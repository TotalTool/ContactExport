#——————————————————————– 
# Name: Load CSV into SharePoint List 
# NOTE: No warranty is expressed or implied by this code – use it at your 
# own risk. If it doesn’t work or breaks anything you are on your own 
#——————————————————————–


# Setup the correct modules for SharePoint Manipulation 
if ( (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null ) 
{ 
   Add-PsSnapin Microsoft.SharePoint.PowerShell 
} 
$host.Runspace.ThreadOptions = "ReuseThread"


#Open SharePoint List 
$SPServer="http://sharepoint.totaltool.int/itv2"
$SPAppList="/Lists/Test CSV Upload" 
$spWeb = Get-SPWeb $SPServer 
$spData = $spWeb.GetList($spWeb.ServerRelativeURL + $SPAppList)


$InvFile="ContactUpload.csv" 
# Get Data from Inventory CSV File 
$FileExists = (Test-Path $InvFile -PathType Leaf) 
if ($FileExists) { 
   "Loading $InvFile for processing…" 
   $tblData = Import-CSV $InvFile 
} else { 
   "$InvFile not found – stopping import!" 
   exit 
}

# Loop through Applications add each one to SharePoint

"Uploading data to SharePoint…."

foreach ($row in $tblData) 
{ 
   "Adding entry for "+$row."First Name".ToString() 
   $spItem = $spData.AddItem() 
   $spItem["First Name"] = $row."First Name".ToString() 
   $spItem["Last Name"] = $row."Last Name".ToString() 
   $spItem["Email Address"] = $row."Email Address".ToString() 
   $spItem["Business Phone"] = $row."Business Phone".ToString() 
   $spItem["Mobile Phone"] = $row."Mobile Phone".ToString()
   $spItem.Update() 
}

"—————" 
"Upload Complete"

$spWeb.Dispose()