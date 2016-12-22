# Setup the correct modules for SharePoint Manipulation 
if ( (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null ) 
{ 
   Add-PsSnapin Microsoft.SharePoint.PowerShell 
} 
$host.Runspace.ThreadOptions = "ReuseThread"


#Open SharePoint List 
$SPServer="http://sharepoint.totaltool.int/sales"
$SPAppList="/Lists/Contacts" 
$spWeb = Get-SPWeb $SPServer 
$spData = $spWeb.GetList($spWeb.ServerRelativeURL + $SPAppList)


$InvFile="C:\Scripts\1ContactUpload.csv" 
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
   "Adding entry for "+$row."GivenName".ToString() 
   $spItem = $spData.AddItem()
   $spItem["Company Name"] = $row."CompanyName".ToString() 
   $spItem["First Name"] = $row."GivenName".ToString() 
   $spItem["Last Name"] = $row."Surname".ToString() 
   $spItem["Email Address"] = $row."Email1EmailAddress".ToString() 
   $spItem["Business Phone"] = $row."BusinessPhone".ToString() 
   $spItem["Mobile Phone"] = $row."MobilePhone".ToString()
   $spItem["Categories"] = $row."Categories".ToString()
   $spItem["Street Address"] = $row."BusinessStreet".ToString()
   $spItem["City"] = $row."BusinessCity".ToString()
   $spItem["State"] = $row."BusinessState".ToString()
   $spItem["Zip"] = $row."BusinessPostalCode".ToString()
   $spItem.Update() 
}

"—————" 
"Upload Complete"

$spWeb.Dispose()