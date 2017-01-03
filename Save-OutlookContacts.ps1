Function Get-EWSContacts{
[cmdletbinding()]
Param(
    [string]$EmailAccount,
    [System.Management.Automation.PSCredential]$creds
    )
# This requires the Exchange Web Services Managed API to be installed on the computer where this script is being ran

# Download at - http://www.microsoft.com/en-us/download/confirmation.aspx?id=42022

Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"


#Connetion - https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.exchangeversion%28v=exchg.80%29.aspx

#Change the Exchange Version to work with your environment

$EWS = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)

#Change the “UseDefaultCredentials” to false if you want to specify alternate creds
#$psCred = Get-Credential -Credential totaltool\ttmailop  
$ewscreds = New-Object System.Net.NetworkCredential($creds.UserName.ToString(),$creds.GetNetworkCredential().password.ToString())    
$ews.Credentials = $ewscreds
$EWS.UseDefaultCredentials = $false

#$EWS.AutodiscoverUrl($EmailAccount,{$true})
$uri=[system.URI] "https://mail.totaltool.com/ews/exchange.asmx"    
$ews.Url = $uri 
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAccount)
$ExportCollection = @()

$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$EmailAccount)     
$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ews,$folderid)
#$contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ews,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$EmailAccount)
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(5000)      
$fiItems = $null 
do{      
    $fiItems = $ews.FindItems($Contacts.Id,$ivItemView)
    foreach($Item in $fiItems.Items){       
        if($Item -is [Microsoft.Exchange.WebServices.Data.Contact]){  
            $expObj = "" | select CompanyName,DisplayName,GivenName,Surname,Email1EmailAddress,Categories,BusinessPhone,MobilePhone,BusinessStreet,BusinessCity,BusinessState,BusinessPostalCode
            $expObj.CompanyName = $Item.CompanyName
            $expObj.DisplayName = $Item.DisplayName  
            $expObj.GivenName = $Item.GivenName  
            $expObj.Surname = $Item.Surname
            
            #Get Email Information
            if($Item.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1)){                  
                $expObj.Email1EmailAddress = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address  
            }  
            
            #Get Categorie Infomation
            $expObj.Categories=$Item.Categories
            
            #Get Phone Information
            $BusinessPhone = $null  
            $MobilePhone = $null  
            if($Item.PhoneNumbers -ne $null){  
                if($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone,[ref]$BusinessPhone)){  
                    $expObj.BusinessPhone = $BusinessPhone  
                }  
                if($Item.PhoneNumbers.TryGetValue([Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone,[ref]$MobilePhone)){  
                    $expObj.MobilePhone = $MobilePhone
                }
            }

            #Get Address Information
            $BusinessAddress = $null
            if($item.PhysicalAddresses -ne $null){
                if($item.PhysicalAddresses.TryGetValue([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business,[ref]$BusinessAddress)){  
                    $expObj.BusinessStreet = $BusinessAddress.Street  
                    $expObj.BusinessCity = $BusinessAddress.City  
                    $expObj.BusinessState = $BusinessAddress.State  
                    $expObj.BusinessPostalCode =$BusinessAddress.PostalCode
                }  
            }  

            $ExportCollection += $expObj  
         }
    }
    $ivItemView.Offset += $fiItems.Items.Count
}while($fiItems.MoreAvailable -eq $true) 

$FileName = "S:\ContactExports\" + $EmailAccount + ".csv" 
$ExportCollection | Export-Csv -NoTypeInformation -Path $FileName 
"Exported to " + $FileName
}

Function Split-Categories{
[cmdletbinding()]
Param(
    [string]$filePath
    )
$file = import-csv $filePath
$expandfile = @()
foreach($row in $file){ 
    if(($row.Categories -ne "Personal") -and ($row.companyname -notmatch "Total Tool*")){
        $tempfile = "" | select CompanyName,givenname,surname,Email1EmailAddress,businessphone,mobilephone,Categories,businessstreet,businesscity,businessstate,businesspostalcode
        $tempfile.'companyname' = $row.companyname
        $tempfile.'givenname' = $row.givenname
        $tempfile.'surname' = $row.surname
        $tempfile.'Email1EmailAddress' = $row.email1emailaddress
        $tempfile.'businessphone' = $row.businessphone
        $tempfile.'mobilephone' = $row.mobilephone
        
        $split = $row.categories
        $newCategories =";#"+($split -replace ",",";#")+";#"
        $tempfile.'categories' = $newCategories

        $tempfile.'businessstreet' = $row.businessstreet
        $tempfile.'businesscity' = $row.businesscity
        $tempfile.'businessstate' = $row.businessstate
        $tempfile.'businesspostalcode' = $row.businesspostalcode
    
    $expandfile += $tempfile
    }
}

$expandfile | Export-Csv -NoTypeInformation -Path $filePath 
"Exported to " + $filePath
}



$mailcred = Get-Credential -Credential totaltool\ttmailop 
$EmailAccounts = Import-Csv "C:\Scripts\Contacts Export\SalesReps.csv"
foreach($EmailAccount in $EmailAccounts){
    Get-EWSContacts -EmailAccount $EmailAccount.email -creds $mailcred
}
Get-ChildItem "S:\ContactExports" | 
Foreach-Object {
    Split-Categories -filePath $_.FullName
}