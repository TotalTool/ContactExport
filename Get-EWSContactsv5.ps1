# This requires the Exchange Web Services Managed API to be installed on the computer where this script is being ran

# Download at - http://www.microsoft.com/en-us/download/confirmation.aspx?id=42022

Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"


#Connetion - https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.exchangeversion%28v=exchg.80%29.aspx

#$EmailAccount = "dan.swenson@totaltool.com"

#Change the Exchange Version to work with your environment

$EWS = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)

#Change the “UseDefaultCredentials” to false if you want to specify alternate creds
$psCred = Get-Credential    
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())    
$ews.Credentials = $creds
$EWS.UseDefaultCredentials = $false

#$EWS.AutodiscoverUrl($EmailAccount,{$true})
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAccount)
$ExportCollection = @()

$EmailAccounts = Import-Csv 'C:\Scripts\Outlook CRM\SalesReps.csv'
foreach($EmailAccount in $EmailAccounts){
$EWS.AutodiscoverUrl($EmailAccount,{$true})
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$EmailAccount)     
$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ews,$folderid)
#$contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ews,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$EmailAccount)
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)      
$fiItems = $null 
do{      
    $fiItems = $ews.FindItems($Contacts.Id,$ivItemView)
    foreach($Item in $fiItems.Items){       
        if($Item -is [Microsoft.Exchange.WebServices.Data.Contact]){  
            $expObj = "" | select DisplayName,GivenName,Surname,Email1EmailAddress,Categories,BusinessPhone,MobilePhone,BusinessStreet,BusinessCity,BusinessState,BusinessPostalCode
            
            #Get Name Information
            $expObj.DisplayName = $Item.DisplayName  
            $expObj.GivenName = $Item.GivenName  
            $expObj.Surname = $Item.Surname
            
            #Get Email Information
            if($Item.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1)){                  
                #$expObj.Email1DisplayName = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name  
                #$expObj.Email1Type = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].RoutingType  
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
}
$FileName = "c:\temp\" + $EmailAccount + "-ContactsExportnew.csv" 
$ExportCollection | Export-Csv -NoTypeInformation -Path $FileName 
"Exported to " + $FileName