
Import-Module Get-EWSContacts
$mailcred = Get-Credential -Credential totaltool\ttmailop 
$EmailAccounts = Import-Csv 'C:\Scripts\Outlook CRM\SalesReps.csv'
foreach($EmailAccount in $EmailAccounts){
    Get-EWSContacts -EmailAccount $EmailAccount.email -creds $mailcred
}