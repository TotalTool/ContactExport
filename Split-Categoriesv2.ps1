$file = import-csv C:\OutlookCRMFiles\RawExports\dan.swenson@totaltool.com-ContactsExportnew.csv
$expandfile = @()
foreach($row in $file){
    if($row.Categories -ne "Personal"){
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
#$expandfile | select companyname,displayname,emailaddress,categories | Export-Csv C:\OutlookCRMFiles\CleanedExports\$file + ".csv"
$FileName = "C:\OutlookCRMFiles\CleanedExports\Dan-ContactsExportnew.csv" 
$expandfile | Export-Csv -NoTypeInformation -Path $FileName 
"Exported to " + $FileName