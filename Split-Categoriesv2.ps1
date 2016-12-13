$file = import-csv C:\OutlookCRMFiles\RawExports\dan.swenson@totaltool.com-ContactsExportnew.csv
$expandfile = @()
foreach($row in $file){
    $tempfile = New-Object psobject -Property @{
        $tempfile['companyname'] = '$row.companyname'
        $tempfile['displayname'] = '$row.displayname'
        $tempfile['emailaddress'] = '$row.email1emailaddress'
        $split = $row.categories.split(",")
        for($i=0;$i -le $split.count; $i++){
            $tempfile['Cat$($i+1)'] = '$split[$i]'
            }
        
    }
    $expandfile += $tempfile
}
$expandfile | select companyname,displayname,emailaddress,cat1,cat2,cat3,cat4 | Export-Csv C:\OutlookCRMFiles\CleanedExports\$file + '.csv'
