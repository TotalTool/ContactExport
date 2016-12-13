$Path = C:\OutlookCRMFiles\RawExports
$files = Get-ChildItem -Path $Path
ForEach ($file in $files) { 
    Split-Categories $file
    }