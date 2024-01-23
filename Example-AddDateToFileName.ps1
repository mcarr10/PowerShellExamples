$currentdate = (get-date).ToString("yyyyMMdd")
$CurrentPath= Get-Location
New-Item $CurrentPath'_'$currentdate'.txt' -type file