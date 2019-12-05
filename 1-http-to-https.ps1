# url rewrite installation
if (Test-Path "C:\Windows\System32\inetsrv\rewrite.dll" -PathType Any) {
    Write-Host "IIS URL Rewrite Module 2 is installed." -ForegroundColor Green
    Write-Host "Skipping." -ForegroundColor Green
} 

else { 
    Write-Host "IIS URL Rewrite Module 2 is NOT installed." -ForegroundColor Red
    Write-Host "IIS URL Rewrite Module 2 WILL BE installed." -ForegroundColor Green
    Write-Host "Continue? y/n" -ForegroundColor Green
    $askInstall = Read-Host " "
    if ($askInstall -eq "yes" -or $askInstall -eq "y") {
        Write-Host "Installing..." -ForegroundColor Green
        msiexec.exe /i "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi" /qb
    }
    else {
        Return
    }
}
# searching all iis sites
$site = Get-ChildItem -Path IIS:\Sites
$allsites = Get-ChildItem -Path IIS:\Sites | measure
$countsites = $allsites.Count
$x = 0
write-host 
do {
    $site[$x].name
    Set-Variable -name "id" -value $x
    $x = $x + 1
    Get-Variable -name "id" | format-table -HideTableHeaders -autosize
}
While ($x -lt $countsites)
Write-Host "Enter the site ID:" -ForegroundColor Green
$x = Read-Host " "
$sitetorewrite = $site[$x].name
Write-host "You chose the site" $sitetorewrite

# creating http to https rewrite role in iis
$sitetorewrite = 'IIS:\Sites\' + $sitetorewrite
$name = 'http to https redirect role'
$inbound = '(.*svc)'
$outbound = 'https://{HTTP_HOST}{REQUEST_URI}'
$range = 'off'
$root = 'system.webServer/rewrite/rules'
$filter = "{0}/rule[@name='{1}']" -f $root, $name
Add-WebConfigurationProperty -PSPath $sitetorewrite -filter $root -name '.' -value @{name=$name; patterSyntax='Regular Expressions'}
Set-WebConfigurationProperty -PSPath $sitetorewrite -filter "$filter/match" -name 'url' -value $inbound
Set-WebConfigurationProperty -PSPath $sitetorewrite -filter "$filter/match" -name 'negate' -value $true
Set-WebConfigurationProperty -PSPath $sitetorewrite -filter "$filter/conditions" -name '.' -value @{input='{HTTPS}'; matchType='0'; pattern=$range; ignoreCase='True'; negate='False'}
Set-WebConfigurationProperty -PSPath $sitetorewrite -filter "$filter/action" -name 'type' -value 'Redirect'
Set-WebConfigurationProperty -PSPath $sitetorewrite -filter "$filter/action" -name 'redirectType' -value 'SeeOther'
Set-WebConfigurationProperty -PSPath $sitetorewrite -filter "$filter/action" -name 'url' -value $outbound
