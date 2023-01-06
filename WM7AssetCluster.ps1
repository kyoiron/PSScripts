$WM7Asset_Path = "$env:SystemDrive\WM7Asset"
$isInstall =!((Test-Path $WM7Asset_Path\WM7Asset.bat) -and (Test-Path $WM7Asset_Path\WM7Assetreport.xml) -and (Test-Path $WM7Asset_Path\WM7LiteGreen.exe))
if($isInstall){
    if(!(Test-Path $WM7Asset_Path)){New-Item -Path $WM7Asset_Path -ItemType Directory}
    $url_exe  = "http://download.moj/files/VANS/INTRA/WM7AssetCluster.exe"
    Start-Job -Name WebReq -ScriptBlock { param($p1, $p2)
        Invoke-WebRequest -Uri $p1 -OutFile $p2
    } -ArgumentList $url_exe,"$WM7Asset_Path\WM7AssetCluster.exe"
    Wait-Job -Name WebReq -Force
    Remove-Job -Name WebReq -Force
    Start-Process -FilePath "$WM7Asset_Path\WM7AssetCluster.exe"  -Wait
    Remove-Item "$WM7Asset_Path\WM7AssetCluster.exe" -Force
}