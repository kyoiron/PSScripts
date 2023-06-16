New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null
$RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
$RegUninstallPaths += Get-ChildItem HKU: | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object {"HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"}
$Software_installeds =@()
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$Log_Folder_Path = $Log_Path +"\LINE"
foreach ($Path in $RegUninstallPaths) {
    if (Test-Path $Path) {
        $Software_installeds += Get-ItemProperty $Path | Where-Object{ $_.DisplayName -like "LINE*"} 
    }
}
Remove-PSDrive -Name HKU
if($Software_installeds.Count -gt 0){
    if(!(Test-Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
     foreach($item in $Software_installeds){       
        $LogFileName =$env:Computername + "_"+$item.DisplayName + "_" + $item.DisplayVersion  +".txt"
        $item | Out-File -FilePath  "$env:systemdrive\temp\$LogFileName"        
        if(Test-Path "$env:systemdrive\temp\$LogFileName") { Move-Item  "$env:systemdrive\temp\$LogFileName" -Destination $Log_Folder_Path -Force }
     }         
}

<#
    DisplayName     : LINE
    UninstallString : C:\Users\kyoiron\AppData\Local\LINE\bin\LineUnInst.exe
    DisplayVersion  : 7.16.1.3000
    URLInfoAbout    : http://line.me
    Publisher       : LINE Corporation
    DisplayIcon     : C:\Users\kyoiron\AppData\Local\LINE\bin\LineLauncher.exe
    PSPath          : Microsoft.PowerShell.Core\Registry::HKEY_USERS\S-1-5-21-348651901-4118955271-3989848957-18762\Software\Microsoft\Windows\CurrentVersion\Uninstall\LINE
    PSParentPath    : Microsoft.PowerShell.Core\Registry::HKEY_USERS\S-1-5-21-348651901-4118955271-3989848957-18762\Software\Microsoft\Windows\CurrentVersion\Uninstall
    PSChildName     : LINE
    PSDrive         : HKU
    PSProvider      : Microsoft.PowerShell.Core\Registry
#>