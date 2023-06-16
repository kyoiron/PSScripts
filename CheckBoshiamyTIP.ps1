New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null
$RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
$RegUninstallPaths += Get-ChildItem HKU: | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object {"HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"}
$Software_installeds =@()
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$Log_Folder_Path = $Log_Path +"\Boshiamy Input Method"
foreach ($Path in $RegUninstallPaths) {
    if (Test-Path $Path) {
        $Software_installeds += Get-ItemProperty $Path | Where-Object{ $_.DisplayName -like "嘸蝦米輸入法*"} 
    }
}
Remove-PSDrive -Name HKU
if($Software_installeds.Count -gt 0){
    if(!(Test-Path $Log_Folder_Path)){ New-Item -ItemType Directory -Path $Log_Folder_Path -Force }
     foreach($item in $Software_installeds){       
        $LogFileName =$env:Computername + "_"+$item.DisplayName + "_" + $item.DisplayVersion  +".txt"
        $item | Out-File -FilePath  "$env:systemdrive\temp\$LogFileName"
        #if(Test-Path "$env:systemdrive\temp\$LogFileName") {robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogFileName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
        if(Test-Path "$env:systemdrive\temp\$LogFileName") { Move-Item  "$env:systemdrive\temp\$LogFileName" -Destination $Log_Folder_Path -Force }
     }
}

<#
    DisplayName     : 嘸蝦米輸入法 J 授權版 (x64)
    DisplayVersion  : rev. 456
    Publisher       : 行易有限公司
    URLInfoAbout    : http://boshiamy.com
    URLUpdateInfo   : http://boshiamy.com/product.html
    HelpLink        : http://boshiamy.com/contact_mail.html
    Contact         : liu@liu.com.tw
    HelpTelephone   : +886 2 23415677
    UninstallString : C:\Program Files\BoshiamyTIP\unliu64.exe
    DisplayIcon     : C:\Program Files\BoshiamyTIP\BoshiamyConfig.exe
    PSPath          : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BoshiamyTIP
    PSParentPath    : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
    PSChildName     : BoshiamyTIP
    PSDrive         : HKLM
    PSProvider      : Microsoft.PowerShell.Core\Registry
#>