$RegUninstallPaths = @('HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
$Jail_installeds = @()
foreach ($Path in $RegUninstallPaths) {
    if (Test-Path $Path) {
        $Jail_installeds += @(Get-ItemProperty $Path | Where-Object{$_.DisplayName -like "獄政資訊系統"})
    }
}
<#
    ShortcutAppId           : http://nlbap.newjail/CjalPublish/CJALWpf.application#CJALWpf.application, Culture=neutral, PublicKeyToken=0000000000000000, processorArchitecture=msil
    SupportShortcutFileName : 獄政資訊系統 線上支援
    ShortcutSuiteName       : CJAL
    ShortcutFileName        : 獄政資訊系統
    ShortcutFolderName      : 法務部矯正署
    UrlUpdateInfo           : http://nlbap.newjail/CjalPublish/CJALWpf.application
    UninstallString         : rundll32.exe dfshim.dll,ShArpMaintain CJALWpf.application, Culture=neutral, PublicKeyToken=0000000000000000, processorArchitecture=msil
    Publisher               : 法務部矯正署
    DisplayVersion          : 1.1.210.3
    DisplayIcon             : dfshim.dll,2
    DisplayName             : 獄政資訊系統
    PSPath                  : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\6962f8d33a299b0b
    PSParentPath            : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
    PSChildName             : 6962f8d33a299b0b
    PSDrive                 : HKCU
    PSProvider              : Microsoft.PowerShell.Core\Registry
#>

 if($Jail_installeds){    
    $wshell = new-object -com wscript.shell
    $selectedUninstallString = $Jail_installeds.UninstallString
    $wshell.run("cmd /c $selectedUninstallString")
    Start-Sleep 5
    #$wshell.sendkeys("`"OK`"~")
    $wshell.sendkeys("`"OK`"~")
    if(Test-Path  "$env:LOCALAPPDATA\Apps\2.0"){Remove-Item "$env:LOCALAPPDATA\Apps\2.0"  -Recurse -Force}
 }

