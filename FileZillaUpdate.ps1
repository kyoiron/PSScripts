
#exe檔安裝路徑
$FileZillas_Path = "\\172.29.205.114\loginscript\Update\FileZilla"
#LOG檔NAS存放路徑
#$Log_Path = "\\172.29.205.114\Public\sources\audit"
#沒裝的是否要裝$true或$false
$Install_IF_NOT_Installed = $false

$FileZilla_EXE = (Get-ChildItem -Path ($FileZillas_Path+"\*.exe") | Where-Object{$_.VersionInfo.FileDescription -eq "FileZilla FTP Client"} | Sort-Object )  | Sort-Object -Property VersionInfo.FileVersion -Descending | Select-Object -last 1

$FileZilla_EXE_Path = $FileZilla_EXE.FullName

if($FileZilla_EXE_Path){
    $FileZilla_EXE_ProductName = (Get-ItemProperty $FileZilla_EXE_Path).VersionInfo.ProductName
    $FileZilla_EXE_ProductVersion = (Get-ItemProperty $FileZilla_EXE_Path).VersionInfo.ProductVersion
    <#
        PSPath            : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\FileZilla\FileZilla_3.62.2_win64-setup.exe
        PSParentPath      : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\FileZilla
        PSChildName       : FileZilla_3.62.2_win64-setup.exe
        PSProvider        : Microsoft.PowerShell.Core\FileSystem
        PSIsContainer     : False
        Mode              : -a----
        VersionInfo       : File:             \\172.29.205.114\loginscript\Update\FileZilla\FileZilla_3.62.2_win64-setup.exe
                            InternalName:     
                            OriginalFilename: FileZilla_3.62.2_win32-setup.exe
                            FileVersion:      3.62.2
                            FileDescription:  FileZilla FTP Client
                            Product:          FileZilla
                            ProductVersion:   3.62.2
                            Debug:            False
                            Patched:          False
                            PreRelease:       False
                            PrivateBuild:     False
                            SpecialBuild:     False
                            Language:         英文 (美國)
                            
        BaseName          : FileZilla_3.62.2_win64-setup
        Target            : 
        LinkType          : 
        Name              : FileZilla_3.62.2_win64-setup.exe
        Length            : 11905648
        DirectoryName     : \\172.29.205.114\loginscript\Update\FileZilla
        Directory         : \\172.29.205.114\loginscript\Update\FileZilla
        IsReadOnly        : False
        Exists            : True
        FullName          : \\172.29.205.114\loginscript\Update\FileZilla\FileZilla_3.62.2_win64-setup.exe
        Extension         : .exe
        CreationTime      : 2023/1/7 上午 11:18:48
        CreationTimeUtc   : 2023/1/7 上午 03:18:48
        LastAccessTime    : 2023/1/9 上午 09:49:53
        LastAccessTimeUtc : 2023/1/9 上午 01:49:53
        LastWriteTime     : 2023/1/7 上午 11:18:53
        LastWriteTimeUtc  : 2023/1/7 上午 03:18:53
        Attributes        : Archive
    #>
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $FileZilla_installeds = Get-ItemProperty $Path | Where-Object{$_.PSChildName -match ($FileZilla_EXE_ProductName + " Client")} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
        }
    }
    <#
        UninstallString : "C:\Program Files\FileZilla FTP Client\uninstall.exe"
        InstallLocation : C:\Program Files\FileZilla FTP Client
        DisplayName     : FileZilla 3.62.2
        DisplayIcon     : C:\Program Files\FileZilla FTP Client\FileZilla.exe
        DisplayVersion  : 3.62.2
        URLInfoAbout    : https://filezilla-project.org/
        URLUpdateInfo   : https://filezilla-project.org/
        HelpLink        : https://filezilla-project.org/
        Publisher       : Tim Kosse
        VersionMajor    : 3
        VersionMinor    : 62
        NoModify        : 1
        NoRepair        : 1
        EstimatedSize   : 42322
        PSPath          : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\FileZilla Client
        PSParentPath    : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName     : FileZilla Client
        PSDrive         : HKLM
        PSProvider      : Microsoft.PowerShell.Core\Registry
    #>
    $installed = $FileZilla_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    if(($installed -eq $null) -and ($Install_IF_NOT_Installed -ne $true)){exit}
    if(([version]$FileZilla_EXE_ProductVersion -le [version]$installed.DisplayVersion)){exit}
    $LogName = $env:Computername + "_"+$FileZilla_EXE_ProductName +"_"+ $FileZilla_EXE_ProductVersion + ".txt"
    $EXE_FIleName = $FileZilla_EXE.Name
    $arguments = "$env:systemdrive\temp\$EXE_FIleName /S user=all"
    robocopy $FileZillas_Path "$env:systemdrive\temp" $EXE_FIleName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    unblock-file ($env:systemdrive+"\temp\"+$EXE_FIleName)
    start-process ($env:systemdrive+"\temp\"+$EXE_FIleName) -arg $arguments -WindowStyle Hidden 
    #安裝指令無log參數故無回傳LOG檔
}
