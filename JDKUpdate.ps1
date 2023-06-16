$JDKs_EXE_Path = "\\172.29.205.114\loginscript\Update\JDK"
$Log_Path = "\\172.29.205.114\Public\sources\audit"

$Global:PatternJDK32 = "Java SE Development Kit ([0-9]|[0-9][0-9]) Update ([0-9]|[0-9][0-9]|[0-9][0-9][0-9])$"
$Global:PatternJDK64 = "Java SE Development Kit ([0-9]|[0-9][0-9]) Update ([0-9]|[0-9][0-9]|[0-9][0-9][0-9]) \(64-bit\)$" 
$RegUninstallPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
$JDKs_EXE_32 = Get-ChildItem -Path ($JDKs_EXE_Path+"\jdk-*-i586.exe") | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1
$JDKs_EXE_64 = Get-ChildItem -Path ($JDKs_EXE_Path+"\jdk-*-x64.exe") | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1
if($JDKs_EXE_32.FullName){    
    $JDKs_EXE_32_ProductName = $JDKs_EXE_32.VersionInfo.ProductName
    $JDKs_EXE_32_ProductVersion = $JDKs_EXE_32.VersionInfo.ProductVersion
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $JDK_32_installeds += Get-ItemProperty $Path | Where-Object{$_.DisplayName -match $Global:PatternJDK32} 
            <#
                AuthorizedCDFPrefix : 
                Comments            : 
                Contact             : https://java.com
                DisplayVersion      : 8.0.3510.10
                HelpLink            : https://java.com/help
                HelpTelephone       : 
                InstallDate         : 20230418
                InstallLocation     : C:\Program Files\Java\jdk1.8.0_351\
                InstallSource       : C:\Users\tndadmin\AppData\LocalLow\Oracle\Java\jdk1.8.0_351_x64\
                ModifyPath          : MsiExec.exe /X{64A3A4F4-B792-11D6-A78A-00B0D0180351}
                NoModify            : 1
                NoRepair            : 1
                Publisher           : Oracle Corporation
                Readme              : C:\Program Files\Java\jdk1.8.0_351\README.html
                Size                : 
                EstimatedSize       : 245785
                UninstallString     : MsiExec.exe /X{64A3A4F4-B792-11D6-A78A-00B0D0180351}
                URLInfoAbout        : https://java.com
                URLUpdateInfo       : https://www.oracle.com/technetwork/java/javase/downloads
                VersionMajor        : 8
                VersionMinor        : 0
                WindowsInstaller    : 1
                Version             : 134221238
                Language            : 1033
                DisplayName         : Java SE Development Kit 8 Update 351 (64-bit)
                PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{64A3A4F4-B792-11D6-A78A-00B0D0180351}
                PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
                PSChildName         : {64A3A4F4-B792-11D6-A78A-00B0D0180351}
                PSDrive             : HKLM
                PSProvider          : Microsoft.PowerShell.Core\Registry
              
            #>            
        }
    }
    $JDK_32_Lastest_installed  = $JDK_32_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
    $LogName = $env:Computername + "_"+$JDKs_EXE_32_ProductName+"_"+ $JDKs_EXE_32_ProductVersion + ".txt"
    $arguments = "/s REBOOT=0 ADDLOCAL=""ToolsFeature,SourceFeature"" /LV* ""$env:systemdrive\temp\$LogName"""    
    if($JDK_32_Lastest_installed){
        #有安裝狀況---只有安裝exe檔比已經安裝中最新的還要新再裝
        if([version]$JDK_32_Lastest_installed.DisplayVersion -lt [version]$JDKs_EXE_32_ProductVersion){
            robocopy $JDKs_EXE_Path "$env:systemdrive\temp" ""$JDKs_EXE_32.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
            unblock-file ($env:systemdrive+"\temp\"+ $JDKs_EXE_32.Name)
            start-process ($env:systemdrive+"\temp\"+ $JDKs_EXE_32.Name) -arg $arguments -wait
            #在確認已安裝中最新的jdk版本
            foreach ($Path in $RegUninstallPaths) {
                if (Test-Path $Path) {
                    $JDK_32_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -match $Global:PatternJDK32} 
                    $JDK_32_Lastest_installed  = $JDK_32_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
                }
            }                     
        }                
    }else{
        #無安裝狀況
        robocopy $JDKs_EXE_Path "$env:systemdrive\temp" ""$JDKs_EXE_32.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
        unblock-file ($env:systemdrive+"\temp\"+ $JDKs_EXE_32.Name)
        start-process ($env:systemdrive+"\temp\"+ $JDKs_EXE_32.Name) -arg $arguments -wait
        #在確認已安裝中最新的jdk版本
        foreach ($Path in $RegUninstallPaths) {
            if (Test-Path $Path) {
                $JDK_32_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -match $Global:PatternJDK32} 
                $JDK_32_Lastest_installed  = $JDK_32_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
            }
        }     
    }
}
if($JDKs_EXE_64.FullName){    
    $JDKs_EXE_64_ProductName = $JDKs_EXE_64.VersionInfo.ProductName
    $JDKs_EXE_64_ProductVersion = $JDKs_EXE_64.VersionInfo.ProductVersion    
    $JDK_64_installeds = Get-ItemProperty $RegUninstallPaths[0] | Where-Object{$_.DisplayName -match $Global:PatternJDK64}          
    $JDK_64_Lastest_installed  = $JDK_64_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
    $LogName = $env:Computername + "_"+$JDKs_EXE_64_ProductName+" (x64)"+"_"+ $JDKs_EXE_64_ProductVersion + ".txt"
    $arguments = "/s  REBOOT=0 ADDLOCAL=""ToolsFeature,SourceFeature"" /LV* ""$env:systemdrive\temp\$LogName"""
    if($JDK_64_Lastest_installed){
        #有安裝狀況---只有安裝exe檔比已經安裝中最新的還要新再裝
        if([version]$JDK_64_Lastest_installed.DisplayVersion -lt [version]$JDKs_EXE_64_ProductVersion){
            robocopy $JDKs_EXE_Path "$env:systemdrive\temp" ""$JDKs_EXE_64.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
            unblock-file ($env:systemdrive+"\temp\"+ $JDKs_EXE_64.Name)
            start-process ($env:systemdrive+"\temp\"+ $JDKs_EXE_64.Name) -arg $arguments -wait
            #在確認已安裝中最新的jdk版本
            $JDK_64_installeds = Get-ItemProperty $RegUninstallPaths[0] | Where-Object{$_.DisplayName -match $Global:PatternJDK64}          
            $JDK_64_Lastest_installed  = $JDK_64_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1 
        }
    }else{
        #無安裝狀況
            robocopy $JDKs_EXE_Path "$env:systemdrive\temp" ""$JDKs_EXE_64.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
            unblock-file ($env:systemdrive+"\temp\"+ $JDKs_EXE_64.Name)
            start-process ($env:systemdrive+"\temp\"+ $JDKs_EXE_64.Name) -arg $arguments -wait
            #在確認已安裝中最新的jdk版本
            $JDK_64_installeds = Get-ItemProperty $RegUninstallPaths[0] | Where-Object{$_.DisplayName -match $Global:PatternJDK64}          
            $JDK_64_Lastest_installed  = $JDK_64_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
    }
}
<#
    Get-ChildItem exe
    PSPath            : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\JDK\jdk-8u351-windows-x64.exe
    PSParentPath      : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\JDK
    PSChildName       : jdk-8u351-windows-x64.exe
    PSProvider        : Microsoft.PowerShell.Core\FileSystem
    PSIsContainer     : False
    Mode              : -a----
    VersionInfo       : File:             \\172.29.205.114\loginscript\Update\JDK\jdk-8u351-windows-x64.exe
                        InternalName:     Setup Launcher
                        OriginalFilename: wrapper_jdk.exe
                        FileVersion:      8.0.3510.10
                        FileDescription:  Java Platform SE binary
                        Product:          Java Platform SE 8 U351
                        ProductVersion:   8.0.3510.10
                        Debug:            False
                        Patched:          False
                        PreRelease:       False
                        PrivateBuild:     False
                        SpecialBuild:     False
                        Language:         英文 (美國)
                        
    BaseName          : jdk-8u351-windows-x64
    Target            : 
    LinkType          : 
    Name              : jdk-8u351-windows-x64.exe
    Length            : 184064648
    DirectoryName     : \\172.29.205.114\loginscript\Update\JDK
    Directory         : \\172.29.205.114\loginscript\Update\JDK
    IsReadOnly        : False
    Exists            : True
    FullName          : \\172.29.205.114\loginscript\Update\JDK\jdk-8u351-windows-x64.exe
    Extension         : .exe
    CreationTime      : 2023/4/17 上午 11:43:06
    CreationTimeUtc   : 2023/4/17 上午 03:43:06
    LastAccessTime    : 2023/4/17 下午 03:31:04
    LastAccessTimeUtc : 2023/4/17 上午 07:31:04
    LastWriteTime     : 2023/4/17 上午 11:43:32
    LastWriteTimeUtc  : 2023/4/17 上午 03:43:32
    Attributes        : Archive
    
#>
#移除舊版JDK
if($JDK_32_Lastest_installed){
    foreach($exe32 in $JDK_32_installeds){
        if([version]$exe32.DisplayVersion -eq [version]$JDK_32_Lastest_installed.DisplayVersion){continue}
        $uninstall = ($exe32.UninstallString  -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()
        $LogFile = $env:systemdrive+"\temp\" + $env:Computername + "_"+ $exe32.DisplayName+"_"+ $JDKs_EXE_32_ProductVersion + "_Remove.txt"
        start-process "msiexec.exe" -arg "/X $uninstall /quiet /passive /norestart /log ""$LogFile""" -Wait -WindowStyle Hidden
    }
}
if($JDK_64_Lastest_installed){
    foreach($exe64 in $JDK_64_installeds){
        if([version]$exe64.DisplayVersion -eq [version]$JDK_64_Lastest_installed.DisplayVersion){continue}
        $uninstall = ($exe64.UninstallString  -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()
        $LogFile = $env:systemdrive+"\temp\"+$env:Computername + "_"+ $exe64.DisplayName+"_"+ $JDKs_EXE_64_ProductVersion + "_Remove.txt"
        start-process "msiexec.exe" -arg "/X $uninstall /quiet /passive /norestart /log ""$LogFile """ -Wait -WindowStyle Hidden
    }
}
$Log_Folder_Path = $Log_Path +"\"+ "JDK"
if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
$LogPattern="${env:Computername}_"+"Java"+"*SE*.txt"
if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null }