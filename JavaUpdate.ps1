#java不更新清單
$JavaExcludePC=@("TND-GASE-061")
#$JavaExcludePC=@()

if($JavaExcludePC.Contains($env:Computername)){exit}
$Javas_EXE_Path = "\\172.29.205.114\loginscript\Update\Java"
$Log_Path = "\\172.29.205.114\Public\sources\audit"

$Global:PatternJava32 = "Java ([0-9]|[0-9][0-9]) Update ([0-9]|[0-9][0-9]|[0-9][0-9][0-9])$"
$Global:PatternJava64 = "Java ([0-9]|[0-9][0-9]) Update ([0-9]|[0-9][0-9]|[0-9][0-9][0-9]) \(64-bit\)$" 
$RegUninstallPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
$Javas_EXE_32 = Get-ChildItem -Path ($Javas_EXE_Path+"\*-i586.exe") | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1
$Javas_EXE_64 = Get-ChildItem -Path ($Javas_EXE_Path+"\*-x64.exe") | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1
$msiexecProcessName = "msiexec.exe"
#自造字
$FontClient_EXE_Path = "$env:systemdrive\CMEX_FontClient\FontClient.exe"
$FontClient_AutoUpdate_EXE_Path = "$env:systemdrive\CMEX_FontClient\AutoUpdate.exe"

if($Javas_EXE_32.FullName){    
    $Javas_EXE_32_ProductName = $Javas_EXE_32.VersionInfo.ProductName
    $Javas_EXE_32_ProductVersion = $Javas_EXE_32.VersionInfo.ProductVersion
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $Java_32_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -match $Global:PatternJava32} 
            <#
                AuthorizedCDFPrefix : 
                Comments            : 
                Contact             : https://java.com
                DisplayVersion      : 8.0.2610.12
                HelpLink            : https://java.com/help
                HelpTelephone       : 
                InstallDate         : 20200716
                InstallLocation     : C:\Program Files (x86)\Java\jre1.8.0_261\
                InstallSource       : C:\temp\
                ModifyPath          : MsiExec.exe /X{26A24AE4-039D-4CA4-87B4-2F32180261F0}
                NoModify            : 1
                NoRepair            : 1
                Publisher           : Oracle Corporation
                Readme              : [INSTALLDIR]README.txt
                Size                : 
                EstimatedSize       : 221468
                UninstallString     : MsiExec.exe /X{26A24AE4-039D-4CA4-87B4-2F32180261F0}
                URLInfoAbout        : https://java.com
                URLUpdateInfo       : https://java.sun.com
                VersionMajor        : 8
                VersionMinor        : 0
                WindowsInstaller    : 1
                Version             : 134220338
                Language            : 1033
                DisplayName         : Java 8 Update 261
                sEstimatedSize2     : 110734
                PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{26A24AE4-039D-4CA4-87B4-2F32180261F0}
                PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
                PSChildName         : {26A24AE4-039D-4CA4-87B4-2F32180261F0}
                PSDrive             : HKLM
                PSProvider          : Microsoft.PowerShell.Core\Registry                
            #>            
        }
    }
    $Java_32_Lastest_installed  = $Java_32_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
    $LogName = $env:Computername + "_"+$Javas_EXE_32_ProductName+"_"+ $Javas_EXE_32_ProductVersion + ".txt"
    $arguments = "/s AUTO_UPDATE=0 REBOOT=0 /LV* ""$env:systemdrive\temp\$LogName"""    
    if($Java_32_Lastest_installed){
        #有安裝狀況---只有安裝exe檔比已經安裝中最新的還要新再裝
        if([version]$Java_32_Lastest_installed.DisplayVersion -lt [version]$Javas_EXE_32_ProductVersion){
            robocopy $Javas_EXE_Path "$env:systemdrive\temp" ""$Javas_EXE_32.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
            unblock-file ($env:systemdrive+"\temp\"+ $Javas_EXE_32.Name)
            $FontClient = Get-Process -Name "FontClient" -ErrorAction SilentlyContinue
            if(!$FontClient.HasExited){
                Stop-Process -Name "FontClient" -Force -ErrorAction SilentlyContinue
                Wait-Process -Name "FontClient" -ErrorAction SilentlyContinue
            }
            start-process ($env:systemdrive+"\temp\"+ $Javas_EXE_32.Name) -arg $arguments -Wait
            #在確認已安裝中最新的java版本
            foreach ($Path in $RegUninstallPaths) {
                if (Test-Path $Path) {
                    $Java_32_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -match $Global:PatternJava32} 
                    $Java_32_Lastest_installed  = $Java_32_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
                }
            }                     
        }                
    }else{
        #無安裝狀況
        robocopy $Javas_EXE_Path "$env:systemdrive\temp" ""$Javas_EXE_32.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
        unblock-file ($env:systemdrive+"\temp\"+ $Javas_EXE_32.Name)
        $FontClient = Get-Process -Name "FontClient" -ErrorAction SilentlyContinue
        if(!$FontClient.HasExited){
            Stop-Process -Name "FontClient" -Force
            Wait-Process -Name "FontClient" -ErrorAction SilentlyContinue
        }
        start-process ($env:systemdrive+"\temp\"+ $Javas_EXE_32.Name) -arg $arguments -wait
        #在確認已安裝中最新的java版本
        foreach ($Path in $RegUninstallPaths) {
            if (Test-Path $Path) {
                $Java_32_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -match $Global:PatternJava32} 
                $Java_32_Lastest_installed  = $Java_32_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
            }
        }     
    }
}
if($Javas_EXE_64.FullName){
    
    $Javas_EXE_64_ProductName = $Javas_EXE_64.VersionInfo.ProductName
    $Javas_EXE_64_ProductVersion = $Javas_EXE_64.VersionInfo.ProductVersion    
    $Java_64_installeds = Get-ItemProperty $RegUninstallPaths[0] | Where-Object{$_.DisplayName -match $Global:PatternJava64}          
    $Java_64_Lastest_installed  = $Java_64_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
    $LogName = $env:Computername + "_"+$Javas_EXE_64_ProductName+" (x64)"+"_"+ $Javas_EXE_64_ProductVersion + ".txt"
    $arguments = "/s AUTO_UPDATE=0 REBOOT=0 /l*vx ""$env:systemdrive\temp\$LogName"""
    if($Java_64_Lastest_installed){
        #有安裝狀況---只有安裝exe檔比已經安裝中最新的還要新再裝
        if([version]$Java_64_Lastest_installed.DisplayVersion -lt [version]$Javas_EXE_64_ProductVersion){
            robocopy $Javas_EXE_Path "$env:systemdrive\temp" ""$Javas_EXE_64.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
            unblock-file ($env:systemdrive+"\temp\"+ $Javas_EXE_64.Name)
            $FontClient = Get-Process -Name "FontClient" -ErrorAction SilentlyContinue
            if(!$FontClient.HasExited){
                Stop-Process -Name "FontClient" -Force -ErrorAction SilentlyContinue
                Wait-Process -Name "FontClient" -ErrorAction SilentlyContinue
            }
            start-process ($env:systemdrive+"\temp\"+ $Javas_EXE_64.Name) -arg $arguments -wait
            do {
                $Javas_EXE_64_Process = Get-Process -Name $Javas_EXE_64.Name -ErrorAction SilentlyContinue
                if ($Javas_EXE_64_Process -ne $null) {
                    Start-Sleep -Seconds 1
                }
            } while ($Javas_EXE_64_Process -ne $null)
            #在確認已安裝中最新的java版本
            $Java_64_installeds = Get-ItemProperty $RegUninstallPaths[0] | Where-Object{$_.DisplayName -match $Global:PatternJava64}          
            $Java_64_Lastest_installed  = $Java_64_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1 
        }
    }else{
        #無安裝狀況
            robocopy $Javas_EXE_Path "$env:systemdrive\temp" ""$Javas_EXE_64.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
            unblock-file ($env:systemdrive+"\temp\"+ $Javas_EXE_64.Name)
            $FontClient = Get-Process -Name "FontClient" -ErrorAction SilentlyContinue
            if(!$FontClient.HasExited){
                Stop-Process -Name "FontClient" -Force -ErrorAction SilentlyContinue
                Wait-Process -Name "FontClient" -ErrorAction SilentlyContinue
            }
            start-process ($env:systemdrive+"\temp\"+ $Javas_EXE_64.Name) -arg $arguments -wait
            #在確認已安裝中最新的java版本
            $Java_64_installeds = Get-ItemProperty $RegUninstallPaths[0] | Where-Object{$_.DisplayName -match $Global:PatternJava64}          
            $Java_64_Lastest_installed  = $Java_64_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
    }
}
<#
    Get-ChildItem exe
    PSPath            : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\Java\jre-8u261-windows-i586.exe
    PSParentPath      : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\Java
    PSChildName       : jre-8u261-windows-i586.exe
    PSProvider        : Microsoft.PowerShell.Core\FileSystem
    PSIsContainer     : False
    Mode              : -a----
    VersionInfo       : File:             \\172.29.205.114\loginscript\Update\Java\jre-8u261-windows-i586.exe
                        InternalName:     Setup Launcher
                        OriginalFilename: wrapper_jre_offline.exe
                        FileVersion:      8.0.2610.12
                        FileDescription:  Java Platform SE binary
                        Product:          Java Platform SE 8 U261
                        ProductVersion:   8.0.2610.12
                        Debug:            False
                        Patched:          False
                        PreRelease:       False
                        PrivateBuild:     False
                        SpecialBuild:     False
                        Language:         英文 (美國)
                        
    BaseName          : jre-8u261-windows-i586
    Target            : 
    LinkType          : 
    Name              : jre-8u261-windows-i586.exe
    Length            : 72990856
    DirectoryName     : \\172.29.205.114\loginscript\Update\Java
    Directory         : \\172.29.205.114\loginscript\Update\Java
    IsReadOnly        : False
    Exists            : True
    FullName          : \\172.29.205.114\loginscript\Update\Java\jre-8u261-windows-i586.exe
    Extension         : .exe
    CreationTime      : 2020/8/31 下午 05:22:29
    CreationTimeUtc   : 2020/8/31 上午 09:22:29
    LastAccessTime    : 2020/8/31 下午 05:25:11
    LastAccessTimeUtc : 2020/8/31 上午 09:25:11
    LastWriteTime     : 2020/8/31 下午 05:23:09
    LastWriteTimeUtc  : 2020/8/31 上午 09:23:09
    Attributes        : Archive
#>
#移除舊版Java
if($Java_32_Lastest_installed){
    foreach($exe32 in $Java_32_installeds){
        if([version]$exe32.DisplayVersion -eq [version]$Java_32_Lastest_installed.DisplayVersion){continue}
        $uninstall = ($exe32.UninstallString  -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()
        $LogFile = $env:systemdrive+"\temp\" + $env:Computername + "_"+ $exe32.DisplayName+"_"+ $Javas_EXE_32_ProductVersion + "_Remove.txt"
        $FontClient = Get-Process -Name "FontClient" -ErrorAction SilentlyContinue
        if(!$FontClient.HasExited){
            Stop-Process -Name "FontClient" -Force -ErrorAction SilentlyContinue
            Wait-Process -Name "FontClient" -ErrorAction SilentlyContinue
        }
        start-process "msiexec.exe" -arg "/X $uninstall /quiet /l*vx ""$LogFile""" -Wait            
    }
}
if($Java_64_Lastest_installed){
    foreach($exe64 in $Java_64_installeds){
        if([version]$exe64.DisplayVersion -eq [version]$Java_64_Lastest_installed.DisplayVersion){continue}
        $uninstall = ($exe64.UninstallString  -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()
        $LogFile = $env:systemdrive+"\temp\"+$env:Computername + "_"+ $exe64.DisplayName+"_"+ $Javas_EXE_64_ProductVersion + "_Remove.txt"
        $FontClient = Get-Process -Name "FontClient" -ErrorAction SilentlyContinue
        if(!$FontClient.HasExited){
            Stop-Process -Name "FontClient" -Force
            Wait-Process -Name "FontClient" -ErrorAction SilentlyContinue
        }
        start-process "msiexec.exe" -arg "/X $uninstall /quiet /l*vx ""$LogFile""" -Wait      
    }
}
$Log_Folder_Path = $Log_Path +"\"+ "Java"
if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
$LogPattern="${env:Computername}_"+"Java"+"*.txt"
if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern " /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null }

if((Get-Process | Where-Object {$_.MainModule.FileName -eq $FontClient_EXE_Path}) -eq 0){ 
    Start-Process -FilePath $FontClient_EXE_Path -ArgumentList "-gui"
    Start-Process -FilePath $FontClient_AutoUpdate_EXE_Path
}
          