#檢查及安裝JDK
$JDK_EXE_Path = "\\172.29.205.114\loginscript\Update\JDK"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
#$Global:PatternJava32 = "Java ([0-9]|[0-9][0-9]) Update ([0-9]|[0-9][0-9]|[0-9][0-9][0-9])$"
$Global:PatternJDK64 = "Java(TM) SE Development Kit *" 
$RegUninstallPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*")
$JDK_64_EXE = Get-ChildItem -Path ($JDK_EXE_Path+"\"+"*-windows-x64.exe") | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1
if($JDK_64_EXE.FullName){
    $JDK_64_EXE_ProductName = $JDK_64_EXE.VersionInfo.ProductName
    $JDK_64_EXE_ProductVersion = $JDK_64_EXE.VersionInfo.ProductVersion
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $JDK_64_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $Global:PatternJDK64 } 
            <#
                AuthorizedCDFPrefix : 
                Comments            : 
                Contact             : http://java.com
                DisplayVersion      : 16.0.1.0
                HelpLink            : http://java.com/help
                HelpTelephone       : 
                InstallDate         : 20210707
                InstallLocation     : C:\Program Files\Java\jdk-16.0.1\
                InstallSource       : C:\Users\tndadmin\AppData\LocalLow\Oracle\Java\jdk16.0.1_x64\
                ModifyPath          : MsiExec.exe /X{75CDB88B-F917-5456-AB2D-5504DE7F43DE}
                NoModify            : 1
                NoRepair            : 1
                Publisher           : Oracle Corporation
                Readme              : C:\Program Files\Java\jdk-16.0.1\README.html
                Size                : 
                EstimatedSize       : 292577
                UninstallString     : MsiExec.exe /X{75CDB88B-F917-5456-AB2D-5504DE7F43DE}
                URLInfoAbout        : http://java.com
                URLUpdateInfo       : http://www.oracle.com/technetwork/java/javase/downloads
                VersionMajor        : 16
                VersionMinor        : 0
                WindowsInstaller    : 1
                Version             : 268435457
                Language            : 1033
                DisplayName         : Java(TM) SE Development Kit 16.0.1 (64-bit)
                PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{75CDB88B-F917-5456-AB2D-5504DE7F43DE}
                PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
                PSChildName         : {75CDB88B-F917-5456-AB2D-5504DE7F43DE}
                PSDrive             : HKLM
                PSProvider          : Microsoft.PowerShell.Core\Registry           
            #>            
        }
    }
    $JDK_64_Lastest_installed = $JDK_64_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    $LogName = $env:Computername + "_"+$JDK_64_EXE_ProductName+"_"+ $JDK_64_EXE_ProductVersion + ".txt"
    $arguments = "/s REBOOT=0 /LV* ""$env:systemdrive\temp\$LogName"""    
    if($JDK_64_Lastest_installed){
        #有安裝狀況---只有安裝exe檔比已經安裝中最新的還要新再裝
        if([version]$JDK_64_Lastest_installed.DisplayVersion -lt [version]$JDK_64_EXE_ProductVersion){
            robocopy $JDK_EXE_Path "$env:systemdrive\temp" ""$JDK_64_EXE.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
            unblock-file ($env:systemdrive+"\temp\"+ $JDK_64_EXE.Name)
            start-process ($env:systemdrive+"\temp\"+ $JDK_64_EXE.Name) -arg $arguments -wait
            #在確認已安裝中最新的java版本
            foreach ($Path in $RegUninstallPaths) {
                if (Test-Path $Path) {
                    $JDK_64_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $Global:PatternJDK64} 
                    $JDK_64_Lastest_installed  = $JDK_64_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
                }
            }                     
        }                
    }else{
        #無安裝狀況
        robocopy $JDK_EXE_Path "$env:systemdrive\temp" ""$JDK_64_EXE.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
        unblock-file ($env:systemdrive+"\temp\"+ $JDK_64_EXE.Name)
        start-process ($env:systemdrive+"\temp\"+ $JDK_64_EXE.Name) -arg $arguments -wait
        #在確認已安裝中最新的java版本
        foreach ($Path in $RegUninstallPaths) {
            if (Test-Path $Path) {
                $JDK_64_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $Global:PatternJDK64} 
                $JDK_64_Lastest_installed  = $JDK_64_installeds | Sort-Object -Property Version -Descending | Select-Object -first 1
            }
        }     
    }
}
if($JDK_64_Lastest_installed){
    foreach($exe64 in $JDK_64_installeds){
        if([version]$exe64.DisplayVersion -eq [version]$JDK_64_Lastest_installed.DisplayVersion){continue}
        $uninstall = ($exe64.UninstallString  -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()
        $LogFile = $env:systemdrive+"\temp\"+$env:Computername + "_"+ $exe64.DisplayName +".txt"
        start-process "msiexec.exe" -arg "/X $uninstall /quiet /passive /norestart /log ""$LogFile """ -Wait -WindowStyle Hidden        
    }
}
$Log_Folder_Path = $Log_Path +"\"+ "JDK"
if(!(Test-Path -Path $Log_Folder_Path)){ New-Item -ItemType Directory -Path $Log_Folder_Path -Force }
$LogPattern="${env:Computername}_"+"Java(TM) Platform SE"+"*.txt"
if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null }

#設定JDK環境變數
if($env:JAVA_HOME -ne $JDK_64_installeds.InstallLocation.TrimEnd('\')){
    [System.Environment]::SetEnvironmentVariable('JAVA_HOME',$JDK_64_installeds.InstallLocation.TrimEnd('\'),[System.EnvironmentVariableTarget]::Machine)    
}


#檢查att有沒有安裝及路徑
$RegUninstallPaths32 = @("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
 foreach ($Path in $RegUninstallPaths32) {
    if (Test-Path $Path) { 
        $ATT_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like "指紋差勤管理系統標準版" } 
        <#
            Inno Setup: Setup Version         : 5.4.3 (a)
            Inno Setup: App Path              : C:\Program Files (x86)\Att
            InstallLocation                   : C:\Program Files (x86)\Att\
            Inno Setup: Icon Group            : 指紋差勤管理系統標準版
            Inno Setup: User                  : tndadmin
            Inno Setup: Setup Type            : full
            Inno Setup: Selected Components   : admin,zksensor
            Inno Setup: Deselected Components : 
            Inno Setup: Language              : cnt
            DisplayName                       : 指紋差勤管理系統標準版
            UninstallString                   : "C:\Program Files (x86)\Att\unins000.exe"
            QuietUninstallString              : "C:\Program Files (x86)\Att\unins000.exe" /SILENT
            NoModify                          : 1
            NoRepair                          : 1
            InstallDate                       : 20210707
            EstimatedSize                     : 18484
            PSPath                            : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\指紋差勤管理系統標準版_is1
            PSParentPath                      : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
            PSChildName                       : 指紋差勤管理系統標準版_is1
            PSDrive                           : HKLM
            PSProvider                        : Microsoft.PowerShell.Core\Registry           
        #>            
    }
}

#檢查
$Card_folder = "D:\CARD"
$UgClient_folder = $Card_folder + "\" + "UgClient"
$UgClient_Source = "\\172.29.205.114\loginscript\Update\UgClient"
if(!(Test-Path -Path $Card_folder)){New-Item -ItemType Directory -Path $Card_folder -Force}
Robocopy $UgClient_Source $UgClient_folder "/E /XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
#匯入排程
if((Get-ScheduledTaskInfo -TaskName "掌型機刷卡檔上傳法務部" -ErrorAction Ignore) -eq $null){
    $schtasksOutput = schtasks.exe /create /RU "NT AUTHORITY\SYSTEM" /TN "掌型機刷卡檔上傳法務部" /XML "$UgClient_folder\掌型機刷卡檔上傳法務部.xml" /F
}else{
    $schtasksOutput = "排程已建立"    
}
