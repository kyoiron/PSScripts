#將tnduser加入本機管理者帳號
Add-LocalGroupMember -Group "Administrators" -Member "$env:COMPUTERNAME\tnduser" -ErrorAction SilentlyContinue
#將tnduser加入本機管理者帳號
net user "tndadmin" /active:yes
Get-LocalUser -Name tnduser | Select-Object * | Out-File $env:SystemDrive\temp\${env:computername}_tnduserStatus.txt
if(test-path("$env:SystemDrive\temp\${env:computername}_tnduserStatus.txt")){Copy-Item "$env:SystemDrive\temp\${env:computername}_tnduserStatus.txt" -Destination  "\\172.29.205.114\Public\sources\audit\tnduser" -Force}
#解鎖Bitlocker槽
$BitlockerRecoveryKey='102157-408870-455730-463155-149358-296956-700711-045573'
if ($env:BitlockerDataDrive -ne $null){
    manage-bde -unlock $env:BitlockerDataDrive -RecoveryPassword $BitlockerRecoveryKey
    #manage-bde -off $env:BitlockerDataDrive
}

#防毒更新病毒碼及政策
    cmd /c "start smc -updateconfig"
#更改tnduser密碼
<#
if($env:COMPUTERNAME -eq "TND-5EES-068" ){
    $Password = "me@TND1234"
    $UserAccount = Get-LocalUser -Name "tnduser"
    $UserAccount | Set-LocalUser -Password (ConvertTo-SecureString -AsPlainText "$Password" -Force)
}
#>

#筆硯新版元件
    # powershell  "$env:SystemDrive\temp\eic.ps1"
#修復筆硯列印工具[機關別]選項列皆為空白問題
<#$Fix_ORG_Disappear_PC=@('TND-BUSE-072')
if($Fix_ORG_Disappear_PC.Contains($env:computername)){
    powershell "$env:SystemDrive\temp\EicPrint_Fix_ORG_Disappear.ps1"
}
#>
<#         
$Sign_officer_Computers=@("TND-HEAD-150","TND-DEPUTY-014","TND-SEOF-062","TND-STOF-113")
if($Sign_officer_Computers.Contains($env:computername)){
    $ShortcutPath_EICSignTSR = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\筆硯簽章工具.lnk"
    $ShortcutPath_EicPrint  = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\筆硯列印工具.lnk"
    if(!(Test-Path $ShortcutPath_EICSignTSR)){
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($ShortcutPath_EICSignTSR)
        $Shortcut.TargetPath = "C:\eic\EICSignTSR\EicSignTSR.exe"
        $Shortcut.IconLocation = "C:\eic\EICSignTSR\eictool_sign.ico"
        $Shortcut.Save()
    }
    if(!(Test-Path  $ShortcutPath_EicPrint)){
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($ShortcutPath_EicPrint)
        $Shortcut.TargetPath = "C:\eic\EICSignTSR\EicPrint.exe"
        $Shortcut.IconLocation = "C:\eic\EicPrint\EicPrint.ico"
        $Shortcut.Save()
    }
}#>

#HiCOS更新
    #不更新Hicos的電腦名稱
    $NotUpdateHicos_ComputerName=@()    
    if(!$NotUpdateHicos_ComputerName.Contains($env:computername)){
        powershell  "$env:SystemDrive\temp\HiCOS_Update.ps1"        
    }
#Chrome更新
    powershell "$env:SystemDrive\temp\ChromeUpdate.ps1"
#Adobe Reader更新
    powershell "$env:SystemDrive\temp\AdobeReaderUpdateV2.ps1"
#Java更新
    powershell "$env:SystemDrive\temp\JavaUpdate.ps1"
#7zip更新
    powershell "$env:SystemDrive\temp\7zipUpdateV3.ps1"
#經費結報系統元件GeasBatchsign更新
    powershell "$env:SystemDrive\temp\GeasBatchsign.ps1" 
#MariaDB ODBC Driver 64-bit更新
    powershell "$env:SystemDrive\temp\MariadbConnectorOdbc.ps1"
#ThreatSonarPC檢測
    #要安裝的電腦名稱
    #台數最多11台，如果要新增電腦，請從最後者加起，並刪除最前者。
    #$SpecificPC=@("TND-PEOF-057","TND-ASSE-031","TND-GASE-061","TND-SEOF-152","TND-ASSE-022","TND-ACOF-060","TND-GCSE-076","TND-DEPUTY-151","TND-GASE-085","TND-SEOF-065","TND-ASSE-031","TND-STOF-112","TND-CPSC-064")
    <#
        $SpecificPC=@("TND-STOF-113")
        if($SpecificPC.Contains($env:computername)){
            powershell "$env:SystemDrive\temp\ThreatSonarPCv2.ps1"
        }else{
            if(Get-ScheduledTask -TaskName "ThreatSonar" -ErrorAction Ignore){
                Unregister-ScheduledTask -TaskName "ThreatSonar" -Confirm:$False
            }
        }
    #>
    powershell "$env:SystemDrive\temp\ThreatSonarPCv2.ps1"
#印表機更名：將中括號[,]替換成【】，因為中括號容易造成字串字元判斷困難
    Get-printer |Where-Object{$_.Name -like  ("*"+[regex]::escape(']'))} |Where-Object{ Rename-Printer -name $_.Name -NewName ((($_.Name -replace  [regex]::escape('['),'【') -replace  [regex]::escape(']'),'】'))}

#PC基本資料蒐集
    powershell "$env:SystemDrive\temp\PCChecker.ps1" 
#刪除不要的軟體
    powershell "$env:SystemDrive\temp\UninstallSoftware.ps1" 

#安裝VANs軟體
     powershell "$env:SystemDrive\temp\WM7AssetCluster.ps1" 

#安裝新版軟體FileZilla
     powershell "$env:SystemDrive\temp\FileZillaUpdate.ps1" 

#安裝新版軟體XnView
     powershell "$env:SystemDrive\temp\XnViewUpdate.ps1" 
#安裝新版軟體K-LiteMegaCodecPack
     powershell "$env:SystemDrive\temp\K-LiteMegaCodecPackUpdate.ps1" 

#異地辦公室個人電腦匯入印表機設定
<#
$DormPC = @("TND-RMSE-047","TND-DEPUTY-151","TND-ACOF-040","TND-PEOF-031","TND-SASE-173","TND-SEOF-152","TND-GASE-055","TND-GASE-088","TND-GASE-044","TND-PEOF-030","TND-SASE-155","TND-BUSE-159","TND-GASE-045","TND-STOF-113","TND-ACOF-040","TND-ACOF-032","TND-5EES-068","TND-STOF-119")
$DormRemovePrinter = @("Kyocera ECOSYS P3050dn KX 【25號宿舍】","Kyocera ECOSYS P5025cdn KX 【20號宿舍】","Kyocera ECOSYS P3050dn KX 【19號宿舍】")

if($DormPC.Contains($env:computername)){
    (Get-printer).Name | Where-Object{ $DormRemovePrinter  -Contains $_} | where-object{ Remove-Printer -name $_ }
    powershell "$env:SystemDrive\temp\DormPrinterImport.ps1" 
}
#>


#指定電腦（們）匯入特定電腦的印表機封裝檔
#要匯入的電腦
#$InstallPrinterPC=@("TND-SASE-016","TND-SASE-089","TND-SASE-091","TND-SASE-095","TND-SASE-102","TND-SASE-107")
$InstallPrinterPC=@()
#指定哪個電腦的匯出當匯入範本
$ImportFromeComputername = "TND-STOF-113"
if($InstallPrinterPC.Contains($env:computername)){
    $PrinterExportFileLocation = "\\172.29.205.114\mig\Printer"
    $File_Name = $ImportFromeComputername +"x64.printerExport"
    $File_FullName = $PrinterExportFileLocation + "\" + $File_Name
    if(Test-Path $File_FullName){
        robocopy $PrinterExportFileLocation "$env:systemdrive\temp" $File_Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        $tempFile = "$env:systemdrive\temp\" + $File_Name
        start-Process cmd.exe -Verb RunAs -Args '/c',"${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe -R -F $tempFile -O FORCE" -Wait
        Get-printer | ForEach-Object{Set-Printer $_.Name -PermissionSDDL ((Get-Printer -Name 'Microsoft Print to PDF' -full).PermissionSDDL)}
        Remove-Item $tempFile -Force
        remove-item -Path ($PrinterExportFileLocation+"\"+$env:computername+"x64.printerExport") -Force
        #powershell "$env:SystemDrive\temp\SpecificPrintersImport.ps1"
    }
}

<#
    #修復Windows10 列印出現「藍白畫面」或無法完全列印。
    #參考連結https://3c.ltn.com.tw/news/43558
    if([environment]::OSVersion.Version.Major -like "10"){    
        powershell "$env:SystemDrive\temp\RepairWin32kfull.ps1" 
    }
#>
#Get-ChildItem -Path C:\Temp -Include * -Recurse -Force | foreach { $_.Delete()}
<#
if((Get-ScheduledTaskInfo -TaskName "PCinspection" -ErrorAction Ignore).LastTaskResult -ne 0){
    #年份
    $year = "110"
    #pc.exe存放位置
    $PCInspection_exeNASPath="\\172.29.205.114\loginscript\Update\PCinspection"
    #log存放位置
    $Log_Folder_Path = "\\172.29.205.114\Public\sources\audit\PCinspection" + '\' + $year

    #指訂排程執行的時與分（24H制）
    $SpecificTime = '12:10'
    if((get-date) -gt (get-date $SpecificTime)){
       $Exe_Date = (get-date).AddDays(1).ToString("yyyy-MM-dd")
    }else{
       $Exe_Date = (get-date).ToString("yyyy-MM-dd")
    }
    $DATEandTIME = $Exe_Date +"T"+ $SpecificTime + ":00"
    $Now_DATEandTIME = (get-date -format yyyy-MM-ddTHH:mm:ss).ToString()
    ((((((Get-Content -Path "${env:SystemDrive}\temp\PCinspectionTemplate.xml") ) -replace '%DATEandTIME%' , $DATEandTIME) -replace '%PCInspection_exeNASPath%', $PCInspection_exeNASPath) -replace "%Log_Folder_Path%",$Log_Folder_Path) -replace "%YEAR%",$year) -replace "%RegistrationInfoDate%", $Now_DATEandTIME | Set-Content -Path "${env:SystemDrive}\temp\PCinspection.xml" -Force
    #if((Get-ScheduledTask -TaskName "PCInspection") -ne $null){Unregister-ScheduledTask -TaskName "PCInspection" -Confirm:$false }   
    $schtasksOutput = schtasks.exe /create /RU "NT AUTHORITY\SYSTEM" /TN "PCInspection" /XML "${env:SystemDrive}\temp\PCinspection.xml" /F
    $schtasksOutput | Out-File ($Log_Folder_Path+"\"+$env:computername+"_schtaskOutput.txt")
    if(Test-Path -Path ($env:SystemDrive+"\temp\"+$env:computername+'_Result_Success' + ".txt")){robocopy ($env:SystemDrive+"\temp") $Log_Folder_Path ($env:computername+'_Result_Success' + ".txt") "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
    if(Test-Path -Path ($env:SystemDrive+"\temp\"+$env:computername+'_Result_Fail' + ".txt")){robocopy ($env:SystemDrive+"\temp")  $Log_Folder_Path ($env:computername+'_Result_Fail' + ".txt") "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
    #write-host $schtasksOutput
    #powershell "$env:SystemDrive\temp\PCinspection.ps1"
}
#>
#$Rebuid_EICSignTSR_PC=@("TND-BUSE-072")

$Rebuid_EICSignTSR_PC=@()
if($Rebuid_EICSignTSR_PC.Contains($env:computername)){   
    get-item -Path "$env:PUBLIC\EICSignTSR\EicSignTSR.ini"
    Stop-Process -Name EicSignTSR -Force -ErrorAction Continue
    Remove-Item -Path "$env:PUBLIC\EICSignTSR\EicSignTSR.ini" -Force
    Start-Process "$env:SystemDrive\eic\EICSignTSR\EicSignTSR.exe"
    Start-Sleep -s 3 
    Stop-Process -Name EicSignTSR -Force 
}

#指紋機驅動程式安裝
$FingerPrint_PC =@("TND-COU-141","TND-RMSE-142"."TND-GCSE-143","TND-CENTRAL-144","TND-GCSE-146","TND-GCSE-147")
#$FingerPrint_PC =@("TND-STOF-113")
#$FingerPrint_PC =@()
if($FingerPrint_PC.Contains($env:computername)){   
    $RegUninstallPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $Driver_RET = "\\172.29.205.114\loginscript\Update\Driver_RET"
    $Log_Path = "\\172.29.205.114\Public\sources\audit"    
    $DigitalPersona_installeds = Get-ItemProperty $RegUninstallPath | Where-Object{$_.DisplayName -match "DigitalPersona SDK Runtime"}               
    $Log_Folder_Path = $Log_Path +"\"+ "DigitalPersona SDK Runtime"
    if(($DigitalPersona_installeds -eq $null) -or ([Version]$DigitalPersona_installeds.DisplayVersion -lt  [Version]3.4.0.127)){
        robocopy $Driver_RET "$env:systemdrive\temp"  /E /XO /NJH /NJS /NDL /NC /NS | Out-Null

        #start-process $env:systemdrive\temp\RET\InstallOnly.bat
        $arguments="/s /v""REBOOT=ReallySuppress /qn /l*v $env:systemdrive\temp\$env:Computername" + "_ururte_install.log"""        
        start-process ($env:systemdrive+"\temp\RET\Setup.exe") -arg $arguments -wait -WindowStyle Hidden                
    }
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    $LogPattern="${env:Computername}"+"_ururte_install.log"
    if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
}

#安裝差勤上傳程式
<#
#差勤上傳機器電腦名稱"TND-PEOF-015"
if($env:computername -eq "TND-PEOF-015"){    
    Robocopy "\\172.29.205.114\loginscript\PSScripts" "$env:SystemDrive\temp" "Card.ps1" "/PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
    powershell "$env:SystemDrive\temp\Card.ps1"
}
#>
#印表機備份程式
    powershell  "$env:SystemDrive\temp\PrinterBackup.ps1"
#印表機權限
$Temp_PermissionSDDL="G:SYD:(A;;LCSWSDRCWDWO;;;WD)(A;OIIO;RPWPSDRCWDWO;;;WD)(A;;SWRC;;;AC)(A;CIIO;RC;;;AC)(A;OIIO;RPWPSDRCWDWO;;;AC)(A;;LCSWSDRCWDWO;;;CO)(A;OIIO;RPWPSDRCWDWO;;;CO)(A;OIIO;RPWPSDRCWDWO;;;BA)(A;;LCSWSDRCWDWO;;;BA)"
Get-printer | ForEach-Object{ Set-Printer $_.Name -PermissionSDDL $Temp_PermissionSDDL}
#同電腦名稱印表機匯入作業
<#
$ImportPrinterPC=@("TND-PEOF-030")
if ($ImportPrinterPC.Contains($env:computername)){    
    $PrinterExportFileLocation = "\\172.29.205.114\mig\Printer"
    $File_Name = $env:computername+"x64.printerExport"
    $File_FullName = $PrinterExportFileLocation + "\" + $File_Name
    robocopy $PrinterExportFileLocation "$env:systemdrive\temp" $File_Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    $tempFile = "$env:systemdrive\temp\" + $File_Name
    start-Process cmd.exe -Verb RunAs -Args '/c',"${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe -R -F $tempFile" -Wait
    Get-printer | ForEach-Object{ Set-Printer $_.Name -PermissionSDDL ((Get-Printer -Name 'Microsoft Print to PDF' -full).PermissionSDDL)}
    Remove-Item $tempFile
    Move-Item -Path $File_FullName -Destination "\\172.29.205.114\mig\Printer_BACKUP" -Force
}
#>
#蒐集SEP LOG檔
<#$SET_GetLOG_PC=@("TND-2EES-110","TND-WOME-079")
if($SET_GetLOG_PC.Contains($env:computername)){
    powershell "$env:SystemDrive\temp\getSEP_installLOG.ps1"
}#>
#powershell "$env:SystemDrive\temp\getSEP_installLOG.ps1"

#安裝財產申報系統-所長、副所長、會計主任、政風主任、總務科長
<#$Pdis_PC=@("TND-HEAD-150","TND-DEPUTY-014","TND-ACOF-060","TND-GEOF-131","TND-GASE-041")
if($Pdis_PC.Contains($env:computername)){
    powershell  "$env:SystemDrive\temp\Pdis.ps1"
}#>

#更新SEP版本至14.3.558.0000
<#
$Sep_Registry = "HKLM:\software\wow6432node\symantec\symantec endpoint protection\smc"
$Sep_NeedReboot_Registry = 'HKLM:\\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\RebootMgr'
if(((Get-ItemProperty -Path $Sep_Registry).ProductVersion -ne '14.3.558.0000') -and (!(Test-Path -Path  $Sep_NeedReboot_Registry))){
    powershell "$env:SystemDrive\temp\SEP_AutoUpdate.ps1"
}
#>
#檢查UAC有沒有開啟，沒開啟則開啟。
if((Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA -eq 0){
    Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name EnableLUA -Value 1
}

#進行WindowsUpdate
    #如果沒有PSWindowsUpdate模組則安裝
    if((Get-Module -ListAvailable -Name PSWindowsUpdate) -eq $null){
        $PSWindowsUpdate_Path = "\\172.29.205.114\loginscript\PSWindowsUpdate"
        $PSModule_Path = "$Env:ProgramFiles\WindowsPowerShell\Modules\PSWindowsUpdate"    
        robocopy $PSWindowsUpdate_Path $PSModule_Path "/e /XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        Import-Module PSWindowsUpdate             
    }
    $LogPath = "\\172.29.205.114\Public\sources\audit\WSUS"
    $temp = "$env:SystemDrive\temp"
    $ServiceID_WindowsUpdate = (Get-WUServiceManager | Where-Object{$_.Name -like "Windows Update"}).ServiceID
    $ServiceID_WSUS = (Get-WUServiceManager | Where-Object{$_.Name -like "Windows Server Update Service"}).ServiceID
    #$PSWUSettings = @{SmtpServer="smtp.moj.gov.tw";From="tndi@mail.moj.gov.tw";To="kyoiron@mail.moj.gov.tw";Port=25}
    start-job -ScriptBlock {    
        Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WSUS -Verbose *>&1 | Out-File "$env:SystemDrive\temp\${env:computername}_WindowsUpdate.txt" -Force -Append
    }
    Get-WUHistory | Out-File "$temp\${env:computername}_WindowsUpdate_History.txt" -Force
    Robocopy $temp $LogPath "${env:computername}_WindowsUpdate.txt" "${env:computername}_WindowsUpdate_History.txt" "/XO /NJH /NJS /NDL /NC /NS".Split(' ')| Out-Null

#Windows 7的話則安裝WMF5.1
$BuildVersion = [System.Environment]::OSVersion.Version
if($BuildVersion.Major -lt '10'){       
    Robocopy "\\172.29.205.114\loginscript\Update\Win7-KB3191566-x86" "$env:SystemDrive\temp" "Win7-KB3191566-x86.msu" "/XO /NJH /NJS /NDL /NC /NS".Split(' ')| Out-Null
    Robocopy "\\172.29.205.114\loginscript\Update\Win7-KB3191566-x86" "$env:SystemDrive\temp" "Install-WMF5.1.ps1" "/XO /NJH /NJS /NDL /NC /NS".Split(' ')| Out-Null
    powershell "$env:SystemDrive\temp\Install-WMF5.1.ps1" | out-file "$env:SystemDrive\temp\${env:computername}_WMF5.1_Log.txt"
    if("$env:SystemDrive\temp\${env:computername}_WMF5.1_Log.txt"){
        robocopy "$env:systemdrive\temp" "\\172.29.205.114\Public\sources\audit\WMF5.1\" "$env:SystemDrive\temp\${env:computername}_WMF5.1_Log.txt" "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        #Copy-Item -Path "$env:SystemDrive\temp\${env:computername}_WMF5.1_Log.txt" -Destination "\\172.29.205.114\Public\sources\audit\WMF5.1\"
    }
}
<#
if($env:COMPUTERNAME -eq "TND-HEAD-150"){
    $Bitlocker_Status = (manage-bde -status d:)
    if($Bitlocker_Status[8] -like "*已完全加密"){
        manage-bde -off d:
    }else{
       $Bitlocker_Status | Out-file $env:systemdrive\temp\"${env:COMPUTERNAME}_bitlockerStatus.txt" 
       $LogPattern="${env:Computername}"+"_bitlockerStatus.txt"      
       if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" "\\172.29.205.114\Public\sources\audit\BitlockerStatus" $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
    }
}
#>
