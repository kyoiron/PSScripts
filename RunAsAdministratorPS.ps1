#將tnduser加入本機管理者帳號
Add-LocalGroupMember -Group "Administrators" -Member "$env:COMPUTERNAME\tnduser" -ErrorAction SilentlyContinue

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
    powershell  "$env:SystemDrive\temp\HiCOS_Update.ps1"
#Chrome更新
    powershell "$env:SystemDrive\temp\ChromeUpdate.ps1"
#Adobe Reader更新
    powershell "$env:SystemDrive\temp\AdobeReaderUpdate.ps1"
#Java更新
    #powershell "$env:SystemDrive\temp\JavaUpdate.ps1"
#經費結報系統元件GeasBatchsign更新
    powershell "$env:SystemDrive\temp\GeasBatchsign.ps1" 
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

#異地辦公室個人電腦匯入印表機設定
$DormPC = @("TND-RMSE-047","TND-DEPUTY-151","TND-ACOF-020","TND-PEOF-031","TND-SASE-111","TND-SEOF-152","TND-GASE-055","TND-GASE-088","TND-GASE-044","TND-ACOF-032","TND-PEOF-030","TND-SASE-155","TND-BUSE-159","TND-ACOF-040","TND-GASE-045","TND-GCSE-086","TND-GCSE-051","TND-STOF-112")
if($DormPC.Contains($env:computername)){
    powershell "$env:SystemDrive\temp\DormPrinterImport.ps1" 
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

#蒐集SEP LOG檔
<#$SET_GetLOG_PC=@("TND-2EES-110","TND-WOME-079")
if($SET_GetLOG_PC.Contains($env:computername)){
    powershell "$env:SystemDrive\temp\getSEP_installLOG.ps1"
}#>
#powershell "$env:SystemDrive\temp\getSEP_installLOG.ps1"
#安裝財產申報系統-所長、副所長、會計主任、政風主任、總務科長
$Pdis_PC=@("TND-HEAD-150","TND-DEPUTY-014","TND-ACOF-060","TND-GEOF-131","TND-GASE-041")
if($Pdis_PC.Contains($env:computername)){
    powershell  "$env:SystemDrive\temp\Pdis.ps1"
}
#更新SEP版本至14.3.558.0000
$Sep_Registry = "HKLM:\software\wow6432node\symantec\symantec endpoint protection\smc"
$Sep_NeedReboot_Registry = 'HKLM:\\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\RebootMgr'
if(((Get-ItemProperty -Path $Sep_Registry).ProductVersion -ne '14.3.558.0000') -and (!(Test-Path -Path  $Sep_NeedReboot_Registry))){
    powershell "$env:SystemDrive\temp\SEP_AutoUpdate.ps1"
}