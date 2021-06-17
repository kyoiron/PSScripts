#筆硯新版元件
    powershell  "$env:SystemDrive\temp\eic.ps1" 
#HiCOS更新
    powershell  "$env:SystemDrive\temp\HiCOS_Update.ps1"
#Chrome更新
    powershell "$env:SystemDrive\temp\ChromeUpdate.ps1"
#Adobe Reader更新
    powershell "$env:SystemDrive\temp\AdobeReaderUpdate.ps1"
#Java更新
    powershell "$env:SystemDrive\temp\JavaUpdate.ps1"
#ThreatSonarPC檢測
    #要安裝的電腦名稱
    #台數最多11台，如果要新增電腦，請從最後者加起，並刪除最前者。
    $SpecificPC=@("TND-GASE-023","TND-PEOF-029","TND-PEOF-030","TND-PEOF-057","TND-ASSE-031","TND-GASE-061","TND-SEOF-152","TND-ASSE-022","TND-ACOF-060","TND-GCSE-076","TND-DEPUTY-151","TND-GASE-085")
    if($SpecificPC.Contains($env:computername)){
        powershell "$env:SystemDrive\temp\ThreatSonarPC.ps1"
    }else{
        if(Get-ScheduledTask -TaskName "ThreatSonar" -ErrorAction Ignore){
            Unregister-ScheduledTask -TaskName "ThreatSonar" -Confirm:$False
        }
    }
#印表機更名：將中括號[,]替換成【】，因為中括號容易造成字串字元判斷困難
    Get-printer |Where-Object{$_.Name -like  ("*"+[regex]::escape(']'))} |Where-Object{ Rename-Printer -name $_.Name -NewName ((($_.Name -replace  [regex]::escape('['),'【') -replace  [regex]::escape(']'),'】'))}

#PC基本資料蒐集
    powershell "$env:SystemDrive\temp\PCChecker.ps1" 
#刪除不要的軟體
    powershell "$env:SystemDrive\temp\UninstallSoftware.ps1" 

#異地辦公室個人電腦匯入印表機設定
$DormPC = @("TND-STOF-138","TND-BUSE-075","TND-RMSE-047","TND-DEPUTY-151","TND-ACOF-020","TND-PEOF-031","TND-SASE-111","TND-SEOF-152","TND-GASE-055","TND-GASE-088","TND-GASE-044","TND-ACOF-032","TND-PEOF-030","TND-SASE-155","TND-BUSE-159","TND-ACOF-040","TND-GASE-045","TND-GCSE-086")
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

if((Get-ScheduledTaskInfo -TaskName "PCinspection" -ErrorAction Ignore).LastTaskResult -ne 0){
    #年份
    $year = "110"
    #pc.exe存放位置
    $PCInspection_exeNASPath="\\172.29.205.114\loginscript\Update\PCinspection"
    #log存放位置
    $Log_Folder_Path = "\\172.29.205.114\Public\sources\audit\PCinspection"

    #指訂排程執行的時與分（24H制）
    $SpecificTime = '12:10'
    if((get-date) -gt (get-date $SpecificTime)){
       $Exe_Date = (get-date).AddDays(1).ToString("yyyy-MM-dd")
    }else{
       $Exe_Date = (get-date).ToString("yyyy-MM-dd")
    }
    $DATEandTIME = $Exe_Date +"T"+ $SpecificTime + ":00"
    (((((Get-Content -Path "${env:SystemDrive}\temp\PCinspectionTemplate.xml") ) -replace '%DATEandTIME%' , $DATEandTIME) -replace '%PCInspection_exeNASPath%', $PCInspection_exeNASPath) -replace "%Log_Folder_Path%",$Log_Folder_Path) -replace "%YEAR%",$year | Set-Content -Path "${env:SystemDrive}\temp\PCinspection.xml" -Force    
    if((Get-ScheduledTask -TaskName "PCInspection") -ne $null){Unregister-ScheduledTask -TaskName "PCInspection" -Confirm:$false }
    $schtasksOutput = schtasks.exe /create /RU "NT AUTHORITY\SYSTEM" /TN "PCInspection" /XML "${env:SystemDrive}\temp\PCinspection.xml" /F
    #write-host $schtasksOutput
    #powershell "$env:SystemDrive\temp\PCinspection.ps1"
}