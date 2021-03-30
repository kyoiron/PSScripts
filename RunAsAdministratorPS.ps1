#HiCOS更新
powershell  "$env:SystemDrive\temp\HiCOS_Update.ps1"
#Chrome更新
powershell "$env:SystemDrive\temp\ChromeUpdate.ps1"
#Adobe Reader更新
powershell "$env:SystemDrive\temp\AdobeReaderUpdate.ps1"
#Java更新
powershell "$env:SystemDrive\temp\JavaUpdate.ps1"
#ThreatSonarPC檢測
powershell "$env:SystemDrive\temp\ThreatSonarPC.ps1"
#PC基本資料蒐集
powershell "$env:SystemDrive\temp\PCChecker.ps1" 
#刪除不要的軟體
powershell "$env:SystemDrive\temp\UninstallSoftware.ps1" 

#修復Windows10 列印出現「藍白畫面」或無法完全列印。
#參考連結https://3c.ltn.com.tw/news/43558
if([environment]::OSVersion.Version.Major -like "10"){    
    powershell "$env:SystemDrive\temp\RepairKB5000802.ps1" 
}