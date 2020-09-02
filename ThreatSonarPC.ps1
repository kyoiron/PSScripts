$SpecificPC=@("TND-GEOF-131","TND-GEOF-132","TND-SASE-089","TND-SASE-107","TND-GASE-023","TND-PEOF-029","TND-PEOF-030","TND-PEOF-057","TND-ASSE-031","TND-GASE-061")
if($SpecificPC.Contains($env:computername)){
    $Name="台南看守所"
    $ThreatSonar_Path = "$env:SystemDrive\ThreatSonar"
    $ThreatSonar_exe = "$ThreatSonar_Path\ThreatSonar_073.exe"
    $ThreatSonar_bat = "$ThreatSonar_Path\TS_scan.bat"
    if(!(Test-Path -Path $ThreatSonar_Path)){New-Item -ItemType directory -Path $ThreatSonar_Path}
    $url_exe = "http://download.moj/files/工具檔案/T5/所屬PC/$Name-PC/ThreatSonar/ThreatSonar_073.exe"
    $url_bat = "http://download.moj/files/工具檔案/T5/所屬PC/$Name-PC/ThreatSonar/TS_scan.bat"
    Invoke-WebRequest -Uri $url_exe -OutFile $ThreatSonar_exe
    Invoke-WebRequest -Uri $url_bat -OutFile $ThreatSonar_bat
    start-process "cmd.exe" "/c $ThreatSonar_bat" -Wait
    if(Get-ScheduledTask -TaskName "ThreatSonar" -ErrorAction Ignore){ 
        $ScheduledTask_Check = "ThreatSonar排程已建立"
    }else{
        $ScheduledTask_Check = "ThreatSonar排程不存在，請重新檢查"
    }
    if(Test-Path -Path $ThreatSonar_exe){
        $ThreatSonarExe_Check = "$ThreatSonar_exe 檔案存在"
    }else{
        $ThreatSonarExe_Check = "$ThreatSonar_exe 檔案不存在"
    }    
    (get-date).ToString() + "`r`n$ScheduledTask_Check`r`n$ThreatSonarExe_Check`r`n" | Out-File -FilePath "$env:SystemDrive\temp\${env:COMPUTERNAME}_ThreatSonar_Check.txt"
}