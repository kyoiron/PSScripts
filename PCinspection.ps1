#設定變數
$Year = '110'
$PCInspection_exeNASPath="\\172.29.205.114\loginscript\Update\PCinspection"
$PcFolder=$env:SystemDrive+"\temp";$CheckFolder='3D36C52EB2';$CheckFolder_Path=$PcFolder+'\'+$CheckFolder;
$Log_Folder_Path = "\\172.29.205.114\Public\sources\audit\PCinspection"+"\"+"$Year"
#$Logfile = $NetworkFolder + "\" + $Year + "\" + $env:computername+'_Result' + ".txt"
$PCInspection_exePath = $PcFolder+'\'+"pc.exe"

#執行前刪除資料夾
if (Test-Path $CheckFolder_Path){Remove-Item -LiteralPath $CheckFolder_Path -Force -Recurse}
robocopy $PCInspection_exeNASPath "pc.exe" "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
if(Test-Path($PCInspection_exePath)){
    #執行資安健檢檔start-job  
    $ISDJob = Start-Job -ScriptBlock { 
        write-host $args[0]
        Start-Process $args[0] -Verb 'RunAs'
    } -ArgumentList $PCInspection_exePath
    #檢查edcprobeagent.exe、Threat_Sonar.exe、GPOAction.exe是否已經結束。
    #結束代表檢查已完成。
    do{
        Start-Sleep -Seconds 60
        $Threat_Sonar = Get-WmiObject Win32_Process | Where-Object{$_.name -eq 'Threat_Sonar.exe'}
        $edcprobeagent = Get-WmiObject Win32_Process | Where-Object{$_.name -eq 'edcprobeagent.exe'}
        $DSWAgent = Get-WmiObject Win32_Process | Where-Object{$_.name -eq 'DSWAgent.exe'}    
    }while(($Threat_Sonar) -or ($edcprobeagent) -or ($DSWAgent))
    
    #檢查資料夾是否為空，空則成功；否則
    $directoryInfo = Get-ChildItem $CheckFolder_Path  -ErrorAction SilentlyContinue | Measure-Object
    if((Test-Path -path $CheckFolder_Path) -and ($directoryInfo.count -eq 0)){
        $Logfile =  $env:computername+'_Result_Success' + ".txt"
        (get-date).ToString() + " 檢查3D36C52EB2資料夾為空"| Out-File -FilePath ($PcFolder+"\"+$Logfile)
    }else{
        $Logfile =  $env:computername+'_Result_Fail' + ".txt"    
        "3D36C52EB2資料夾不為空，可能檢測未成功" | Out-File -FilePath ($PcFolder+"\"+$Logfile)
    }
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    robocopy $PcFolder ($Log_Folder_Path+"\"+$Year) $Logfile "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
}
