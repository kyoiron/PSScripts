$Year = "110";$PcFolder=$env:SystemDrive+"\temp";$CheckFolder="3D36C52EB2";$CheckFolder_Path=$PcFolder+"\"+$CheckFolder;$NetworkFolder="\\172.29.205.114\Public\sources\audit\PCinspection";$PCInspection_exePath = $PcFolder+"\"+"pc.exe"
if (Test-Path $CheckFolder_Path){Remove-Item -LiteralPath $CheckFolder_Path -Force -Recurse}
if(Test-Path($PCInspection_exePath)){    
    $ISDJob = Start-Job -ScriptBlock {Start-Process $args[0] -Verb 'RunAs'} -ArgumentList $PCInspection_exePath
    do{Start-Sleep -Seconds 60;$Threat_Sonar = Get-WmiObject Win32_Process | Where-Object{$_.name -eq 'Threat_Sonar.exe'};$edcprobeagent = Get-WmiObject Win32_Process | Where-Object{$_.name -eq 'edcprobeagent.exe'};$GPOAction = Get-WmiObject Win32_Process | Where-Object{$_.name -eq 'GPOAction.exe'}}while(($Threat_Sonar) -or ($edcprobeagent) -or ($GPOAction))    
    #檢查資料夾是否為空，空則成功；否則
    $directoryInfo = Get-ChildItem $CheckFolder_Path  -ErrorAction SilentlyContinue | Measure-Object;
    if($directoryInfo.count -eq 0){$Logfile=$NetworkFolder+"\"+$Year+"\"+$env:computername+'_Result_Success'+".txt";"get-date" | Out-File -FilePath $Logfile}else{$Logfile=$NetworkFolder+"\"+ $Year +"\"+ $env:computername+'_Result_Fail'+".txt";"3D36C52EB2資料夾不為空，可能檢測未成功" | Out-File -FilePath $Logfile}}
