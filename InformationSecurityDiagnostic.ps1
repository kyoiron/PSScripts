#設定變數
$Year = '108'
$PcFolder = $env:SystemDrive + "\temp\"
$PcIsdFolder = $PcFolder  + "3D36C52EB2"
$NetworkFolder = "\\172.29.205.114\Public\sources\audit\InformationSecurityDiagnostic"
$Logfile = $NetworkFolder + "\" + $Year + "\" + $env:computername + ".txt"

#拷貝資安健檢執行檔
robocopy $NetworkFolder $PcFolder '法務部矯正署臺南看守所pc.exe' /PURGE /XO /NJH /NJS /NDL /NC /NS
#執行前刪除資料夾
if (Test-Path ($PcFolder + "3D36C52EB2")){Remove-Item -LiteralPath $PcIsdFolder -Force -Recurse}

#執行資安健檢檔start-job
$exe = $PcFolder+"法務部矯正署臺南看守所pc.exe"   

$ISDJob = Start-Job -ScriptBlock { 
    $exe = $args[0] + "法務部矯正署臺南看守所pc.exe"   
    write-host $exe
    Start-Process $exe -Verb 'RunAs'  
} -ArgumentList $PcFolder

#檢查edcprobeagent.exe、Threat_Sonar.exe、GPOAction.exe是否已經結束。
#結束代表檢查已完成。
do{
    Start-Sleep -Seconds 60
    $Threat_Sonar = Get-WmiObject Win32_Process | Where-Object{$_.name -eq 'Threat_Sonar.exe'}
    $edcprobeagent = Get-WmiObject Win32_Process | Where-Object{$_.name -eq 'edcprobeagent.exe'}
    $GPOAction = Get-WmiObject Win32_Process | Where-Object{$_.name -eq 'GPOAction.exe'}    
}while(($Threat_Sonar) -or ($edcprobeagent) -or ($GPOAction))

#檢查資料夾是否為空，空則成功；否則
$directoryInfo = Get-ChildItem $PcIsdFolder -ErrorAction SilentlyContinue | Measure-Object
if($directoryInfo.count -eq 0){
    ""| Out-File -FilePath $Logfile
}else{
    "3D36C52EB2資料夾不為空，可能檢測未成功" | Out-File -FilePath $Logfile
}