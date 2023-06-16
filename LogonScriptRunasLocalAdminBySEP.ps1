$lastBootUpTime = (gcim Win32_OperatingSystem).LastBootUpTime
$uptime = (Get-Date) - $lastBootUpTime
$Day_Threshold = 7
if ($uptime.TotalDays -ge $Day_Threshold) {
    # 電腦自從上一次重開機以來已經運行已經超過7天
    $taskName = "LogonScripts"  # 將YourTaskName替換為您的排程名稱
    Start-ScheduledTask  "LogonScripts"
}