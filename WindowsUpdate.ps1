#本檔案存參用。
if((Get-Module -ListAvailable -Name PSWindowsUpdate) -eq $null){
    #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    #Install-Module PSWindowsUpdate
    #Save-Module –Name PSWindowsUpdate –Path $PSWindowsUpdate_Path  
    $PSWindowsUpdate_Path = "\\172.29.205.114\loginscript\PSWindowsUpdate"
    $PSModule_Path = "$Env:ProgramFiles\WindowsPowerShell\Modules\PSWindowsUpdate"
    #$PSModule_Path = "$env:SystemRoot\system32\WindowsPowerShell\v1.0\Modules"    
    robocopy $PSWindowsUpdate_Path $PSModule_Path /e /XO /NJH /NJS /NDL /NC /NS | Out-Null
    Import-Module PSWindowsUpdate             
}
$LogPath = "\\172.29.205.114\Public\sources\audit\WSUS"
$temp = "$env:SystemDrive\temp"
$ServiceID_WindowsUpdate = (Get-WUServiceManager | Where-Object{$_.Name -like "Windows Update"}).ServiceID
$ServiceID_WSUS = (Get-WUServiceManager | Where-Object{$_.Name -like "Windows Server Update Service"}).ServiceID
#連線機關WSUS
#Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WSUS -Verbose  #| Out-File "c:\logs\$(get-date -f yyyy-MM-dd)-WindowsUpdate.log" -force
#連線機關Microsoft Windows Update
#Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot -MicrosoftUpdate -Verbose
#Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WindowsUpdate -Verbose
$PSWUSettings = @{SmtpServer="smtp.moj.gov.tw";From="tndi@mail.moj.gov.tw";To="kyoiron@mail.moj.gov.tw";Port=25}
start-job -ScriptBlock {
    #Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WSUS  -SendReport -SendHistory –PSWUSettings $PSWUSettings -Verbose
    Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WSUS -Verbose *>&1 | Out-File "$env:SystemDrive\temp\${env:computername}_WindowsUpdate.txt" -Force -Append
    Robocopy $temp $LogPath "${env:computername}_WindowsUpdate.txt" "/XO /NJH /NJS /NDL /NC /NS".Split(' ')
}
Get-WUHistory | Out-File "$temp\${env:computername}_WindowsUpdate_History.txt" -Force
Robocopy $temp $LogPath "${env:computername}_WindowsUpdate_History.txt" "/XO /NJH /NJS /NDL /NC /NS".Split(' ')


#Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WSUS -SendReport –PSWUSettings @{SmtpServer="smtp.moj.gov.tw";From="tndi@mail.moj.gov.tw";To="kyoiron@mail.moj.gov.tw";Port=25}  -Verbos

#Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WindowsUpdate -SendReport -SendHistory –PSWUSettings $PSWUSettings -ScheduleJob ((get-date).AddMinutes(1)) -Verbose
#Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WSUS -SendReport -SendHistory –PSWUSettings $PSWUSettings -ScheduleJob ((get-date).AddMinutes(1)) -Verbose

