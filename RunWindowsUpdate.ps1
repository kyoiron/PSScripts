if((Get-Module -ListAvailable -Name PSWindowsUpdate) -eq $null){
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-Module PSWindowsUpdate          
}
$PSWUSettings = @{SmtpServer="smtp.moj.gov.tw";From="tndi@mail.moj.gov.tw";To="tndi@mail.moj.gov.tw";Port=25}
Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot -MicrosoftUpdate -SendReport -PSWUSettings $PSWUSettings 
Get-WUHistory -SendReport -PSWUSettings $PSWUSettings