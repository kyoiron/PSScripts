if([System.Environment]::OSVersion.Version.Major -lt "10"){exit}
$Local_Output_Path = "$env:systemdrive\temp"
$Nas_Output_Path = "\\172.29.205.114\mig\Printer"
$Nas_Output_Bak_Path ="\\172.29.205.114\mig\Printer_BACKUP"
$PrinterExport_FileName = $env:computername+"x64.printerExport"
$Local_Output_FileName = $Local_Output_Path  + "\"  + $PrinterExport_FileName
$Nas_Output_FileName = $Nas_Output_Path  + "\"  + $PrinterExport_FileName
$Nas_Output_Bak_Path_FileName  = $Nas_Output_Bak_Path + "\"  + $PrinterExport_FileName
#保留幾天內的匯出檔
$Days = 30
$Currentlytime = (Get-Date).AddDays(-$Days)
if(Test-Path $Nas_Output_FileName){
    if((Get-ChildItem -path $Nas_Output_FileName).CreationTime -gt $Currentlytime) {
        exit
    }else{
        Move-Item -Path $Nas_Output_FileName -Destination $Nas_Output_Bak_Path_FileName -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $Nas_Output_FileName -Force -ErrorAction SilentlyContinue
    }
}
#將本機印表機檔案匯出
    #刪除不使用的IP Port
        $IPs_all = (Get-PrinterPort |Where-Object {$_.CimClass -like "ROOT/StandardCimv2:MSFT_TcpIpPrinterPort"}).Name
        $IPs_Used = (Get-Printer ).PortName
        $IPs_Delete = $IPs_all | Where-Object {$IPs_Used -notcontains $_}
        $IPs_Delete | Where-Object {Remove-PrinterPort -Name $_}
    #執行備份指令
        Start-Process cmd.exe -Verb RunAs -Args '/c',"title 進行印表機匯出程序，請稍後... & del $Local_Output_FileName /F /Q & ${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe -B -F $Local_Output_FileName" -Wait 
        Robocopy  $Local_Output_Path $Nas_Output_Path $PrinterExport_FileName "/PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null