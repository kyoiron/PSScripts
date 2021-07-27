$Local_Output_Path = "$env:systemdrive\temp"
$Nas_Output_Path = "\\172.29.205.114\mig\Printer"
$PrinterExport_FileName = $env:computername+"x64.printerExport"
$Local_Output_FileName = $Local_Output_Path  + "\"  + $PrinterExport_FileName
$Nas_Output_FileName = $Nas_Output_Path  + "\"  + $PrinterExport_FileName
if(!(Test-Path $Nas_Output_FileName)){
    if([System.Environment]::OSVersion.Version.Major -eq "10"){
        Start-Process cmd.exe -Verb RunAs -Args '/c',"title 進行印表機匯出程序，請稍後... & del $Local_Output_FileName /F /Q & ${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe -B -F $Local_Output_FileName" -Wait 
        Robocopy  $Local_Output_Path $Nas_Output_Path $PrinterExport_FileName "/PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
    }
}