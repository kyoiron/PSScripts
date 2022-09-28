$ImportFromeComputername = "TND-ASSE-020" 
$Printers_Name = @("Kyocera ECOSYS P3050dn KX 【孝一】")
$NowPrinters = @()
$NowPrinters = (Get-printer).Name
#Get-printer|Where-Object{$_.Name -like "*宿舍]*"} | Where-Object{ Rename-Printer -name $_.Name -NewName ((($_.Name -replace  [regex]::escape('['),'【') -replace  [regex]::escape(']'),'】'))}
<#
foreach($item in $Printers_Name){
    if(!($NowPrinters -Contains $item)){
        $PrinterExportFileLocation = "\\172.29.205.114\mig\Printer"
        $File_Name = $ImportFromeComputername+"x64.printerExport"
        $File_FullName = $PrinterExportFileLocation + "\" + $File_Name
        robocopy $PrinterExportFileLocation "$env:systemdrive\temp" $File_Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        $tempFile = "$env:systemdrive\temp\" + $File_Name
        start-Process cmd.exe -Verb RunAs -Args '/c',"${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe -R -F $tempFile" -Wait
        Get-printer | ForEach-Object{ Set-Printer $_.Name -PermissionSDDL ((Get-Printer -Name 'Microsoft Print to PDF' -full).PermissionSDDL)}
        Remove-Item $tempFile
        break
    }
}
#>
#Rename-Printer -name "Kyocera ECOSYS P3050dn KX" -NewName "Kyocera ECOSYS P3050dn KX【孝一】"
Remove-Printer -Name "FUJI XEROX DocuPrint P455 d 【輔導科辦公室】"
$PrinterExportFileLocation = "\\172.29.205.114\mig\Printer"
$File_Name = $ImportFromeComputername+"x64.printerExport"
$File_FullName = $PrinterExportFileLocation + "\" + $File_Name
robocopy $PrinterExportFileLocation "$env:systemdrive\temp" $File_Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
$tempFile = "$env:systemdrive\temp\" + $File_Name
start-Process cmd.exe -Verb RunAs -Args '/c',"${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe -R -F $tempFile" -Wait
Get-printer | ForEach-Object{Set-Printer $_.Name -PermissionSDDL ((Get-Printer -Name 'Microsoft Print to PDF' -full).PermissionSDDL)}
Remove-Item $tempFile
