#$DormPC = @("TND-STOF-112","TND-BUSE-075","TND-RMSE-047","TND-DEPUTY-014","TND-ACOF-020","TND-PEOF-031","TND-SASE-102","TND-SEOF-062","TND-GASE-055","TND-GASE-088","TND-GASE-044","TND-ACOF-032","TND-PEOF-030","TND-SASE-089","TND-BUSE-159","TND-ACOF-040","TND-GASE-045")
$Printers_Name = @("Kyocera ECOSYS P3050dn KX 【25號宿舍】","Kyocera TASKalfa 3011i KX 【25號宿舍】","Kyocera ECOSYS P5025cdn KX 【20號宿舍】","Kyocera ECOSYS P3050dn KX 【19號宿舍】")
$NowPrinters = @()
$NowPrinters = (Get-printer).Name
#Get-printer|Where-Object{$_.Name -like "*宿舍]*"} | Where-Object{ Rename-Printer -name $_.Name -NewName ((($_.Name -replace  [regex]::escape('['),'【') -replace  [regex]::escape(']'),'】'))}
Rename-Printer -name "Kyocera TASKalfa 3011i KX 【20號宿舍】" -NewName "Kyocera TASKalfa 3011i KX 【25號宿舍】"

foreach($item in $Printers_Name){
    if(!($NowPrinters -Contains $item)){
        $PrinterExportFileLocation = "\\172.29.205.114\loginscript\Update\PrinterExport"
        $File_Name = "DormPrinterPackage"+"_x64.printerExport"
        $File_FullName = $PrinterExportFileLocation + "\" + $File_Name
        robocopy $PrinterExportFileLocation "$env:systemdrive\temp" $File_Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        $tempFile = "$env:systemdrive\temp\" + $File_Name
        start-Process cmd.exe -Verb RunAs -Args '/c',"${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe -R -F $tempFile" -Wait
        Get-printer | ForEach-Object{ Set-Printer $_.Name -PermissionSDDL ((Get-Printer -Name 'Microsoft Print to PDF' -full).PermissionSDDL)}
        Remove-Item $tempFile
        break
    }
}