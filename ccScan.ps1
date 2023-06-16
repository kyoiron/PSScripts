$ccScan_PC =@("*")
$ccScan_Path = "\\172.29.205.114\loginscript\Update\ccScan"
$ccScan_FileName = "ccScan_TND_PC.exe"
if($ccScan_PC.Contains($env:computername) -or $ccScan_PC.Contains("*")){
    if(!(test-path -path  "$env:systemdrive\temp\corecloud")){
        if(!(test-path -path "$env:systemdrive\temp\$ccScan_FileName")){
            Robocopy $ccScan_Path "$env:SystemDrive\temp" $ccScan_FileName "/XO /NJH /NJS /NDL /NC /NS".Split(' ')| Out-Null
            unblock-file ($env:systemdrive+"\temp\"+$ccScan_FileName)
        }
        #Start-Job -ScriptBlock { Start-Process "$env:systemdrive\temp\$ccScan_FileName" -NoNewWindow}
        Start-Process "$env:systemdrive\temp\$ccScan_FileName" 
    }
}