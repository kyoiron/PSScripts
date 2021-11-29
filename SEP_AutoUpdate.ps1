$SEP_EXE_Path = "\\172.29.205.114\loginscript\Update\SEP"
$SEP_EXE = Get-ChildItem -Path ($SEP_EXE_Path+"\*.exe") | Select-Object -first 1
robocopy $SEP_EXE_Path "$env:systemdrive\temp" $SEP_EXE.Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
Start-Process -FilePath ($env:systemdrive+"\temp\"+$SEP_EXE.Name)