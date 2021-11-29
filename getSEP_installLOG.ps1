$Path_Destination_AllLog = "\\172.29.205.114\Public\sources\audit\SEP_Install_LOG" 
$Path_Sources_SIS_INST_log_Path = (Get-Item -Path "$env:ProgramData\Symantec\Symantec Endpoint Protection\14.*\Data\Install\Logs").FullName
$FileName_SIS_INST_log = "SIS_INST.log"
Copy-Item -Path ($Path_Sources_SIS_INST_log_Path + "\" + $FileName_SIS_INST_log ) -Destination ($Path_Destination_AllLog+"\"+$env:COMPUTERNAME+"_"+$FileName_SIS_INST_log) -ErrorAction SilentlyContinue
#Rename-Item -Path ($Path_Destination_AllLog + "\"+$FileName_SIS_INST_log) -NewName ($env:COMPUTERNAME+"_"+$FileName_SIS_INST_log ) -ErrorAction SilentlyContinue

$FileName_SEP_INST_log = "SEP_INST.log"
Copy-Item -Path ($env:temp + "\" + $FileName_SEP_INST_log) -Destination ($Path_Destination_AllLog+"\"+ $env:COMPUTERNAME+"_TEMP_"+$FileName_SEP_INST_log )  -ErrorAction SilentlyContinue
#Rename-Item -Path ($Path_Destination_AllLog + "\"+ $FileName_SEP_INST_log ) -NewName ($env:COMPUTERNAME+"_TEMP_"+$FileName_SEP_INST_log)  -ErrorAction SilentlyContinue
Copy-Item -Path ($env:systemroot + "\temp\" + $FileName_SEP_INST_log) -Destination ($Path_Destination_AllLog+"\"+$env:COMPUTERNAME+"_SYSTEMROOT_"+$FileName_SEP_INST_log)  -ErrorAction SilentlyContinue
#Rename-Item -Path ($Path_Destination_AllLog + "\"+ $FileName_SEP_INST_log ) -NewName ($env:COMPUTERNAME+"_SYSTEMROOT_"+$FileName_SEP_INST_log)  -ErrorAction SilentlyContinue
