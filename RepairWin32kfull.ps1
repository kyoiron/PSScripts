#if([environment]::OSVersion.Version.Major -notlike "10"){exit}
#LOG檔NAS存放路徑
    $Log_Path = "\\172.29.205.114\Public\sources\audit"
#Patch檔NAS存放路徑
    $Patch_Folder = "\\172.29.205.114\loginscript\Update\WindowsUpdatePatch"

$Version = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' –Name ReleaseID –ErrorAction Stop).ReleaseID
Switch ($Version){
    '1607' {$Patch_KeyWord = 'kb5001633';$Problem_Patch='KB5000803';}
    '1809' {$Patch_KeyWord = 'kb5001634';$Problem_Patch='KB5000809';}
}
if((Get-HotFix -id $Problem_Patch -ErrorAction Ignore) -ne $null){
    if((Get-HotFix -id $Patch_KeyWord -ErrorAction Ignore) -eq $null){
        $Patch_File = Get-ChildItem -Path $Patch_Folder | Where-Object{$_.Name -match $Patch_KeyWord}
        robocopy $Patch_Folder "$env:systemdrive\temp" $Patch_File.Name /XO /NJH /NJS /NDL /NC /NS 
        $execute_Patch = "$env:systemdrive\temp\" + $Patch_File.Name
        $Log_FILE_Name =  $env:Computername + "_" + $Problem_Patch + "修補.txt"
        $Log_FILE_PathAndName = "$env:systemdrive\temp\" + $Log_FILE_Name 
        #wusa.exe "$execute_Patch /quiet /norestart /log:""$Log_FILE_PathAndName"""
        #wusa.exe $execute_Patch /log:"$Log_FILE_PathAndName"
        start-process "wusa.exe" -arg "$execute_Patch /quiet /norestart /log:""$Log_FILE_PathAndName""" -Wait -WindowStyle Hidden    
        $Log_Folder_Path = $Log_Path + "\" + $Patch_KeyWord 
        if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force} 
        if(Test-Path -Path $Log_Folder_Path){robocopy "$env:systemdrive\temp" $Log_Folder_Path  $Log_FILE_Name /XO /NJH /NJS /NDL /NC /NS}   
    }
}