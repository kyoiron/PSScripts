#LOG檔NAS存放路徑
    $Log_Path = "\\172.29.205.114\Public\sources\audit"
#安裝軟體允許清單位址
    $SoftwareDisallowList_Path = "\\172.29.205.114\loginscript\PSScripts\SoftwareDisallowList.txt"
    #$SoftwareAllowList_Path = "$env:SystemDrive\temp\SoftwareAllowList.txt"

New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null
$RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
$RegUninstallPaths += Get-ChildItem HKU: | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object {"HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"}
$Softwares = @()
foreach($Path in $RegUninstallPaths){        
    $Softwares += (Get-ItemProperty $Path | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate , PSPath , QuietUninstallString , UninstallString , UninstallString_Hidden)
}
Remove-PSDrive -Name HKU

$BlackList_Software_DisplayName = Get-Content -Path $SoftwareDisallowList_Path -Encoding UTF8
if($null -ne $BlackList_Software_DisplayName){
    $UninstallSoftwares = @()
    $UninstallSoftwares = $Softwares | Where-Object {$_.displayName -in $BlackList_Software_DisplayName}    
    $UninstallSoftwares | Where-Object{
        if($null -ne $_.QuietUninstallString){
            $QuietUninstallString = $_.QuietUninstallString
            $QuietUninstallStringEXE = ($QuietUninstallString.trim() -split '"')[1]
            ($QuietUninstallString.trim() -split '"').length
            $QuietUninstallStringArgs=""
            for ($i=2; $i -lt ($QuietUninstallString.trim() -split '"').length; $i++) {
    	        $QuietUninstallStringArgs = $QuietUninstallStringArgs +" "+($QuietUninstallString.trim() -split '"')[$i]
            }
            $QuietUninstallStringArgs = $QuietUninstallStringArgs.Trim()    
            start-process -FilePath $QuietUninstallStringEXE  -ArgumentList "$QuietUninstallStringArgs" -Wait -WindowStyle Hidden           
        }else{
            $Log_Folder_Path =  $Log_Path +"\"+ $_.DisplayName
            $Uninstall_LogName = $env:Computername + "_"+ $_.DisplayName +"_Remove" + ".txt" 
            if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
            if($_.UninstallString -match "msiexec.exe"){
                $uninstall = $_.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
                $uninstall = $uninstall.Trim()                
                start-process "msiexec.exe" -arg "/X $uninstall /qn /log ""$env:systemdrive\temp\$Uninstall_LogName""" -Wait -WindowStyle Hidden                                                                 
            }elseif($_.UninstallString -notmatch "/S" ){
                $uninstall = $_.UninstallString.Trim()             
                start-process -FilePath $uninstall -ArgumentList " /S"  -Wait -WindowStyle Hidden 
                "加入參數/S，嘗試移除：" + $_.DisplayName | Out-File  "$env:systemdrive\temp\$Uninstall_LogName"  
            }
            else{
                
                "無靜默安裝功能，無法移除軟體：" + $_.DisplayName | Out-File  "$env:systemdrive\temp\$Uninstall_LogName"
            } 
            if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $Uninstall_LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}          
        }
    }
    New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null
    $Softwares_AfterRomve = @()
    foreach($Path in $RegUninstallPaths){        
        $Softwares_AfterRemove += (Get-ItemProperty $Path | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate , PSPath , QuietUninstallString , UninstallString , UninstallString_Hidden)
    }
    Remove-PSDrive -Name HKU
    $UninstallSoftwares_Check = @()
    $UninstallSoftwares_Check = $Softwares_AfterRemove| Where-Object {$_.displayName -in $BlackList_Software_DisplayName}
    $Log_Folder_Path_Check =  $Log_Path +"\"+ "DisallowList_ALLClear"
    $Uninstall_LogName_Check  = $env:Computername + "_DisallowList_ALLClear_Remove" + ".txt"
    if(!(Test-Path -Path $Log_Folder_Path_Check)){ New-Item -ItemType Directory -Path $Log_Folder_Path_Check -Force}    
    if($null -eq $UninstallSoftwares_Check){        
        "成功移除所有黑名單軟體"  | Out-File  "$env:systemdrive\temp\$Uninstall_LogName_Check"
    }else{                    
        $UninstallSoftwares_Check | Where-Object{ 
                $Log_Folder_Path_Check  =  $Log_Path +"\"+ $_.DisplayName ;
                $Uninstall_LogName = $env:Computername + "_"+ $_.DisplayName +"_NotRemove" + ".txt" ;
                "尚未移除：" + $_.DisplayName+"軟體" | Out-File  "$env:systemdrive\temp\$Uninstall_LogName";
                 if(Test-Path -Path "$env:systemdrive\temp\$Uninstall_LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path_Check $Uninstall_LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
            }
    }
    if(Test-Path -Path "$env:systemdrive\temp\$Uninstall_LogName_Check"){ robocopy "$env:systemdrive\temp" $Log_Folder_Path_Check $Uninstall_LogName_Check "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}    
}