New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null
$RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
$RegUninstallPaths += Get-ChildItem HKU: | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object {"HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"}
$Software_installeds =@()
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$Log_Folder_Path = $Log_Path +"\Going"
foreach ($Path in $RegUninstallPaths) {
    if (Test-Path $Path) {
        $Software_installeds += Get-ItemProperty $Path | Where-Object{ $_.DisplayName -like "Natural Input Method Professional Edition*"} 
    }
}
Remove-PSDrive -Name HKU
if($Software_installeds.Count -gt 0){
    if(!(Test-Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
     foreach($item in $Software_installeds){       
        $LogFileName =$env:Computername + "_"+$item.DisplayName + "_" + $item.DisplayVersion  +".txt"
        $item | Out-File -FilePath  "$env:systemdrive\temp\$LogFileName"
        if(Test-Path "$env:systemdrive\temp\$LogFileName") { Move-Item  "$env:systemdrive\temp\$LogFileName" -Destination $Log_Folder_Path -Force }
     }         
}

<#
    
#>