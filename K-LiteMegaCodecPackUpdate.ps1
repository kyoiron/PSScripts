# Define variables
$RemoveFirstPC = @()
$KLiteMegaCodecPacks_Path = "\\172.29.205.114\loginscript\Update\KLiteMegaCodecPack"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$Force_Install = $false

# Function to get the latest K-Lite Mega Codec Pack installer
function Get-LatestKLiteInstaller {
    Get-ChildItem -Path "$KLiteMegaCodecPacks_Path\*.exe" | 
    Where-Object { $_.VersionInfo.ProductName.Trim() -eq "K-Lite Mega Codec Pack" } | 
    Sort-Object -Property VersionInfo -Descending | 
    Select-Object -First 1
}

# Function to get installed K-Lite Codec Packs
function Get-InstalledKLiteCodecPacks {
    $RegPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $RegPaths | ForEach-Object {
        if (Test-Path $_) {
            Get-ItemProperty $_ | Where-Object {
                ($_.DisplayName -eq "$KLiteMegaCodecPack_EXE_ProductName $($_.DisplayVersion)") -or 
                ($_.DisplayName -like "K-Lite Codec Pack*")
            }
        }
    }
}

# Function to uninstall K-Lite Codec Packs
function Uninstall-KLiteCodecPack($item) {
    $uninstall_Char = ($item.UninstallString -split "  ")
    $LogFile = "$env:systemdrive\temp\$($env:Computername)_$($item.DisplayName -replace $item.DisplayVersion, '').Trim()_Uninstall_$($item.DisplayVersion).txt"
    
    $NeedRemove = if (Test-Path $LogFile) {
        $TimeDiff = (Get-Date) - (Get-Item $LogFile).LastWriteTime
        $TimeDiff.Days -gt 3
    } else { $true }

    if ($NeedRemove) {
        $arguments = " /VERYSILENT /NORESTART /ALLUSERS /FORCECLOSEAPPLICATIONS /LOGCLOSEAPPLICATIONS /LOG=""$LogFile"""
        Start-Process $uninstall_Char[0] -ArgumentList $arguments -Wait
        $script:Force_Install = $true
    }
}

# Function to find Smc.exe
function Get-SmcPath {
    $PossiblePaths = @("${env:ProgramFiles(x86)}\Symantec\Symantec Endpoint Protection\Smc.exe",
                       "${env:ProgramFiles}\Symantec\Symantec Endpoint Protection\Smc.exe")
    foreach ($Path in $PossiblePaths) {
        if (Test-Path $Path) {
            return $Path
        }
    }
    return $null
}

# Function to decode the Symantec password
function Get-DecodedSymantecPassword {
    $encodedPassword = "c3ltYW50ZWM="  # This is "symantec" in Base64
    $decodedBytes = [System.Convert]::FromBase64String($encodedPassword)
    return [System.Text.Encoding]::UTF8.GetString($decodedBytes)
}

# Main script execution
$KLiteMegaCodecPack_EXE = Get-LatestKLiteInstaller
if ($KLiteMegaCodecPack_EXE) {
    $KLiteMegaCodecPack_EXE_ProductName = $KLiteMegaCodecPack_EXE.VersionInfo.ProductName.Trim()
    $KLiteMegaCodecPack_EXE_ProductVersion = $KLiteMegaCodecPack_EXE.VersionInfo.ProductVersion.Trim()
    
    $KLiteCodecPack_installeds = Get-InstalledKLiteCodecPacks
    
    if ($KLiteCodecPack_installeds -or $Force_Install) {
        if ($RemoveFirstPC.Contains($env:Computername) -or ($KLiteCodecPack_installeds.Count -ge 2)) {
            foreach ($item in $KLiteCodecPack_installeds) {
                Uninstall-KLiteCodecPack $item
            }
            
            $Log_Folder_Path = "$Log_Path\$KLiteMegaCodecPack_EXE_ProductName"
            $LogPattern = "$env:Computername`_$KLiteMegaCodecPack_EXE_ProductName`_*.txt"
            if (!(Test-Path -Path $Log_Folder_Path)) { New-Item -ItemType Directory -Path $Log_Folder_Path -Force }
            if (Test-Path -Path "$env:systemdrive\temp") {
                robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
            }
        }

        $KLiteCodecPack_installed = $KLiteCodecPack_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -First 1
        if (([version]$KLiteCodecPack_installed.DisplayVersion -ge [version]$KLiteMegaCodecPack_EXE_ProductVersion) -and (!$Force_Install)) { exit }
        
        $LogName = "$env:Computername`_$KLiteMegaCodecPack_EXE_ProductName`_$KLiteMegaCodecPack_EXE_ProductVersion.txt"
        $arguments = " /VERYSILENT /NORESTART /ALLUSERS /FORCECLOSEAPPLICATIONS /LOGCLOSEAPPLICATIONS /LOG=$env:systemdrive\temp\""$LogName"""
        
        robocopy $KLiteMegaCodecPacks_Path "$env:systemdrive\temp" $KLiteMegaCodecPack_EXE.Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        Unblock-File "$env:systemdrive\temp\$($KLiteMegaCodecPack_EXE.Name)"
        
        $SmcPath = Get-SmcPath
        if ($SmcPath) {
            $symantecPassword = Get-DecodedSymantecPassword
            Start-Process -FilePath $SmcPath -ArgumentList " -p ""$symantecPassword"" -stop" -Wait -WindowStyle Hidden
        }
        
        Start-Process "$env:systemdrive\temp\$($KLiteMegaCodecPack_EXE.Name)" -ArgumentList $arguments -Wait
        
        if ($SmcPath) {
            Start-Process -FilePath $SmcPath -ArgumentList " -p ""$symantecPassword"" -start" -WindowStyle Hidden
            $symantecPassword = $null  # Clear the password variable
        }
        
        $Log_Folder_Path = "$Log_Path\$KLiteMegaCodecPack_EXE_ProductName"
        $LogPattern = "$env:Computername`_$KLiteMegaCodecPack_EXE_ProductName`_*.txt"
        if (!(Test-Path -Path $Log_Folder_Path)) { New-Item -ItemType Directory -Path $Log_Folder_Path -Force }
        if (Test-Path -Path "$env:systemdrive\temp") {
            robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        }
    }
}