# Configuration
$JavaExcludePC = @("TND-GASE-061")
$Javas_EXE_Path = "\\172.29.205.114\loginscript\Update\Java"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$Global:PatternJava32 = "Java ([0-9]|[0-9][0-9]) Update ([0-9]|[0-9][0-9]|[0-9][0-9][0-9])$"
$Global:PatternJava64 = "Java ([0-9]|[0-9][0-9]) Update ([0-9]|[0-9][0-9]|[0-9][0-9][0-9]) \(64-bit\)$" 
$RegUninstallPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
$FontClient_EXE_Path = "$env:systemdrive\CMEX_FontClient\FontClient.exe"
$FontClient_AutoUpdate_EXE_Path = "$env:systemdrive\CMEX_FontClient\AutoUpdate.exe"

# Helper Functions
function Get-JavaInstalls {
    param (
        [string]$Pattern
    )
    $installedJava = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $installedJava += Get-ItemProperty $Path | Where-Object{$_.DisplayName -match $Pattern}
        }
    }
    return $installedJava | Sort-Object -Property Version -Descending
}

function Install-Java {
    param (
        [string]$ExePath,
        [string]$Arguments
    )
    try {
        Stop-Process -Name "FontClient" -Force -ErrorAction SilentlyContinue
        Start-Process $ExePath -ArgumentList $Arguments -Wait
        return $true
    }
    catch {
        Write-Error "Failed to install Java: $_"
        return $false
    }
}

function Uninstall-OldJava {
    param (
        [array]$JavaInstalls,
        [string]$LatestVersion,
        [string]$ProductVersion
    )
    foreach($install in $JavaInstalls) {
        if([version]$install.DisplayVersion -eq [version]$LatestVersion) { continue }
        $uninstall = ($install.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()
        $LogFile = "$env:systemdrive\temp\${env:Computername}_$($install.DisplayName)_${ProductVersion}_Remove.txt"
        Stop-Process -Name "FontClient" -Force -ErrorAction SilentlyContinue
        Start-Process "msiexec.exe" -ArgumentList "/X $uninstall /quiet /l*vx ""$LogFile""" -Wait
    }
}

# Main execution
if($JavaExcludePC.Contains($env:Computername)) { exit }

$Javas_EXE_32 = Get-ChildItem -Path ($Javas_EXE_Path+"\*-i586.exe") | Sort-Object -Property VersionInfo -Descending | Select-Object -First 1
$Javas_EXE_64 = Get-ChildItem -Path ($Javas_EXE_Path+"\*-x64.exe") | Sort-Object -Property VersionInfo -Descending | Select-Object -First 1

# 32-bit Java update
if($Javas_EXE_32.FullName) {
    $Java_32_installeds = Get-JavaInstalls -Pattern $Global:PatternJava32
    $Java_32_Lastest_installed = $Java_32_installeds | Select-Object -First 1
    $LogName = "${env:Computername}_$($Javas_EXE_32.VersionInfo.ProductName)_$($Javas_EXE_32.VersionInfo.ProductVersion).txt"
    $arguments = "/s AUTO_UPDATE=0 REBOOT=0 /LV* ""$env:systemdrive\temp\$LogName"""

    if(-not $Java_32_Lastest_installed -or [version]$Java_32_Lastest_installed.DisplayVersion -lt [version]$Javas_EXE_32.VersionInfo.ProductVersion) {
        $installSuccess = Install-Java -ExePath "$env:systemdrive\temp\$($Javas_EXE_32.Name)" -Arguments $arguments
        if($installSuccess) {
            $Java_32_installeds = Get-JavaInstalls -Pattern $Global:PatternJava32
            $Java_32_Lastest_installed = $Java_32_installeds | Select-Object -First 1
        }
    }
}

# 64-bit Java update (similar structure to 32-bit update)
if($Javas_EXE_64.FullName) {
    $Java_64_installeds = Get-JavaInstalls -Pattern $Global:PatternJava64
    $Java_64_Lastest_installed = $Java_64_installeds | Select-Object -First 1
    $LogName = "${env:Computername}_$($Javas_EXE_64.VersionInfo.ProductName) (x64)_$($Javas_EXE_64.VersionInfo.ProductVersion).txt"
    $arguments = "/s AUTO_UPDATE=0 REBOOT=0 /l*vx ""$env:systemdrive\temp\$LogName"""

    if(-not $Java_64_Lastest_installed -or [version]$Java_64_Lastest_installed.DisplayVersion -lt [version]$Javas_EXE_64.VersionInfo.ProductVersion) {
        $installSuccess = Install-Java -ExePath "$env:systemdrive\temp\$($Javas_EXE_64.Name)" -Arguments $arguments
        if($installSuccess) {
            $Java_64_installeds = Get-JavaInstalls -Pattern $Global:PatternJava64
            $Java_64_Lastest_installed = $Java_64_installeds | Select-Object -First 1
        }
    }
}

# Uninstall old Java versions
if($Java_32_Lastest_installed) {
    Uninstall-OldJava -JavaInstalls $Java_32_installeds -LatestVersion $Java_32_Lastest_installed.DisplayVersion -ProductVersion $Javas_EXE_32.VersionInfo.ProductVersion
}
if($Java_64_Lastest_installed) {
    Uninstall-OldJava -JavaInstalls $Java_64_installeds -LatestVersion $Java_64_Lastest_installed.DisplayVersion -ProductVersion $Javas_EXE_64.VersionInfo.ProductVersion
}

# Log file management
$Log_Folder_Path = Join-Path $Log_Path "Java"
if(!(Test-Path -Path $Log_Folder_Path)) { New-Item -ItemType Directory -Path $Log_Folder_Path -Force }
$LogPattern = "${env:Computername}_Java*.txt"
if(Test-Path -Path "$env:systemdrive\temp") {
    robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS" | Out-Null
}

# Start FontClient if not running
if(-not (Get-Process | Where-Object {$_.MainModule.FileName -eq $FontClient_EXE_Path})) { 
    Start-Process -FilePath $FontClient_EXE_Path -ArgumentList "-gui"
    Start-Process -FilePath $FontClient_AutoUpdate_EXE_Path
}