# eic.ps1

[CmdletBinding()]
param (
    [Parameter()]
    [switch]$ForceReinstall = $false,

    [Parameter()]
    [switch]$InstallNonStartup = $true,

    [Parameter()]
    [ValidateSet("EicPrint", "EICSignTSR", "Both")]
    [string]$InstallScope = "Both"
)

function Get-MsiInformation
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                    PositionalBinding=$false,
                    ConfirmImpact='Medium')]
    [Alias("gmsi")]
    Param(
        [parameter(Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    HelpMessage = "Provide the path to an MSI")]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$Path,
        [parameter(Mandatory=$false)]
        [ValidateSet( "ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage" )]
        [string[]]$Property = ( "ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage" )
    )

    Begin
    {
        # Do nothing for prep
    }
    Process
    {
        
        ForEach ( $P in $Path )
        {
            if ($pscmdlet.ShouldProcess($P, "Get MSI Properties"))
            {            
                try
                {
                    Write-Verbose -Message "Resolving file information for $P"
                    $MsiFile = Get-Item -Path $P
                    Write-Verbose -Message "Executing on $P"
                    
                    # Read property from MSI database
                    $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
                    $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiFile.FullName, 0))
                    
                    # Build hashtable for retruned objects properties
                    $PSObjectPropHash = [ordered]@{File = $MsiFile.FullName}
                    ForEach ( $Prop in $Property )
                    {
                        Write-Verbose -Message "Enumerating Property: $Prop"
                        $Query = "SELECT Value FROM Property WHERE Property = '$( $Prop )'"
                        $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
                        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
                        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
                        $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
                        # Return the value to the Property Hash
                        $PSObjectPropHash.Add($Prop, $Value)
                    }
                    
                    # Build the Object to Return
                    $Object = @( New-Object -TypeName PSObject -Property $PSObjectPropHash )
                    
                    # Commit database and close view
                    $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
                    $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
                    $MSIDatabase = $null
                    $View = $null
                }
                catch
                {
                    Write-Error -Message $_.Exception.Message
                }
                finally
                {
                    Write-Output -InputObject @( $Object )
                }
            } # End of ShouldProcess If
        } # End For $P in $Path Loop

    }
    End
    {
        # Run garbage collection and release ComObject
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}

$Eic_Path = "\\172.29.205.114\loginscript\Update\eic"
$Log_Path = "\\172.29.205.114\Public\sources\audit\eic"

function Get-NewestMSI {
    param (
        [string]$ProductName
    )

    $MSIs = if ($InstallNonStartup) {
        Get-ChildItem -Path ($Eic_Path + "\$ProductName*非啟動*.msi") 
    } else {
        Get-ChildItem -Path ($Eic_Path + "\$ProductName*.msi") | Where-Object { $_.Name -notlike "*非啟動*" }
    }

    $NewestMSI = $MSIs | ForEach-Object { Get-MsiInformation -Path ($_.FullName) } | Sort-Object -Descending ProductVersion | Select-Object -first 1
    return $NewestMSI
}

function Stop-RelatedProcesses {
    param (
        [string]$ProductName
    )

    $processesToStop = @()
    if ($ProductName -like "*EicPrint*" -or $ProductName -eq "筆硯列印工具") {
        $processesToStop += @("EicPrint", "筆硯列印工具")
    } elseif ($ProductName -like "*EICSignTSR*" -or $ProductName -eq "筆硯簽章工具") {
        $processesToStop += @("EICSignTSR", "筆硯簽章工具")
    }

    foreach ($processName in $processesToStop) {
        $runningProcesses = Get-Process | Where-Object { $_.ProcessName -eq $processName -or $_.MainWindowTitle -eq $processName }
        if ($runningProcesses) {
            Write-Host "Stopping $processName processes..."
            $runningProcesses | ForEach-Object {
                try {
                    $_ | Stop-Process -Force
                    Write-Host "Successfully stopped process: $($_.ProcessName) (ID: $($_.Id))"
                } catch {
                    Write-Host "Failed to stop process: $($_.ProcessName) (ID: $($_.Id)). Error: $($_.Exception.Message)"
                }
            }
            Start-Sleep -Seconds 2  # 等待進程完全停止
        } else {
            Write-Host "No running processes found for $processName"
        }
    }

    # 再次檢查以確保進程已經停止
    foreach ($processName in $processesToStop) {
        $stillRunning = Get-Process | Where-Object { $_.ProcessName -eq $processName -or $_.MainWindowTitle -eq $processName }
        if ($stillRunning) {
            Write-Host "Warning: Some $processName processes are still running after attempt to stop them."
        } else {
            Write-Host "Confirmed: All $processName processes have been stopped."
        }
    }
}

function Uninstall-ExistingProducts {
    param (
        [string]$ProductName
    )

    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $Installed = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $Installed += Get-ItemProperty $Path | Where-Object{$_.DisplayName -eq $ProductName}
        }
    }

    if ($Installed.Count -eq 0) {
        Write-Host "No existing installations of $ProductName found."
        return $false
    }

    Write-Host "Found $($Installed.Count) installation(s) of $ProductName. Uninstalling all..."

    foreach ($product in $Installed) {
        Write-Host "Uninstalling $($product.DisplayName) version $($product.DisplayVersion)..."
        $uninstallString = $product.UninstallString

        if ($uninstallString -like "*msiexec*") {
            $uninstallArgs = $uninstallString -replace "/I", "/x"
            $uninstallArgs = $uninstallArgs -replace "msiexec.exe", ""
            $uninstallArgs = $uninstallArgs.Trim() + " /qn"
            Start-Process "msiexec.exe" -ArgumentList $uninstallArgs -Wait
        } else {
            Start-Process "cmd.exe" -ArgumentList "/c $uninstallString" -Wait
        }
    }

    Write-Host "Uninstallation of all $ProductName versions completed."
    return $true
}

function Install-MSI {
    param (
        [Parameter(Mandatory=$true)]
        [PSObject]$MSIInfo,
        [string]$ProductName
    )

    # 停止相關進程
    Stop-RelatedProcesses -ProductName $ProductName

    $existingInstallation = $false
    if (-not $ForceReinstall) {
        # 檢查是否有現有安裝
        $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
        foreach ($Path in $RegUninstallPaths) {
            if (Test-Path $Path) {
                $installed = Get-ItemProperty $Path | Where-Object{$_.DisplayName -eq $ProductName}
                if ($installed) {
                    $existingInstallation = $true
                    $installedVersion = [version]($installed.DisplayVersion -split '\s+')[0]  # 取第一個部分作為版本號
                    $newVersion = [version]($MSIInfo.ProductVersion -split '\s+')[0]  # 取第一個部分作為版本號
                    if ($newVersion -le $installedVersion) {
                        Write-Host "Existing version ($installedVersion) is newer or same as the version to be installed ($newVersion). Skipping installation."
                        return
                    }
                    break
                }
            }
        }
    }

    # 如果強制重新安裝或有更新版本，則進行卸載
    if ($ForceReinstall -or $existingInstallation) {
        Uninstall-ExistingProducts -ProductName $ProductName
    }

    $LogName = $env:Computername + "_" + $ProductName + "_" + $MSIInfo.ProductVersion + ".txt"
    $MSI_fileName = (Get-Item $MSIInfo.File).Name 

    Write-Host "Installing $ProductName version $($MSIInfo.ProductVersion)..."
    $arguments = "/i $env:systemdrive\temp\$MSI_fileName ALLUSERS=1 /qn /log ""$env:systemdrive\temp\$LogName"""
    robocopy $Eic_Path "$env:systemdrive\temp" (Get-ChildItem -Path $MSIInfo.File).Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    Start-Process "msiexec" -ArgumentList $arguments -Wait

    if(Test-Path -Path "$env:systemdrive\temp\$LogName"){
        robocopy "$env:systemdrive\temp" $Log_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    }

    Write-Host "Installation of $ProductName version $($MSIInfo.ProductVersion) completed."

    # 檢查是否安裝了啟動版本，如果是，則立即執行程式
    if (-not $InstallNonStartup) {
        $exePath = ""
        if ($ProductName -like "*EicPrint*") {
            $exePath = "$env:systemdrive\eic\EicPrint\EicPrint.exe"
        } elseif ($ProductName -like "*EICSignTSR*") {
            $exePath = "$env:systemdrive\eic\EICSignTSR\EicSignTSR.exe"
        }

        if (Test-Path $exePath) {
            Write-Host "Starting $ProductName..."
            Start-Process $exePath
        } else {
            Write-Host "Warning: Executable for $ProductName not found at $exePath"
        }
    }
}

# Main execution logic
if ($InstallScope -eq "EicPrint" -or $InstallScope -eq "Both") {
    $EicPrint_Newest_MSI = Get-NewestMSI -ProductName "EicPrint"
    if ($EicPrint_Newest_MSI) {
        Install-MSI -MSIInfo $EicPrint_Newest_MSI -ProductName $EicPrint_Newest_MSI.ProductName
    }
}

if ($InstallScope -eq "EICSignTSR" -or $InstallScope -eq "Both") {
    $EICSignTSR_Newest_MSI = Get-NewestMSI -ProductName "EICSignTSR"
    if ($EICSignTSR_Newest_MSI) {
        Install-MSI -MSIInfo $EICSignTSR_Newest_MSI -ProductName $EICSignTSR_Newest_MSI.ProductName
    }
}