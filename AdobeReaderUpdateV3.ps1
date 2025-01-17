# 函數：取得 MSI 資訊
function Get-MsiInformation {
    [CmdletBinding(SupportsShouldProcess=$true, PositionalBinding=$false, ConfirmImpact='Medium')]
    [Alias("gmsi")]
    Param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="提供 MSI 檔案的路徑")]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$Path,
        [parameter(Mandatory=$false)]
        [ValidateSet("ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage")]
        [string[]]$Property = ("ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage")
    )

    Process {
        foreach ($P in $Path) {
            if ($pscmdlet.ShouldProcess($P, "取得 MSI 屬性")) {
                try {
                    $MsiFile = Get-Item -Path $P
                    $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
                    $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiFile.FullName, 0))
                    
                    $PSObjectPropHash = [ordered]@{File = $MsiFile.FullName}
                    foreach ($Prop in $Property) {
                        $Query = "SELECT Value FROM Property WHERE Property = '$Prop'"
                        $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
                        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
                        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
                        $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
                        $PSObjectPropHash.Add($Prop, $Value)
                    }
                    
                    $Object = New-Object -TypeName PSObject -Property $PSObjectPropHash
                    
                    $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
                    $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
                    $MSIDatabase = $null
                    $View = $null
                }
                catch {
                    Write-Error -Message $_.Exception.Message
                }
                finally {
                    Write-Output -InputObject $Object
                }
            }
        }
    }
    
    End {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}

# 設定
$config = @{
    LogPath = "\\172.29.205.114\Public\sources\audit"
    AdobeReaderFolder = "\\172.29.205.114\loginscript\Update\AdobeReader"
    AdobeReader64Folder = "\\172.29.205.114\loginscript\Update\AdobeAcrobatDC(64-bit)"
    InstallIfNotInstalled = $true
    MsiexecProcessName = "msiexec.exe"
    TempFolder = "$env:systemdrive\temp"
    LogFileName = "$env:COMPUTERNAME`_Adobe_Reader_Update.log"
}

# 函數：記錄訊息
function Write-Log {
    param (
        [string]$Message
    )
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Add-Content -Path $script:LogFile -Value $logMessage
    Write-Host $logMessage
}


# 函數：取得最新的 MSP 檔案
function Get-LatestMSP($folder) {
    return Get-ChildItem -Path "$folder\*.msp" | Sort-Object -Property Name -Descending | Select-Object -First 1
}

# 函數：取得已安裝的 Adobe Reader
function Get-InstalledAdobeReader($productCode) {
    $paths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    return Get-ItemProperty $paths | Where-Object { $_.PSChildName -eq $productCode } | 
           Sort-Object -Property DisplayVersion -Descending | Select-Object -First 1
}

# 函數：更新 setup.ini 檔案
function Update-SetupIni($tempFolder, $mspName) {
    $setupPath = "$tempFolder\setup.ini"
    $content = Get-Content -Path $setupPath
    $pattern = 'PATCH=AcroRdrDCUpd|PATCH=AcroRdrDCx64Upd'
    ($content -replace $pattern, "PATCH=$mspName") | Set-Content $setupPath
    Write-Log "已更新 setup.ini 檔案，設置 PATCH=$mspName"
}

# 函數：等待 msiexec 進程結束
function Wait-Msiexec {
    do {
        $msiexecProcess = Get-Process -Name $config.MsiexecProcessName -ErrorAction SilentlyContinue
        if ($msiexecProcess) { 
            Start-Sleep -Seconds 1
            Write-Log "等待 msiexec 進程結束..."
        }
    } while ($msiexecProcess)
    Write-Log "msiexec 進程已結束"
}

# 函數：安裝或更新 Adobe Reader
function Install-AdobeReader($msiPath, $mspPath, $logName) {
    $msiLogPath = "$($config.TempFolder)\$logName"
    $arguments = "/i `"$msiPath`" /update `"$mspPath`" /quiet /norestart /l*vx `"$msiLogPath`""
    Unblock-File $mspPath
    Wait-Msiexec
    Write-Log "開始安裝/更新 Adobe Reader: $arguments"
    Start-Process "msiexec" -ArgumentList $arguments -Wait
    Write-Log "Adobe Reader 安裝/更新完成"
    
    # 檢查 MSI 紀錄檔是否存在
    if (Test-Path $msiLogPath) {
        Write-Log "MSI 安裝/更新紀錄檔已生成: $msiLogPath"
        
        # 複製 MSI 紀錄檔到網路位置
        $networkLogFolder = Join-Path $config.LogPath "Adobe Acrobat Reader DC - Chinese Traditional"
        if (!(Test-Path $networkLogFolder)) { 
            New-Item -ItemType Directory -Path $networkLogFolder -Force
            Write-Log "建立網路紀錄檔資料夾: $networkLogFolder"
        }
        Copy-Item $msiLogPath $networkLogFolder -Force
        Write-Log "已複製 MSI 安裝/更新紀錄檔到網路位置: $networkLogFolder\$logName"
    } else {
        Write-Log "警告：未找到 MSI 安裝/更新紀錄檔: $msiLogPath"
    }
}

# 函數：解析 MSP 檔案名稱以取得版本
function Get-VersionFromMSP($mspName) {
    if ($mspName -match '(\d{2})(\d{3})(\d{5})') {
        $major = [int]$matches[1]
        $minor = [int]$matches[2]
        $build = [int]$matches[3]
        return "$major.$minor.$build"
    }
    throw "無法從 MSP 檔案名稱解析版本: $mspName"
}

# 函數：清理暫存檔案
function Clear-TempFiles {
    param (
        [string]$MspName
    )
    $filesToRemove = @("AcroRead.msi", "abcpy.ini", "Data1.cab", "setup.exe", "setup.ini", $MspName)
    foreach ($file in $filesToRemove) {
        $filePath = Join-Path $config.TempFolder $file
        if (Test-Path $filePath) {
            Remove-Item $filePath -Force
            Write-Log "已刪除暫存檔案: $filePath"
        }
    }
}

# 主要腳本
try {
    $msp = Get-LatestMSP $config.AdobeReaderFolder
    $msp64 = Get-LatestMSP $config.AdobeReader64Folder
    $acroReadMsi = Get-MsiInformation -Path "$($config.AdobeReaderFolder)\AcroRead.msi"
    $acroPro64Msi = Get-MsiInformation -Path "$($config.AdobeReader64Folder)\AcroPro.msi"

    $mspVersion = Get-VersionFromMSP ($msp.BaseName -replace 'AcroRdrDCUpd','')
    $msp64Version = Get-VersionFromMSP ($msp64.BaseName -replace 'AcroRdrDCx64Upd','')

    $installedReader = Get-InstalledAdobeReader $acroReadMsi.ProductCode
    $installedReader64 = Get-InstalledAdobeReader $acroPro64Msi.ProductCode

    # 確定是否為 64 位元版本
    $is64Bit = $false
    if ($installedReader64) {
        $is64Bit = $true
    }

   # 更新日誌檔案名稱
    $archSuffix = if ($is64Bit) { "_x64" } else { "" }
    $config.LogFileName = "$env:COMPUTERNAME`_Adobe_Reader$archSuffix`_Update.log"
    $script:LogFile = Join-Path $config.TempFolder $config.LogFileName

    # 清空現有的紀錄檔
    Clear-Content -Path $script:LogFile -Force

    Write-Log "開始執行 Adobe Reader 安裝/更新腳本"
    Write-Log "取得到最新的 MSP 檔案: $($msp.Name) (版本 $mspVersion), $($msp64.Name) (版本 $msp64Version)"

    Write-Log "檢查已安裝的 Adobe Reader 版本"
    if ($installedReader) { 
        Write-Log "已安裝 32 位元版本: $($installedReader.DisplayVersion)" 
        Write-Log "最新可用 32 位元更新版本: $mspVersion"
    }
    if ($installedReader64) { 
        Write-Log "已安裝 64 位元版本: $($installedReader64.DisplayVersion)" 
        Write-Log "最新可用 64 位元更新版本: $msp64Version"
    }

    # 如果沒有安裝 Reader 且不需要安裝，則退出
    if (!$installedReader -and !$installedReader64 -and !$config.InstallIfNotInstalled) {
        Write-Log "未安裝 Adobe Reader，且不需要安裝。腳本退出。"
        exit
    }

    $updatePerformed = $false

    if ($installedReader) {
        # 32 位元版本已安裝，檢查更新
        Write-Log "檢查 32 位元版本更新。本機安裝版本: $($installedReader.DisplayVersion), 可用更新版本: $mspVersion"
        if ([version]$installedReader.DisplayVersion -ge [version]$mspVersion) {
            Write-Log "已安裝的 32 位元版本是最新的。無需更新。"
        } else {
            Write-Log "準備更新 32 位元版本"
            $logName = "$env:COMPUTERNAME`_$($acroReadMsi.ProductName)`_$mspVersion.txt"
            $msiPath = if (Test-Path "$($installedReader.InstallSource)AcroRead.msi") { 
                "$($installedReader.InstallSource)AcroRead.msi" 
            } else {
                Copy-Item "$($config.AdobeReaderFolder)\*" $config.TempFolder -Include @("AcroRead.msi", "abcpy.ini", "Data1.cab", "setup.exe", "setup.ini")
                Update-SetupIni $config.TempFolder $msp.Name
                "$($config.TempFolder)\AcroRead.msi"
            }
            Install-AdobeReader $msiPath $msp.FullName $logName
            $updatePerformed = $true
        }
    } elseif ($installedReader64) {
        # 64 位元版本已安裝，檢查更新
        Write-Log "檢查 64 位元版本更新。本機安裝版本: $($installedReader64.DisplayVersion), 可用更新版本: $msp64Version"
        if ([version]$installedReader64.DisplayVersion -ge [version]$msp64Version) {
            Write-Log "已安裝的 64 位元版本是最新的。無需更新。"
        } else {
            Write-Log "準備更新 64 位元版本"
            $logName = "$env:COMPUTERNAME`_$($acroPro64Msi.ProductName)`_$msp64Version.txt"
            $msiPath = if (Test-Path "$($installedReader64.InstallSource)AcroPro.msi") { 
                "$($installedReader64.InstallSource)AcroPro.msi" 
            } else {
                Copy-Item "$($config.AdobeReader64Folder)\*" $config.TempFolder -Include @("AcroPro.msi", "abcpy.ini", "Core.cab", "Languages.cab", "setup.exe", "setup.ini", "WindowsInstaller-KB893803-v2-x86.exe", "Transforms", "VCRT_x64")
                Update-SetupIni $config.TempFolder $msp64.Name
                "$($config.TempFolder)\AcroPro.msi"
            }
            Install-AdobeReader $msiPath $msp64.FullName $logName
            $updatePerformed = $true
        }
     } else {
        # 未安裝 Reader，進行新安裝
        Write-Log "未檢測到已安裝的 Adobe Reader，準備進行新安裝"
        $logName = "$env:COMPUTERNAME`_$($acroReadMsi.ProductName)`_$mspVersion.txt"
        
        Copy-Item "$($config.AdobeReaderFolder)\*" $config.TempFolder -Include @("AcroRead.msi", "abcpy.ini", "Data1.cab", "setup.exe", "setup.ini")
        Update-SetupIni $config.TempFolder $msp.Name
        
        Install-AdobeReader "$($config.TempFolder)\AcroRead.msi" $msp.FullName $logName
        $updatePerformed = $true
    }

    if ($updatePerformed) {
        Write-Log "已執行更新/安裝過程"
    } else {
        Write-Log "未執行更新過程，本機安裝版本已是最新"
    }
}
catch {
    Write-Log "發生錯誤: $_"
    throw
}
finally {
    if ($updatePerformed) {
        Clear-TempFiles -MspName $msp.Name
    }
    Write-Log "腳本執行完成"
    # 複製紀錄檔到網路位置
    $networkLogPath = Join-Path $config.LogPath $acroReadMsi.ProductName
    if (!(Test-Path $networkLogPath)) { 
        New-Item -ItemType Directory -Path $networkLogPath -Force 
        Write-Log "建立網路紀錄檔資料夾: $networkLogPath"
    }
    Copy-Item $script:LogFile $networkLogPath -Force
    Write-Log "已複製執行紀錄檔到網路位置: $networkLogPath"
}
