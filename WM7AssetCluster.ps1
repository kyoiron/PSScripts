# 特殊電腦名稱設定
$SpecialComputerName = "TND-STOF-112"

# 參數設定
$WM7Asset_Path = "$env:SystemDrive\WM7Asset"
$RequiredFiles = @("WM7Asset.bat", "WM7Assetreport.xml", "WM7LiteGreen.exe")
$ComputerName = $env:COMPUTERNAME
$TempPath = "$env:SystemDrive\temp"
$LogFile = "$TempPath\WM7Asset_Install_Log_$ComputerName.txt"
$NetworkLogPath = "\\172.29.205.114\Public\sources\audit\WM7AssetCluster"
$WM7AssetCluster_EXE_Path = "\\172.29.205.114\loginscript\Update\WM7AssetCluster"
$VersionJsonPath = "$WM7AssetCluster_EXE_Path\WM7AssetVersion.json"
$DownloadUrl = "http://download.moj/files/VANS/VANS_WM7_20240723.exe"

# 確保臨時目錄存在
if (-not (Test-Path $TempPath)) {
    New-Item -ItemType Directory -Path $TempPath | Out-Null
}

# 函數定義
function Write-Log {
    param([string]$Message)    
    $TimeStamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $LogMessage = "$TimeStamp - $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
}

function Initialize-Log {
    if (Test-Path $LogFile) {
        Remove-Item $LogFile -Force
    }
    Write-Log "開始新的 WM7Asset 安裝檢查 (電腦名稱: $ComputerName)"
}

function Get-InstalledVersion {
    $exePath = "$WM7Asset_Path\WM7LiteGreen.exe"
    if (Test-Path $exePath) {
        return [version](Get-Item -Path $exePath).VersionInfo.ProductVersion
    }
    return [version]"0.0.0.0"
}

function Update-VersionJson {
    param([string]$Version)
    $versionInfo = @{
        "Version" = $Version
        "LastUpdated" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }
    $versionInfo | ConvertTo-Json | Set-Content -Path $VersionJsonPath
    Write-Log "已更新版本資訊到 JSON 檔案：$VersionJsonPath"
}

function Get-LatestVersionFromJson {
    if (Test-Path $VersionJsonPath) {
        $versionInfo = Get-Content $VersionJsonPath | ConvertFrom-Json
        return [version]$versionInfo.Version
    }
    return [version]"0.0.0.0"
}

function Get-7ZipPath {
    $possiblePaths = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe"
    )
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return $path
        }
    }
    return $null
}

function Extract-WithSevenZip {
    param(
        [string]$SourceFile,
        [string]$DestinationPath
    )
    $7zipPath = Get-7ZipPath
    if ($7zipPath -eq $null) {
        throw "無法找到 7-Zip，請確保已安裝 7-Zip"
    }
    Write-Log "開始使用 7-Zip 解壓縮文件"
    try {
        $output = & $7zipPath x "$SourceFile" "-o$DestinationPath" -y 2>&1
        $output | ForEach-Object {
            Write-Log "7-Zip: $_"
        }
        if ($LASTEXITCODE -ne 0) {
            throw "7-Zip 解壓縮失敗，退出碼: $LASTEXITCODE"
        }
    }
    catch {
        Write-Log "7-Zip 解壓縮過程中發生錯誤: $($_.Exception.Message)"
        throw
    }
    Write-Log "7-Zip 解壓縮完成"
}

function Process-SpecialComputer {
    Write-Log "正在處理 $SpecialComputerName 特殊邏輯"
    $tempPath = "$env:TEMP\WM7AssetCluster"
    $exePath = "$tempPath\WM7AssetCluster.exe"

    try {
        # 下載檔案
        Write-Log "開始下載 WM7AssetCluster.exe"
        New-Item -ItemType Directory -Force -Path $tempPath | Out-Null
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $exePath
        Write-Log "下載完成"

        # 解壓縮檔案
        Write-Log "開始解壓縮 WM7AssetCluster.exe"
        $7zipPath = Get-7ZipPath
        if ($7zipPath) {
            Write-Log "使用 7-Zip 進行解壓縮 (路徑: $7zipPath)"
            Extract-WithSevenZip -SourceFile $exePath -DestinationPath $tempPath
        } else {
            throw "未安裝 7-Zip，無法進行解壓縮操作"
        }

        # 檢查版本並更新 JSON
        $wm7LiteGreenPath = "$tempPath\WM7LiteGreen.exe"
        if (Test-Path $wm7LiteGreenPath) {
            $latestVersion = (Get-Item $wm7LiteGreenPath).VersionInfo.ProductVersion
            Write-Log "最新的 WM7LiteGreen.exe 版本: $latestVersion"
            Update-VersionJson $latestVersion
        } else {
            throw "無法找到 WM7LiteGreen.exe"
        }

        # 檢查是否需要更新本機軟件
        $installedVersion = Get-InstalledVersion
        if ([version]$latestVersion -gt $installedVersion) {
            Write-Log "本機軟體需要更新，當前版本: $installedVersion，最新版本: $latestVersion"
            return $true
        } else {
            Write-Log "本機軟體已是最新版本，無需更新"
            return $false
        }
    }
    catch {
        Write-Log "處理 $SpecialComputerName 時發生錯誤: $($_.Exception.Message)"
        throw
    }
    finally {
        # 清理臨時檔案
        if (Test-Path $tempPath) {
            Remove-Item $tempPath -Recurse -Force
            Write-Log "已清理臨時檔案"
        }
    }
}

function Get-LatestInstallerFromPath {
    Write-Log "開始在 $WM7AssetCluster_EXE_Path 中搜索最新的安裝檔"
    
    if (-not (Test-Path $WM7Asset_Path)) {
        Write-Log "警告：$WM7Asset_Path 目錄不存在"
        return $null
    }
    
    $latestExe = Get-ChildItem -Path $WM7AssetCluster_EXE_Path -Filter "*.exe" -ErrorAction SilentlyContinue | 
                 Sort-Object LastWriteTime -Descending | 
                 Select-Object -First 1

    if ($latestExe) {
        Write-Log "找到最新的安裝檔：$($latestExe.Name)"
        return $latestExe.FullName
    } else {
        Write-Log "在 $WM7Asset_Path 中沒有找到 .exe 文件"
        return $null
    }
}

function Test-InstallationRequired {
    if (-not (Test-Path $WM7Asset_Path)) {
        Write-Log "WM7Asset 目錄不存在，需要安裝"
        return $true
    }

    $filesExist = $RequiredFiles | ForEach-Object { Test-Path "$WM7Asset_Path\$_" }
    if ($filesExist -contains $false) {
        Write-Log "一些必要文件不存在，需要安裝"
        return $true 
    }

    $installedVersion = Get-InstalledVersion
    $latestVersion = Get-LatestVersionFromJson
    if ($installedVersion -lt $latestVersion) {
        Write-Log "當前版本 ($installedVersion) 低於最新版本 ($latestVersion)，需要更新"
        return $true
    }

    Write-Log "所有檢查通過，不需要安裝或更新"
    return $false
}

function Verify-Installation {
    $installationSuccess = $true
    $exePath = "$env:SystemDrive\WM7Asset\WM7LiteGreen.exe"

    # 檢查 WM7LiteGreen.exe 版本
    if (Test-Path $exePath) {
        $installedVersion = (Get-Item $exePath).VersionInfo.ProductVersion
        $latestVersion = Get-LatestVersionFromJson
        if ([version]$installedVersion -eq [version]$latestVersion) {
            Write-Log "WM7LiteGreen.exe 版本正確：$installedVersion"
        } else {
            Write-Log "WM7LiteGreen.exe 版本不正確。已安裝版本：$installedVersion，最新版本：$latestVersion"
            $installationSuccess = $false
        }
    } else {
        Write-Log "找不到 WM7LiteGreen.exe"
        $installationSuccess = $false
    }

    # 檢查工作排程器
    $scheduledTask = Get-ScheduledTask -TaskName "WM7AssetReport" -ErrorAction SilentlyContinue
    if ($scheduledTask) {
        Write-Log "找到 WM7AssetReport 排程任務"
    } else {
        Write-Log "找不到 WM7AssetReport 排程任務"
        $installationSuccess = $false
    }

    if ($installationSuccess) {
        Write-Log "安裝驗證成功：WM7LiteGreen.exe 版本正確且 WM7AssetReport 排程任務存在"
    } else {
        Write-Log "安裝驗證失敗：WM7LiteGreen.exe 版本不正確或 WM7AssetReport 排程任務不存在"
    }

    return $installationSuccess
}

function Install-WM7Asset {
    $tempInstallerPath = "$env:SystemDrive\temp\WM7AssetInstaller.exe"

    try {
        # 記錄安裝前的版本
        $currentVersion = Get-InstalledVersion
        Write-Log "安裝前的 WM7LiteGreen.exe 版本：$currentVersion"

        if ($ComputerName -eq $SpecialComputerName) {
            # 特殊電腦的安裝邏輯
            Write-Log "特殊電腦：開始下載 WM7AssetCluster.exe"
            Invoke-WebRequest -Uri $DownloadUrl -OutFile $tempInstallerPath
            
            # 對於特殊電腦，我們需要解壓縮檔案來獲取最新版本
            $tempExtractPath = "$env:TEMP\WM7AssetExtract"
            Extract-WithSevenZip -SourceFile $tempInstallerPath -DestinationPath $tempExtractPath
            $latestVersion = (Get-Item "$tempExtractPath\WM7LiteGreen.exe").VersionInfo.ProductVersion
            Remove-Item $tempExtractPath -Recurse -Force
        } else {
            # 普通電腦的安裝邏輯
            $latestInstallerPath = Get-LatestInstallerFromPath
            if (-not $latestInstallerPath) {
                throw "無法找到安裝檔"
            }
            Write-Log "普通電腦：複製最新安裝檔到臨時目錄"
            Copy-Item -Path $latestInstallerPath -Destination $tempInstallerPath -Force
            
            # 獲取最新版本信息
            $latestVersion = Get-LatestVersionFromJson
        }

        # 記錄最新版本
        Write-Log "最新可用的 WM7LiteGreen.exe 版本：$latestVersion"

        # 解除文件封鎖
        Write-Log "開始解除安裝檔案封鎖"
        try {
            Unblock-File -Path $tempInstallerPath
            Write-Log "安裝檔案封鎖解除成功"
        }
        catch {
            Write-Log "解除安裝檔案封鎖時發生錯誤：$($_.Exception.Message)"
            # 即使解除封鎖失敗，我們也繼續安裝過程
        }
        # 檢查並強制停止正在運行的安裝程序
        Write-Log "檢查是否有正在運行的 WM7AssetInstaller.exe 進程"
        $runningInstaller = Get-Process | Where-Object { $_.Path -eq $tempInstallerPath }
        if ($runningInstaller) {
            Write-Log "發現正在運行的 WM7AssetInstaller.exe 進程，嘗試強制停止"
            $runningInstaller | ForEach-Object { 
                $_ | Stop-Process -Force
                Write-Log "已強制停止進程 ID: $($_.Id)"
            }
            Start-Sleep -Seconds 2  # 等待進程完全停止
        } else {
            Write-Log "沒有發現正在運行的 WM7AssetInstaller.exe 進程"
        }

        # 刪除排程任務
        Write-Log "嘗試刪除 WM7AssetReport 排程任務"
        try {
            Unregister-ScheduledTask -TaskName "WM7AssetReport" -Confirm:$false
            Write-Log "成功刪除 WM7AssetReport 排程任務"
        } catch {
            Write-Log "刪除 WM7AssetReport 排程任務時發生錯誤：$($_.Exception.Message)"
        }

        # 刪除 WM7Asset 目錄
        $wm7AssetPath = "$env:SystemDrive\WM7Asset"
        Write-Log "嘗試刪除 WM7Asset 目錄"
        if (Test-Path $wm7AssetPath) {
            try {
                Remove-Item -Path $wm7AssetPath -Recurse -Force
                Write-Log "成功刪除 WM7Asset 目錄"
            } catch {
                Write-Log "刪除 WM7Asset 目錄時發生錯誤：$($_.Exception.Message)"
            }
        } else {
            Write-Log "WM7Asset 目錄不存在，無需刪除"
        }
        Write-Log "開始執行安裝程序"
        Start-Process -FilePath $tempInstallerPath -Wait
        Write-Log "安裝程序執行完畢"

        # 驗證安裝
        if (Verify-Installation) {
            Write-Log "安裝成功並通過驗證"
            $newVersion = Get-InstalledVersion
            Write-Log "安裝後的 WM7LiteGreen.exe 版本：$newVersion"
        } else {
            throw "安裝驗證失敗"
        }
    }
    catch {
        Write-Log "安裝過程中發生錯誤：$($_.Exception.Message)"
        throw
    }
    finally {
        if (Test-Path $tempInstallerPath) {
            Remove-Item $tempInstallerPath -Force
            Write-Log "已移除臨時安裝檔"
        }
    }
}

function Copy-LogToNetwork {
    $NetworkLogFile = Join-Path $NetworkLogPath "WM7Asset_Install_Log_$ComputerName.txt"
    try {
        Copy-Item -Path $LogFile -Destination $NetworkLogFile -Force
        Write-Log "已成功將日誌複製到網路位置: $NetworkLogFile"
    }
    catch {
        Write-Log "複製日誌到網路位置時發生錯誤：$($_.Exception.Message)"
    }
}

# 主要執行邏輯
Initialize-Log

try {
    $needUpdate = $false
    
    if ($ComputerName -eq $SpecialComputerName) {
        Write-Log "這是特殊電腦 $SpecialComputerName，執行特殊處理邏輯"
        $needUpdate = Process-SpecialComputer
    } else {
        Write-Log "這是普通電腦，執行標準檢查邏輯"
        $needUpdate = Test-InstallationRequired
    }

    if ($needUpdate) {
        Write-Log "需要進行安裝或更新，開始安裝程序"
        if (-not (Test-Path $WM7Asset_Path) -and $ComputerName -ne $SpecialComputerName) {
            Write-Log "WM7Asset 目錄不存在，嘗試創建"
            New-Item -ItemType Directory -Path $WM7Asset_Path -Force | Out-Null
            Write-Log "WM7Asset 目錄已創建"
        }
        Install-WM7Asset
        Write-Log "安裝程序完成"
    } else {
        Write-Log "WM7Asset 已是最新版本，無需進行任何操作"
        $currentVersion = Get-InstalledVersion
        Write-Log "當前安裝的 WM7LiteGreen.exe 版本：$currentVersion"
    }
}
catch {
    Write-Log "執行過程中發生錯誤: $($_.Exception.Message)"
}
finally {
    Write-Log "WM7Asset 安裝檢查程序執行完畢"
    Copy-LogToNetwork
}