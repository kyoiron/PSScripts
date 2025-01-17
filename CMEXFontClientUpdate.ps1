# 定義全域變數
$ErrorActionPreference = "Stop"
$CMEXFontClient_Path = "\\172.29.205.114\loginscript\Update\CMEX-FontClient"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$Force_Install = $false
$UninstallOlderThanDays = 3
$AutoUpdateExePath = "$env:systemdrive\CMEX_FontClient\AutoUpdate.exe"

# 設置本機暫存日誌目錄
$localTempPath = "$env:systemdrive\temp"
if (-not (Test-Path $localTempPath)) { 
    New-Item -ItemType Directory $localTempPath -Force | Out-Null 
}

# 啟動記錄（先儲存在本機）
$localLogFile = Join-Path $localTempPath "$env:COMPUTERNAME`_CMEXFontClient_LOG_.txt"
Start-Transcript -Path $localLogFile -Force

# 函數：取得 Symantec Endpoint Protection 路徑
function Get-SymantecPath {
    $paths = @("${env:ProgramFiles(x86)}\Symantec\Symantec Endpoint Protection\Smc.exe",
               "${env:ProgramFiles}\Symantec\Symantec Endpoint Protection\Smc.exe")
    foreach ($path in $paths) {
        if (Test-Path $path) { return $path }
    }
    Write-Warning "找不到 Symantec Endpoint Protection。將跳過 Symantec 相關操作。"
    return $null
}

# 函數：解碼固定的 Symantec 密碼
function Get-DecodedSymantecPassword {
    $encodedPassword = "c3ltYW50ZWM="  # Base64 編碼
    $decodedBytes = [System.Convert]::FromBase64String($encodedPassword)
    return [System.Text.Encoding]::UTF8.GetString($decodedBytes)
}

# 函數：取得最新的 CMEX Font Client 安裝檔
function Get-LatestCMEXFontClientInstaller {
    return Get-ChildItem -Path "$CMEXFontClient_Path\*.exe" | 
           Where-Object { $_.VersionInfo.ProductName.Trim() -like "中推會用戶端更新工具*" } | 
           Sort-Object -Property VersionInfo -Descending | 
           Select-Object -First 1
}

# 函數：取得已安裝的 CMEX Font Client 版本
function Get-InstalledCMEXFontClients {
    $regPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                  'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $installedClients = @()
    foreach ($path in $regPaths) {
        if (Test-Path $path) {
            $installedClients += Get-ItemProperty $path | 
                                 Where-Object { $_.DisplayName -like "中推會用戶端更新工具*" }
        }
    }
    return $installedClients
}

# 函數：解除安裝舊版本
function Uninstall-OldVersions {
    param ($installedClients)
    foreach ($client in $installedClients) {
        $uninstallString = $client.UninstallString -replace '^"?([^"]+)"?.*$', '$1'
        $logFile = "$localTempPath\${env:COMPUTERNAME}_$($client.DisplayName.Replace($client.DisplayVersion, '').Trim())_解除安裝_$($client.DisplayVersion).txt"
        
        $arguments = "/VERYSILENT /NORESTART /ALLUSERS /FORCECLOSEAPPLICATIONS /LOGCLOSEAPPLICATIONS /LOG=`"$logFile`""
        Start-Process $uninstallString -ArgumentList $arguments -Wait
        Write-Host "已解除安裝: $($client.DisplayName) $($client.DisplayVersion)"
    }
}

# 函數：安裝新版本
function Install-NewVersion {
    param ($installer)
    $logName = "${env:COMPUTERNAME}_$($installer.VersionInfo.ProductName.Trim())_$($installer.VersionInfo.ProductVersion.Trim()).txt"
    $arguments = "/VERYSILENT /NORESTART /ALLUSERS /FORCECLOSEAPPLICATIONS /LOGCLOSEAPPLICATIONS /LOG=$localTempPath\`"$logName`""
    
    # 複製安裝檔案到暫存目錄
    $tempInstallerPath = "$localTempPath\$($installer.Name)"
    Copy-Item $installer.FullName $tempInstallerPath -Force
    Unblock-File $tempInstallerPath

    # 執行安裝
    Start-Process $tempInstallerPath -ArgumentList $arguments -Wait
    Write-Host "已安裝: $($installer.VersionInfo.ProductName) $($installer.VersionInfo.ProductVersion)"
}

# 函數：檢查並關閉 AutoUpdate.exe
function Stop-AutoUpdateProcess {
    $process = Get-Process | Where-Object { $_.Path -eq $AutoUpdateExePath } | Select-Object -First 1
    if ($process) {
        Write-Host "正在關閉 AutoUpdate.exe 程序..."
        Stop-Process -Id $process.Id -Force
        Start-Sleep -Seconds 2  # 等待進程完全關閉
        Write-Host "AutoUpdate.exe 已關閉"
    } else {
        Write-Host "AutoUpdate.exe 未運行，無需關閉"
    }
}

# 函數：檢查並啟動 AutoUpdate.exe
function Start-AutoUpdateProcess {
    $process = Get-Process | Where-Object { $_.Path -eq $AutoUpdateExePath } | Select-Object -First 1
    if (-not $process) {
        Write-Host "正在啟動 AutoUpdate.exe..."
        Start-Process -FilePath $AutoUpdateExePath
        Write-Host "AutoUpdate.exe 已啟動"
    } else {
        Write-Host "AutoUpdate.exe 已在運行，無需啟動"
    }
}

# 函數：取得並記錄正在執行的程式清單
function Get-RunningProcesses {
    $runningProcessesLog = Join-Path $localTempPath "${env:COMPUTERNAME}_RunningProcesses_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    Write-Host "正在取得正在執行的程式清單..."
    Get-Process | Select-Object Name, Id, Path | Format-Table -AutoSize | Out-File -FilePath $runningProcessesLog -Encoding UTF8
    Write-Host "正在執行的程式清單已保存至: $runningProcessesLog"
    return $runningProcessesLog
}

# 函數：檢查並重啟 Symantec（如果需要）
function Check-AndRestartSymantec {
    param(
        [DateTime]$startTime,
        [string]$symantecPath,
        [string]$password
    )
    $currentTime = Get-Date
    $timeElapsed = $currentTime - $startTime
    if ($timeElapsed.TotalMinutes -ge 5) {
        Write-Host "已經過去 5 分鐘，正在重新啟動 Symantec Endpoint Protection..."
        Start-Process -FilePath $symantecPath -ArgumentList "-p `"$password`" -start" -WindowStyle Hidden
        return $true
    }
    return $false
}

# 新增函數：檢查並強制關閉 unins000 程序
function Stop-Unins000Process {
    $unins000Processes = Get-Process | Where-Object { $_.ProcessName -eq "unins000" }
    if ($unins000Processes) {
        Write-Host "發現 unins000 程序正在運行，正在嘗試關閉..."
        foreach ($process in $unins000Processes) {
            try {
                Stop-Process -Id $process.Id -Force
                Write-Host "已強制關閉 unins000 程序 (PID: $($process.Id))"
            } catch {
                Write-Warning "無法關閉 unins000 程序 (PID: $($process.Id)): $_"
            }
        }
        Start-Sleep -Seconds 2  # 等待進程完全關閉
    } else {
        Write-Host "未發現 unins000 程序在運行"
    }
}

# 主程序
try {
    Write-Host "腳本開始執行時間: $(Get-Date)"
    
    # 取得並記錄正在執行的程式清單
    $runningProcessesLog = Get-RunningProcesses
    
    $SymantecPath = Get-SymantecPath
    $installer = Get-LatestCMEXFontClientInstaller
    if (-not $installer) {
        throw "找不到 CMEX Font Client 安裝檔。"
    }

    Write-Host "找到最新安裝檔: $($installer.FullName)"
    $installedClients = Get-InstalledCMEXFontClients
    $clientCount = @($installedClients).Count
    Write-Host "找到已安裝的用戶端數量: $clientCount"

    if ($clientCount -gt 0) {
        Write-Host "已安裝的用戶端詳細信息："
        foreach ($client in $installedClients) {
            Write-Host "  - 名稱: $($client.DisplayName)"
            Write-Host "    版本: $($client.DisplayVersion)"
            Write-Host "    安裝日期: $($client.InstallDate)"
            Write-Host ""
        }
    } else {
        Write-Host "未找到已安裝的用戶端。"
    }

    $needInstall = $Force_Install -or -not $installedClients -or 
                   ($installedClients | ForEach-Object { [version]$_.DisplayVersion } | 
                    Measure-Object -Maximum).Maximum -lt [version]$installer.VersionInfo.ProductVersion

    if ($needInstall) {
        Write-Host "需要進行安裝或更新。"

        # 無論如何，都先卸載現有版本
        if ($installedClients) {
            Write-Host "正在卸載現有版本..."
            Uninstall-OldVersions $installedClients
        }

        $symantecStopTime = $null
        # 停止 Symantec
        if ($SymantecPath) {
            Write-Host "正在停止 Symantec Endpoint Protection..."
            $password = Get-DecodedSymantecPassword
            Start-Process -FilePath $SymantecPath -ArgumentList "-p `"$password`" -stop" -Wait -WindowStyle Hidden
            $symantecStopTime = Get-Date
        }

        # 檢查並關閉 AutoUpdate.exe
        Stop-AutoUpdateProcess

        # 檢查並強制關閉 unins000 程序
        Stop-Unins000Process

        # 安裝新版本
        Write-Host "正在安裝新版本..."
        $installStartTime = Get-Date
        Install-NewVersion $installer

        # 檢查安裝時間是否超過 5 分鐘，如果是，重啟 Symantec
        $symantecRestarted = $false
        if ($symantecStopTime -and $SymantecPath) {
            $symantecRestarted = Check-AndRestartSymantec -startTime $symantecStopTime -symantecPath $SymantecPath -password $password
        }

        # 檢查並啟動 AutoUpdate.exe
        Start-AutoUpdateProcess

        # 如果 Symantec 還沒有重啟，現在啟動它
        if ($SymantecPath -and -not $symantecRestarted) {
            Write-Host "正在啟動 Symantec Endpoint Protection..."                
            Start-Process -FilePath $SymantecPath -ArgumentList "-p `"$password`" -start" -WindowStyle Hidden
        }

        $password = $null  # 清除密碼變數
    } else {
        Write-Host "不需要更新。目前版本已是最新。"
    }
} catch {
    Write-Error "發生錯誤: $_"
} finally {
    # 確保 Symantec 在腳本結束時是啟動的
    if ($SymantecPath) {
        $symantecProcess = Get-Process | Where-Object { $_.Path -eq $SymantecPath } | Select-Object -First 1
        if (-not $symantecProcess) {
            Write-Host "確保 Symantec Endpoint Protection 在腳本結束時啟動..."
            $password = Get-DecodedSymantecPassword
            Start-Process -FilePath $SymantecPath -ArgumentList "-p `"$password`" -start" -WindowStyle Hidden
            $password = $null
        }
    }

    Write-Host "腳本結束執行時間: $(Get-Date)"
    Stop-Transcript

    # 將日誌從本機複製到網路位置
    $networkLogFolder = Join-Path $Log_Path "CMEX Font Client"
    if (-not (Test-Path $networkLogFolder)) { 
        New-Item -ItemType Directory $networkLogFolder -Force | Out-Null 
    }

    Write-Host "正在複製日誌檔案到網路位置..."
    # 定義更精確的日誌檔案模式
    $logPatterns = @(
        "${env:COMPUTERNAME}_CMEXFontClient_LOG_*.txt",     # 執行日誌
        "${env:COMPUTERNAME}_中推會用戶端更新工具_*.txt",     # 安裝日誌
        "${env:COMPUTERNAME}_*_解除安裝_*.txt",              # 解除安裝日誌
        "${env:COMPUTERNAME}_RunningProcesses_*.txt"        # 正在執行的程式清單
    )

    foreach ($pattern in $logPatterns) {
        $files = Get-ChildItem -Path $localTempPath -Filter $pattern
        foreach ($file in $files) {
            $destinationPath = Join-Path $networkLogFolder $file.Name
            Copy-Item -Path $file.FullName -Destination $destinationPath -Force
            Write-Host "已複製日誌檔案: $($file.Name)"
        }
    }
}