#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Windows 10 Enterprise LTSC 2021 升級評估和安裝腳本。

.DESCRIPTION
    此腳本執行 Windows 10 Enterprise LTSC 2021 的相容性評估和升級安裝。
    它包括相容性掃描、系統升級、日誌記錄和結果同步等功能。

.PARAMETER ConfigPath
    指定設定檔的路徑。預設為腳本所在目錄下的 config.json。

.EXAMPLE
    .\Windows10-2021Upgrade.ps1
    使用預設設定檔運行腳本。

.EXAMPLE
    .\Windows10-2021Upgrade.ps1 -ConfigPath "C:\CustomConfig.json"
    使用指定的設定檔運行腳本。
#>

param (
    [string]$ConfigPath = "$PSScriptRoot\config.json"
)

# 定義預設設定
$defaultConfig = @{
    TempFolder = "${env:systemdrive}\temp\${env:computername}"
    NetworkPath = "\\172.29.205.114\Public\sources\audit\Win10_x64_2021_LTSC(21H2)"
    ISOPath = "\\172.29.205.114\loginscript\Update\Windows10\Win10_x64_企業版_2021_LTSC(21H2).ISO"
    UpgradeLogFolder = "${env:systemdrive}\temp\${env:computername}\AfterCompat"
    AllowedComputers = @("")
}

# 定義退出代碼及其意義
$script:exitCodes = @{
    "0xC1900210" = "未發現任何相容性問題"
    "0xC1900208" = "發現可操作的相容性問題，如應用程式相容性問題"
    "0xC1900204" = "所選的遷移選項不可用"
    "0xC1900200" = "機器不符合 Windows 10 或以上版本的要求"
    "0xC190020E" = "機器沒有足夠的可用空間進行安裝"
}

# 匯入設定
function Import-Config {
    param ([string]$Path)
    try {
        if (Test-Path $Path) {
            $config = Get-Content $Path -Raw | ConvertFrom-Json
            Write-Host "成功讀取設定檔: $Path"
            return $config
        } else {
            throw "設定檔不存在: $Path"
        }
    } catch {
        Write-Warning "無法讀取設定檔: $Path"
        Write-Warning "錯誤: $_"
        Write-Host "使用內建預設設定。請檢查設定檔是否存在且格式正確。"
        return $defaultConfig
    }
}

# 初始化設定
$script:Config = Import-Config -Path $ConfigPath

# 確保所有必要的設定項目都存在
foreach ($key in $defaultConfig.Keys) {
    if (-not $script:Config.$key) {
        Write-Warning "設定中缺少 $key，使用預設值。"
        $script:Config | Add-Member -NotePropertyName $key -NotePropertyValue $defaultConfig.$key
    }
}
# 初始化日誌
function Initialize-Logging {
    param (
        [string]$LogPath,
        [switch]$Append
    )
    try {
        New-Item -ItemType Directory -Force -Path (Split-Path $LogPath) | Out-Null
        if ($Append) {
            Start-Transcript -Path $LogPath -Append
        } else {
            Start-Transcript -Path $LogPath
        }
        Write-Host "日誌記錄開始於 $(Get-Date)"
    } catch {
        Write-Warning "無法初始化日誌記錄: $_"
    }
}

# 寫入日誌
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage

    $logFilePath = "$($script:Config.TempFolder)\script_log.txt"
    $logFileDir = Split-Path -Parent $logFilePath

    # 確保日誌文件的目錄存在
    if (-not (Test-Path $logFileDir)) {
        try {
            New-Item -ItemType Directory -Force -Path $logFileDir | Out-Null
            Write-Host "Created log directory: $logFileDir"
        }
        catch {
            Write-Host "Error creating log directory: $_"
            return
        }
    }

    # 使用 Add-Content 追加寫入日誌文件
    try {
        Add-Content -Path $logFilePath -Value $logMessage
    }
    catch {
        Write-Host "Error writing to log file: $_"
    }
}

# 獲取 Windows LTSC 版本
function Get-WindowsLTSCVersion {
    $osInfo = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $buildNumber = $osInfo.CurrentBuildNumber
    $ubr = $osInfo.UBR
    $displayVersion = $osInfo.DisplayVersion

    $versionInfo = switch ($buildNumber) {
        "10240" { @{Version = "LTSB 2015"; ReleaseId = "1507"} }
        "14393" { @{Version = "LTSB 2016"; ReleaseId = "1607"} }
        "17763" { @{Version = "LTSC 2019"; ReleaseId = "1809"} }
        "19044" { @{Version = "LTSC 2021"; ReleaseId = "21H2"} }
        default { @{Version = "Unknown"; ReleaseId = "Unknown"} }
    }

    return @{
        BuildNumber = $buildNumber
        UBR = $ubr
        DisplayVersion = $displayVersion
        LTSCVersion = $versionInfo.Version
        ReleaseId = $versionInfo.ReleaseId
    }
}

# 掛載 ISO 檔案
function Mount-ISOFile {
    param ([string]$ISOPath)
    Write-Log "正在掛載 ISO 檔案: $ISOPath"
    try {
        $mountResult = Mount-DiskImage -ImagePath $ISOPath -PassThru
        $driveLetter = ($mountResult | Get-Volume).DriveLetter
        Write-Log "ISO 檔案已掛載到磁區 ${driveLetter}:"
        return $driveLetter
    } catch {
        Write-Log "掛載 ISO 檔案時發生錯誤: $_" -Level ERROR
        throw
    }
}

# 卸載 ISO 檔案
function Unmount-ISOFile {
    param ([string]$ISOPath)
    Write-Log "正在卸載 ISO 檔案..."
    try {
        Dismount-DiskImage -ImagePath $ISOPath
        Write-Log "ISO 檔案已成功卸載"
    } catch {
        Write-Log "卸載 ISO 檔案時發生錯誤: $_" -Level ERROR
    }
}
# 分析 setupact.log 檔案
function Analyze-SetupActLog {
    param (
        [string]$LogPath
    )
    Write-Log "開始分析 setupact.log 檔案"
    if (Test-Path $LogPath) {
        $content = Get-Content $LogPath -Tail 50
        $resultLines = $content | Select-String "MOUPG  C(SetupManager|SetupHost)::Execute\(\d+\): Result = 0x" | Select-Object -Last 2

        if ($resultLines.Count -eq 2) {
            $setupManagerResult = ($resultLines[0] -split "Result = ")[1].Trim()
            $setupHostResult = ($resultLines[1] -split "Result = ")[1].Trim()

            $exitCodeHex = $setupHostResult
            $meaning = $script:exitCodes[$exitCodeHex]
            
            Write-Log "SetupManager Result: $setupManagerResult"
            Write-Log "SetupHost Result: $setupHostResult"
            Write-Log "退出代碼: $exitCodeHex"
            Write-Log "意義: $meaning"

            return @{
                SetupManagerResult = $setupManagerResult
                SetupHostResult = $setupHostResult
                ExitCode = $exitCodeHex
                Meaning = $meaning
            }
        } else {
            Write-Log "警告: 在日誌檔案中找不到預期的結果行" -Level WARNING
            return $null
        }
    } else {
        Write-Log "警告: 找不到 setupact.log 檔案" -Level WARNING
        return $null
    }
}

# 執行相容性掃描
function Invoke-CompatibilityScan {
    param (
        [string]$SetupPath,
        [string]$LogFolder
    )
    # 檢查是否已經執行過評估
    $existingResultFile = Get-ChildItem -Path $LogFolder -Filter "${env:computername}_0x*.txt" -ErrorAction SilentlyContinue
    if ($existingResultFile) {
        Write-Log "檢測到現有的評估結果文件: $($existingResultFile.Name)，跳過評估程序"
        return $null
    }

    Write-Log "開始執行相容性掃描"
    $arguments = "/auto upgrade /quiet /eula accept /DynamicUpdate disable /compat scanonly /copylogs `"$LogFolder`""
    try {
        $startTime = Get-Date
        $process = Start-Process -FilePath $SetupPath -ArgumentList $arguments -Wait -PassThru -NoNewWindow
        $endTime = Get-Date
        $duration = $endTime - $startTime
        
        Write-Log "相容性掃描完成。Exit Code: $($process.ExitCode)"
        Write-Log "相容性掃描開始時間: $startTime"
        Write-Log "相容性掃描結束時間: $endTime"
        Write-Log "相容性掃描總耗時: $($duration.ToString())"
        
        # 分析 setupact.log
        $setupActPath = Join-Path $LogFolder "Panther\setupact.log"
        $analysisResult = Analyze-SetupActLog -LogPath $setupActPath
        
        if ($analysisResult) {
            # 建立結果檔案
            $resultFileName = "${env:computername}_$($analysisResult.ExitCode)_$($analysisResult.Meaning).txt"
            $localResultPath = Join-Path $LogFolder $resultFileName
            $networkResultPath = Join-Path $script:Config.NetworkPath $resultFileName
            
            # 使用陣列和 Join 方法建立結果內容
            $resultContent = @(
                "SetupManager Result: $($analysisResult.SetupManagerResult)",
                "SetupHost Result: $($analysisResult.SetupHostResult)",
                "退出代碼: $($analysisResult.ExitCode)",
                "意義: $($analysisResult.Meaning)",
                "掃描開始時間: $startTime",
                "掃描結束時間: $endTime",
                "掃描總耗時: $($duration.ToString())"
            ) -join "`n"

            # 保存到本地
            $resultContent | Out-File $localResultPath -Encoding UTF8
            Write-Log "建立本地結果檔案: $localResultPath"

            # 保存到網絡位置
            try {
                $resultContent | Out-File $networkResultPath -Encoding UTF8
                Write-Log "建立網絡結果檔案: $networkResultPath"
            } catch {
                Write-Log "無法保存結果檔案到網絡位置: $networkResultPath. 錯誤: $_" -Level WARNING
            }
        }

        return $process.ExitCode
    } catch {
        Write-Log "執行相容性掃描時發生錯誤: $_" -Level ERROR
        throw
    }
}
# 執行系統升級
function Invoke-SystemUpgrade {
    param (
        [string]$SetupPath,
        [string]$LogFolder
    )
    Write-Log "開始執行系統升級"
    $arguments = "/auto upgrade /quiet /eula accept /DynamicUpdate disable /copylogs `"$LogFolder`" /PostOOBE `"$($script:Config.TempFolder)\setupcomplete.cmd`" /noreboot"
    try {
        $startTime = Get-Date
        $process = Start-Process -FilePath $SetupPath -ArgumentList $arguments -Wait -PassThru -NoNewWindow
        $endTime = Get-Date
        $duration = $endTime - $startTime
        Write-Log "升級完成。Exit Code: $($process.ExitCode)"
        Write-Log "升級開始時間: $startTime"
        Write-Log "升級結束時間: $endTime"
        Write-Log "升級總耗時: $($duration.ToString())"
        return $process.ExitCode
    } catch {
        Write-Log "執行系統升級時發生錯誤: $_" -Level ERROR
        throw
    }
}

# 創建 setupcomplete.cmd
function Create-SetupCompleteCmd {
    $setupCompletePath = Join-Path $script:Config.TempFolder "setupcomplete.cmd"
    $setupCompleteContent = @'
@echo off
chcp 65001
setlocal enabledelayedexpansion

:: 設定日誌文件路徑
set "LOGFILE=%SystemDrive%\Windows\Setup\Scripts\setupcomplete.log"

:: 創建日誌文件
echo Setup Complete Script started at %date% %time% > %LOGFILE%

:: 設定源和目標路徑
set "SOURCE=%SystemDrive%\Windows.old\Windows\Setup\Scripts"
set "DESTINATION=%SystemDrive%\Windows\Setup\Scripts"

:: 檢查源路徑是否存在
if not exist "%SOURCE%" (
    echo 源資料夾不存在: %SOURCE% >> %LOGFILE%
    goto :END
)

:: 確保目標路徑存在
if not exist "%DESTINATION%" mkdir "%DESTINATION%"

:: 使用 robocopy 複製文件
robocopy "%SOURCE%" "%DESTINATION%" /E /XO /XX /NP /NDL /NJH /NJS /NC /NS /LOG+:%LOGFILE%

:: 檢查 robocopy 的退出代碼
if %ERRORLEVEL% GEQ 8 (
    echo Robocopy 遇到錯誤。退出代碼: %ERRORLEVEL% >> %LOGFILE%
) else if %ERRORLEVEL% GTR 1 (
    echo Robocopy 完成，但有些文件或目錄被跳過。退出代碼: %ERRORLEVEL% >> %LOGFILE%
) else (
    echo Robocopy 成功完成。退出代碼: %ERRORLEVEL% >> %LOGFILE%
)

:END
echo Setup Complete Script 結束於 %date% %time% >> %LOGFILE%
exit /b 0
'@

    $setupCompleteContent | Out-File -FilePath $setupCompletePath -Encoding UTF8
    Write-Log "已創建 setupcomplete.cmd: $setupCompletePath"
}

# 同步日誌到網路位置
function Sync-LogsToNetwork {
    param (
        [string]$SourceFolder,
        [string]$DestinationFolder
    )
    Write-Log "開始同步日誌到網路位置"
    try {
        if (-not (Test-Path $DestinationFolder)) {
            New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
        }
        
        $robocopyArgs = @(
            $SourceFolder,
            $DestinationFolder,
            "/E",        # 複製子目錄，包括空的子目錄
            "/Z",        # 可恢復模式
            "/R:3",      # 重試 3 次
            "/W:5",      # 重試間隔 5 秒
            "/MT:8",     # 8 個執行緒
            "/NFL",      # 不記錄檔案名
            "/NDL"       # 不記錄目錄名
        )

        $robocopyResult = Start-Process -FilePath "robocopy" -ArgumentList $robocopyArgs -NoNewWindow -Wait -PassThru

        if ($robocopyResult.ExitCode -lt 8) {
            Write-Log "日誌同步完成"
            return $true
        } else {
            Write-Log "日誌同步過程中出現警告或錯誤，Robocopy 退出代碼: $($robocopyResult.ExitCode)" -Level WARNING
            return $false
        }
    } catch {
        Write-Log "同步日誌到網路位置時發生錯誤: $_" -Level ERROR
        return $false
    }
}
# 主要執行邏輯
function Main {
    try {
        $windowsInfo = Get-WindowsLTSCVersion
        $computerName = $env:COMPUTERNAME

        # 使用 append 模式初始化日誌
        Initialize-Logging -LogPath "$($script:Config.TempFolder)\upgrade_script.log" -Append

        Write-Log "開始執行 Windows 10 Enterprise LTSC 2021 升級腳本"
        Write-Log "當前系統版本: Windows 10 Enterprise $($windowsInfo.LTSCVersion)"
        Write-Log "構建號: $($windowsInfo.BuildNumber).$($windowsInfo.UBR)"
        Write-Log "版本號: $($windowsInfo.ReleaseId)"
        Write-Log "電腦名稱: $computerName"

        # 定義需要升級的版本
        $upgradeVersions = @("LTSB 2015", "LTSB 2016", "LTSC 2019")

        if ($windowsInfo.LTSCVersion -eq "LTSC 2021") {
            Write-Log "當前系統已經是 Windows 10 Enterprise LTSC 2021，無需升級"
        }
        elseif ($windowsInfo.LTSCVersion -in $upgradeVersions) {
            Write-Log "檢測到可升級的 Windows 10 Enterprise LTSC/LTSB 版本"
            
            $localResultFile = Get-ChildItem -Path $script:Config.TempFolder -Filter "${computerName}_0x*.txt" -ErrorAction SilentlyContinue
            $networkResultFile = Get-ChildItem -Path $script:Config.NetworkPath -Filter "${computerName}_0x*.txt" -ErrorAction SilentlyContinue

            if (-not $localResultFile -and -not $networkResultFile) {
                # 執行評估程序
                $driveLetter = Mount-ISOFile -ISOPath $script:Config.ISOPath
                $setupPath = "${driveLetter}:\setup.exe"
                $exitCode = Invoke-CompatibilityScan -SetupPath $setupPath -LogFolder $script:Config.TempFolder
                if ($exitCode -ne $null) {
                    Write-Log "評估程序結束，Exit Code: $exitCode"
                }
            }
            elseif (($localResultFile -or $networkResultFile) -and (($localResultFile -and $localResultFile.Name -like "*0xC1900210*") -or ($networkResultFile -and $networkResultFile.Name -like "*0xC1900210*"))) {
                # 檢查是否允許所有電腦或特定電腦
                    if ($script:Config.AllowedComputers.Count -eq 0 -or 
                        $script:Config.AllowedComputers[0] -eq "" -or 
                        $computerName -in $script:Config.AllowedComputers) {
                    
                    # 創建 setupcomplete.cmd
                    Create-SetupCompleteCmd
        
                    # 執行升級程序
                    Write-Log "電腦被允許升級，開始執行升級程序"
                    $driveLetter = Mount-ISOFile -ISOPath $script:Config.ISOPath
                    $setupPath = "${driveLetter}:\setup.exe"
                    $exitCode = Invoke-SystemUpgrade -SetupPath $setupPath -LogFolder $script:Config.UpgradeLogFolder
                    Write-Log "升級程序結束，Exit Code: $exitCode"
                    }
                    else {
                        Write-Log "電腦不在允許升級列表中，不執行升級程序" -Level WARNING
                    }
            }
            else {
                Write-Log "評估結果不符合升級條件，不執行升級程序"
            }

            # 同步日誌到網路位置
            $syncSuccess = Sync-LogsToNetwork -SourceFolder $script:Config.TempFolder -DestinationFolder "$($script:Config.NetworkPath)\$computerName"
            if ($syncSuccess) {
                Write-Log "日誌同步成功"
            }
            else {
                Write-Log "日誌同步失敗" -Level WARNING
            }

            # 如果有升級日誌，也同步它們
            if (Test-Path $script:Config.UpgradeLogFolder) {
                Sync-LogsToNetwork -SourceFolder $script:Config.UpgradeLogFolder -DestinationFolder "$($script:Config.NetworkPath)\$computerName\Upgrade"
            }
        }
        else {
            Write-Log "當前系統不是目標升級版本的 Windows 10 Enterprise LTSC/LTSB，不執行評估或升級"
        }
    }
    catch {
        Write-Log "腳本執行過程中發生錯誤: $_" -Level ERROR
    }
    finally {
        if ($driveLetter) {
            Unmount-ISOFile -ISOPath $script:Config.ISOPath
        }
        # 最後再同步一次，確保捕獲所有日誌
        Sync-LogsToNetwork -SourceFolder $script:Config.TempFolder -DestinationFolder "$($script:Config.NetworkPath)\$computerName"
        Write-Log "腳本執行完成"
        Stop-Transcript
    }
}

# 執行主函數
Main
