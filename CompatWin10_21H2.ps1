# 設定錯誤處理
$ErrorActionPreference = "Stop"

# 設定允許執行升級的電腦列表
$allowedComputers = @("TND-STOF-113")

# 設定日誌函數
function Write-Log {
    param([string]$Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Write-Host $logMessage
    Add-Content -Path "$tempFolder\script_log.txt" -Value $logMessage
}

# 定義退出代碼及其意義
$exitCodes = @{
    "0xC1900210" = "未發現任何相容性問題"
    "0xC1900208" = "發現可操作的相容性問題，如應用程序相容性問題"
    "0xC1900204" = "所選的遷移選項不可用"
    "0xC1900200" = "機器不符合 Windows 10 或以上版本的要求"
    "0xC190020E" = "機器沒有足夠的可用空間進行安裝"
}

# 獲取電腦名稱
$computerName = $env:COMPUTERNAME

# 檢查 Windows 版本
$osVersion = (Get-WmiObject Win32_OperatingSystem).Caption

# 設置路徑
$tempFolder = "${env:systemdrive}\temp\$computerName"
$networkPath = "\\172.29.205.114\Public\sources\audit\Win10_x64_2021_LTSC(21H2)\$computerName"

# 檢查本地臨時資料夾是否已經有結果文件
$localResultFile = Get-ChildItem -Path $tempFolder -Filter "${computerName}_0x*.txt" -ErrorAction SilentlyContinue

if ($localResultFile) {
    Write-Log "本地已有評估結果，正在同步到網絡位置..."
    
    # 確保網絡資料夾存在
    if (-not (Test-Path $networkPath)) {
        New-Item -ItemType Directory -Path $networkPath -Force | Out-Null
    }

    # 複製結果文件到網絡位置
    Copy-Item $localResultFile.FullName $networkPath -Force
    Write-Log "結果文件已同步到網絡位置: $networkPath\$($localResultFile.Name)"
    
    # 設置 $networkResultFile 變量，以便後續使用
    $networkResultFile = Get-Item "$networkPath\$($localResultFile.Name)"
} else {
    # 如果本地沒有結果文件，檢查網絡位置
    $networkResultFile = Get-ChildItem -Path $networkPath -Filter "${computerName}_0x*.txt" -ErrorAction SilentlyContinue
}

# 檢查是否有評估結果文件（本地同步的或網絡上的）
if ($networkResultFile) {
    Write-Log "發現評估結果文件: $($networkResultFile.Name)"
    
    # 檢查當前系統版本
    if ($osVersion -like "*Windows 10 Enterprise LTSC 2021*") {
        Write-Log "當前系統已經是 Windows 10 Enterprise LTSC 2021。"
        
        # 檢查 Windows.old 文件夾和 Scripts 目錄
        if (Test-Path "${env:systemdrive}\Windows.old") {
            Write-Log "檢測到 Windows.old 文件夾。"
            if (-not (Test-Path "${env:windir}\setup\Scripts")) {
                Write-Log "${env:windir}\setup\Scripts 不存在，嘗試從 Windows.old 複製。"
                if (Test-Path "${env:systemdrive}\Windows.old\Windows\setup\Scripts") {
                    Copy-Item -Path "${env:systemdrive}\Windows.old\Windows\setup\Scripts" -Destination "${env:windir}\setup" -Recurse -Force
                    Write-Log "已從 Windows.old 複製 Scripts 文件夾。"
                } else {
                    Write-Log "Windows.old 中也沒有找到 Scripts 文件夾。"
                }
            } else {
                Write-Log "${env:windir}\setup\Scripts 已存在。"
            }
        } else {
            Write-Log "沒有檢測到 Windows.old 文件夾。"
        }
    }
    # 檢查是否為 0xC1900210（未發現任何相容性問題）
    elseif ($networkResultFile.Name -like "*0xC1900210*") {
        Write-Log "檢測到相容性評估結果為 0xC1900210，準備執行升級..."
        
        # 檢查當前電腦是否在允許列表中
        if ($computerName -in $allowedComputers) {
            try {
                # 設置新的日誌文件夾
                $upgradeLogFolder = "${env:systemdrive}\temp\${computerName}_AfterCompat"
                
                # 創建新的日誌文件夾（如果不存在）
                if (-not (Test-Path $upgradeLogFolder)) {
                    New-Item -ItemType Directory -Force -Path $upgradeLogFolder | Out-Null
                    Write-Log "創建升級日誌文件夾: $upgradeLogFolder"
                }

                # 設置 ISO 檔案路徑
                $isoPath = "\\172.29.205.114\loginscript\Update\Windows10\Win10_x64_企業版_2021_LTSC(21H2).ISO"

                # 掛載 ISO 檔案
                Write-Log "正在掛載 ISO 檔案: $isoPath"
                $mountResult = Mount-DiskImage -ImagePath $isoPath -PassThru
                $driveLetter = ($mountResult | Get-Volume).DriveLetter
                Write-Log "ISO 檔案已掛載到光碟機 ${driveLetter}:"

                # 執行 setup.exe
                $setupPath = "${driveLetter}:\setup.exe"
                $arguments = "/auto upgrade /quiet /eula accept /DynamicUpdate disable /copylogs `"$upgradeLogFolder`" /priority low /skipfinalize"
                
                Write-Log "開始執行升級: $setupPath $arguments"
                
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = $setupPath
                $processInfo.Arguments = $arguments
                $processInfo.UseShellExecute = $false
                $processInfo.CreateNoWindow = $true
                $processInfo.RedirectStandardOutput = $true
                $processInfo.RedirectStandardError = $true

                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $processInfo

                # 記錄開始時間
                $startTime = Get-Date
                Write-Log "升級開始時間: $startTime"

                $process.Start() | Out-Null
                $process.WaitForExit()

                # 記錄結束時間並計算執行時間
                $endTime = Get-Date
                $executionTime = $endTime - $startTime
                Write-Log "升級結束時間: $endTime"
                Write-Log "總執行時間: $($executionTime.ToString())"

                $exitCode = $process.ExitCode
                Write-Log "升級程序完成。Exit Code: $exitCode"

                # 獲取輸出和錯誤信息
                $output = $process.StandardOutput.ReadToEnd()
                $errorOutput = $process.StandardError.ReadToEnd()

                if ($output) {
                    Write-Log "Setup.exe Output: $output"
                }
                if ($errorOutput) {
                    Write-Log "Setup.exe Error: $errorOutput"
                }

                # 將執行時間寫入結果文件
                Add-Content -Path $networkResultFile.FullName -Value "`n升級開始時間: $startTime"
                Add-Content -Path $networkResultFile.FullName -Value "升級結束時間: $endTime"
                Add-Content -Path $networkResultFile.FullName -Value "總執行時間: $($executionTime.ToString())"

                # 複製最新的 script_log.txt 到升級日誌文件夾
                $latestLogFile = Get-ChildItem -Path $tempFolder -Filter "script_log.txt" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($latestLogFile) {
                    Copy-Item -Path $latestLogFile.FullName -Destination $upgradeLogFolder -Force
                    Write-Log "已將最新的 script_log.txt 複製到升級日誌文件夾: $upgradeLogFolder"
                } else {
                    Write-Log "警告：未找到 script_log.txt 文件"
                }

                # 將升級日誌文件夾同步到網絡位置
                $networkUpgradeLogPath = Join-Path $networkPath "UpgradeLog"
                if (-not (Test-Path $networkUpgradeLogPath)) {
                    New-Item -ItemType Directory -Path $networkUpgradeLogPath -Force | Out-Null
                }
                Copy-Item -Path "$upgradeLogFolder\*" -Destination $networkUpgradeLogPath -Recurse -Force
                Write-Log "已將升級日誌文件夾同步到網絡位置: $networkUpgradeLogPath"

            } catch {
                Write-Log "執行升級時發生錯誤: $_"
            } finally {
                # 卸載 ISO 檔案
                Write-Log "正在卸載 ISO 檔案..."
                Dismount-DiskImage -ImagePath $isoPath 
                Write-Log "ISO 檔案已卸載"
            }
        } else {
            Write-Log "當前電腦 $computerName 不在允許執行升級的列表中，跳過升級過程。"
        }
    } else {
        Write-Log "相容性評估結果不是 0xC1900210，不執行升級。"
    }
} else {
    Write-Log "未找到評估結果文件，需要執行相容性掃描。"

    if ($osVersion -notlike "*Windows 10 Enterprise LTSC 2021*") {
        try {
            # 創建臨時文件夾
            New-Item -ItemType Directory -Force -Path $tempFolder | Out-Null
            Write-Log "創建臨時文件夾: $tempFolder"

            # 設置 ISO 檔案路徑
            $isoPath = "\\172.29.205.114\loginscript\Update\Windows10\Win10_x64_企業版_2021_LTSC(21H2).ISO"

            # 掛載 ISO 檔案
            Write-Log "正在掛載 ISO 檔案..."
            $mountResult = Mount-DiskImage -ImagePath $isoPath -PassThru
            $driveLetter = ($mountResult | Get-Volume).DriveLetter
            Write-Log "ISO 檔案已掛載到光碟機 ${driveLetter}:"

            # 運行 setup.exe 進行相容性掃描
            $setupPath = "${driveLetter}:\setup.exe"
            $arguments = "/auto upgrade /quiet /eula accept /DynamicUpdate disable /compat scanonly /copylogs `"$tempFolder`""   
            Write-Log "開始執行相容性掃描: $setupPath $arguments"

            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $setupPath
            $processInfo.Arguments = $arguments
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $true
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $process.Start() | Out-Null
            $process.WaitForExit()

            $exitCode = $process.ExitCode
            Write-Log "相容性掃描完成。Exit Code: $exitCode"

            # 獲取輸出和錯誤信息
            $output = $process.StandardOutput.ReadToEnd()
            $errorOutput = $process.StandardError.ReadToEnd()

            if ($output) {
                Write-Log "Setup.exe Output: $output"
            }
            if ($errorOutput) {
                Write-Log "Setup.exe Error: $errorOutput"
            }

            # 分析日誌文件
            $setupactPath = "$tempFolder\Panther\setupact.log"
            if (Test-Path $setupactPath) {
                $content = Get-Content $setupactPath -Tail 50
                $resultLines = $content | Select-String "MOUPG  C(SetupManager|SetupHost)::Execute\(\d+\): Result = 0x" | Select-Object -Last 2

                if ($resultLines.Count -eq 2) {
                    $setupManagerResult = ($resultLines[0] -split "Result = ")[1].Trim()
                    $setupHostResult = ($resultLines[1] -split "Result = ")[1].Trim()

                    $exitCodeHex = $setupHostResult
                    $meaning = $exitCodes[$exitCodeHex]
                    
                    # 創建結果文件
                    $resultFileName = "${computerName}_${exitCodeHex}_$meaning.txt"
                    $resultFilePath = Join-Path $tempFolder $resultFileName

                    $resultContent = @(
                        "SetupManager Result: $setupManagerResult",
                        "SetupHost Result: $setupHostResult",
                        "退出代碼: $exitCodeHex",
                        "意義: $meaning"
                    ) -join [Environment]::NewLine

                    $resultContent | Out-File $resultFilePath -Encoding UTF8

                    Write-Log "創建結果文件: $resultFilePath"
                    Write-Log "SetupManager Result: $setupManagerResult"
                    Write-Log "SetupHost Result: $setupHostResult"
                } else {
                    Write-Log "警告: 在日誌文件中找不到預期的結果行"
                }
            } else {
                Write-Log "警告: 找不到 setupact.log 文件"
            }

            # 複製文件到網絡位置
            if (Test-Path $networkPath) {
                Remove-Item $networkPath -Recurse -Force
            }
            Copy-Item $tempFolder $networkPath -Recurse
            Write-Log "複製文件到網絡位置: $networkPath"

        } catch {
            Write-Log "錯誤: $_"
        } finally {
            # 卸載 ISO 檔案
            Write-Log "正在卸載 ISO 檔案..."
            Dismount-DiskImage -ImagePath $isoPath 
            Write-Log "ISO 檔案已卸載"
        }
    } else {
        Write-Log "當前系統已經是 Windows 10 Enterprise LTSC 2021，不需要評估或升級。"
        
        # 檢查 Windows.old 文件夾和 Scripts 目錄
        if (Test-Path "${env:systemdrive}\Windows.old") {
            Write-Log "檢測到 Windows.old 文件夾。"
            if (-not (Test-Path "${env:windir}\setup\Scripts")) {
                Write-Log "${env:windir}\setup\Scripts 不存在，嘗試從 Windows.old 複製。"
                if (Test-Path "${env:systemdrive}\Windows.old\Windows\setup\Scripts") {
                    Copy-Item -Path "${env:systemdrive}\Windows.old\Windows\setup\Scripts" -Destination "${env:windir}\setup" -Recurse -Force
                    Write-Log "已從 Windows.old 複製 Scripts 文件夾。"
                } else {
                    Write-Log "Windows.old 中也沒有找到 Scripts 文件夾。"
                }
            } else {
                Write-Log "${env:windir}\setup\Scripts 已存在。"
            }
        } else {
            Write-Log "沒有檢測到 Windows.old 文件夾。"
        }
    }
}

Write-Log "腳本執行完成。"