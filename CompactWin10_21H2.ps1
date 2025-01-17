# 設定錯誤處理
$ErrorActionPreference = "Stop"

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
    Write-Host "本地已有評估結果，正在同步到網絡位置..."
    
    # 確保網絡資料夾存在
    if (-not (Test-Path $networkPath)) {
        New-Item -ItemType Directory -Path $networkPath -Force | Out-Null
    }

    # 複製結果文件到網絡位置
    Copy-Item $localResultFile.FullName $networkPath -Force
    Write-Host "結果文件已同步到網絡位置: $networkPath\$($localResultFile.Name)"
    exit
}

# 檢查網絡位置是否已經有結果文件
$networkResultFile = Get-ChildItem -Path $networkPath -Filter "${computerName}_0x*.txt" -ErrorAction SilentlyContinue

if ($networkResultFile) {
    Write-Host "網絡位置已有評估結果，文件: $($networkResultFile.Name)"
    exit
}

if ($osVersion -notlike "*Windows 10 Enterprise LTSC 2021*" -and $osVersion -notlike "*Windows 10 Enterprise*") {
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

        # 運行 setup.exe
        $setupPath = "${driveLetter}:\setup.exe"
        $arguments = "/auto upgrade /quiet /eula accept /DynamicUpdate disable /compat scanonly /copylogs `"$tempFolder`""   
        Unblock-File -Path $setupPath
        Write-Log "開始執行 setup.exe"

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
        Write-Log "Setup.exe Exit Code: $exitCode"

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
            $content = Get-Content $setupactPath -Tail 50 # 讀取最後 50 行
            $resultLines = $content | Select-String "MOUPG  C(SetupManager|SetupHost)::Execute\(\d+\): Result = 0x" | Select-Object -Last 2

            if ($resultLines.Count -eq 2) {
                $setupManagerResult = ($resultLines[0] -split "Result = ")[1].Trim()
                $setupHostResult = ($resultLines[1] -split "Result = ")[1].Trim()

                $exitCodeHex = $setupHostResult  # 使用 SetupHost 的結果作為最終結果
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
        $dismountResult = Dismount-DiskImage -ImagePath $isoPath 
        if ($dismountResult.Attached -eq $false) {
            Write-Log "ISO 檔案已成功卸載"
        } else {
            Write-Log "卸載 ISO 檔案時發生錯誤"
        }

        # 使用 WMI 退出光碟機
        Write-Log "正在退出光碟機 ${driveLetter}:..."
        try {
            $volume = Get-WmiObject Win32_Volume | Where-Object { $_.DriveLetter -eq "${driveLetter}:" }
            if ($volume) {
                $result = $volume.Eject()
                if ($result.ReturnValue -eq 0) {
                    Write-Log "光碟機 ${driveLetter}: 已成功退出"
                } else {
                    Write-Log "退出光碟機 ${driveLetter}: 時發生錯誤。錯誤代碼: $($result.ReturnValue)"
                }
            } else {
                Write-Log "找不到光碟機 ${driveLetter}:"
            }
        } catch {
            Write-Log "退出光碟機時發生異常: $_"
        }

        # 檢查光碟機是否仍然存在
        if (Test-Path "${driveLetter}:") {
            Write-Log "警告: 光碟機 ${driveLetter}: 仍然可見。可能需要手動退出。"
        } else {
            Write-Log "光碟機 ${driveLetter}: 不再可見。"
        }
    }
} else {
    Write-Log "當前系統已經是 Windows 10 Enterprise LTSC 2021 或更新版本，不需要評估。"
}