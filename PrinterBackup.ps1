# 啟用嚴格模式，以幫助擷取錯誤
Set-StrictMode -Version Latest

# 定義路徑變數
$systemDrive = $env:SystemDrive
$computerName = $env:ComputerName
$paths = @{
    Local = "$systemDrive\temp"
    Nas = "\\172.29.205.114\mig\Printer"
    NasBak = "\\172.29.205.114\mig\Printer_BACKUP"
}
$fileName = "${computerName}x64.printerExport"
$files = @{
    Local = Join-Path $paths.Local $fileName
    Nas = Join-Path $paths.Nas $fileName
    NasBak = Join-Path $paths.NasBak $fileName
}

# 定義日誌檔案名稱和路徑
$logFileName = "${computerName}_PrinterExport.log"
$localLogPath = Join-Path $paths.Local $logFileName
$nasLogPath = Join-Path $paths.Nas $logFileName

# 函數：寫入日誌
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Write-Host $logMessage
    Add-Content -Path $localLogPath -Value $logMessage
}

# 函數：清理舊的日誌檔案
function Clean-OldLogs {
    param([string]$Path)
    if (Test-Path $Path) {
        Remove-Item -Path $Path -Force
        Write-Log "已刪除舊的日誌檔案: $Path"
    }
}

# 清理舊的日誌檔案
Clean-OldLogs -Path $localLogPath
Clean-OldLogs -Path $nasLogPath

Write-Log "開始執行印表機匯出腳本"

# 檢查作業系統版本
if ([System.Environment]::OSVersion.Version.Major -lt 10) {
    Write-Log "哎呀！這個腳本需要 Windows 10 以上的版本才能執行喔。"
    exit
}

# 設置保留天數
$retentionDays = 30
$cutoffDate = (Get-Date).AddDays(-$retentionDays)

# 檢查並處理現有的 NAS 檔案
if (Test-Path $files.Nas) {
    $fileInfo = Get-Item $files.Nas
    if ($fileInfo.CreationTime -gt $cutoffDate) {
        Write-Log "哇！發現最近的匯出檔案耶。不用擔心，我們就到這邊囉～"
        exit
    }
    
    try {
        Move-Item -Path $files.Nas -Destination $files.NasBak -Force -ErrorAction Stop
        Remove-Item -Path $files.Nas -Force -ErrorAction Stop
        Write-Log "成功移動並刪除舊的 NAS 檔案。"
    }
    catch {
        Write-Log "處理舊的 NAS 檔案時發生錯誤：$($_.Exception.Message)"
    }
}

# 清理未使用的印表機連接埠
Write-Log "開始清理未使用的印表機連接埠..."
try {
    $allPorts = Get-PrinterPort -ErrorAction Stop
    $usedPorts = (Get-Printer -ErrorAction Stop).PortName
    $portsToDelete = $allPorts | Where-Object { 
        $_.Name -notin $usedPorts -and 
        $_.CimClass -like "ROOT/StandardCimv2:MSFT_TcpIpPrinterPort"
    }

    foreach ($port in $portsToDelete) {
        if ($null -ne $port.Name) {
            try {
                Remove-PrinterPort -Name $port.Name -ErrorAction Stop
                Write-Log "成功刪除未使用的連接埠：$($port.Name)"
            }
            catch {
                Write-Log "刪除連接埠 $($port.Name) 時發生錯誤：$($_.Exception.Message)"
            }
        }
        else {
            Write-Log "警告：發現名稱為 Null 的連接埠，已跳過。"
        }
    }
}
catch {
    Write-Log "獲取印表機或連接埠資訊時發生錯誤：$($_.Exception.Message)"
}

# 執行印表機匯出
if (Test-Path $files.Local) {
    Remove-Item -Path $files.Local -Force -ErrorAction SilentlyContinue
}

$printbrmPath = "${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe"
Write-Log "開始匯出印表機設定囉，請稍等一下下～"
try {
    $printbrmOutput = & $printbrmPath -B -F $files.Local 2>&1
    $printbrmExitCode = $LASTEXITCODE

    switch ($printbrmExitCode) {
        0 {
            Write-Log "印表機設定匯出成功完成。"
        }
        { $_ -gt 0 } {
            Write-Log "印表機設定匯出完成，但有警告。退出碼：$printbrmExitCode"
            Write-Log "Printbrm 輸出：$printbrmOutput"
        }
        default {
            throw "Printbrm 執行失敗。退出碼：$printbrmExitCode"
        }
    }
}
catch {
    Write-Log "匯出印表機設定時發生錯誤：$($_.Exception.Message)"
    if ($printbrmOutput) {
        Write-Log "Printbrm 輸出：$printbrmOutput"
    }
    exit
}

# 複製檔案到 NAS
Write-Log "正在將檔案複製到 NAS，馬上就好！"
$robocopyArgs = @(
    $paths.Local,
    $paths.Nas,
    $fileName,
    "/PURGE",
    "/XO",
    "/NJH",
    "/NJS",
    "/NDL",
    "/NC",
    "/NS"
)
try {
    $robocopyOutput = & robocopy $robocopyArgs 2>&1
    $robocopyExitCode = $LASTEXITCODE
    if ($robocopyExitCode -ge 8) {
        Write-Log "複製檔案到 NAS 時發生錯誤。Robocopy 退出碼：$robocopyExitCode"
        Write-Log "Robocopy 輸出：$robocopyOutput"
    }
    else {
        Write-Log "檔案成功複製到 NAS。Robocopy 退出碼：$robocopyExitCode"
    }
}
catch {
    Write-Log "執行 Robocopy 時發生錯誤：$($_.Exception.Message)"
}

# 複製日誌檔案到 NAS
Write-Log "正在將日誌檔案複製到 NAS..."
try {
    Copy-Item -Path $localLogPath -Destination $nasLogPath -Force
    Write-Log "日誌檔案成功複製到 NAS。"
}
catch {
    Write-Log "複製日誌檔案到 NAS 時發生錯誤：$($_.Exception.Message)"
}

Write-Log "太好了！印表機設定匯出流程完成囉。辛苦啦！"