# 同電腦名稱印表機匯入作業

$ImportPrinterPC = @("TND-GASE-051")
$PrinterExportFileLocation = "\\172.29.205.114\mig\Printer"
$BackupLocation = "\\172.29.205.114\mig\Printer_BACKUP"
$LogFolder = "$env:SystemDrive\temp"
$NetworkLogFolder = "\\172.29.205.114\Public\sources\audit\Printer_LOG"

# 創建日誌函數
function Write-Log {
    param(
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
}


    $Date = Get-Date -Format "yyyyMMdd"
    $LogFileName = "$env:COMPUTERNAME-$Date.log"
    $LogFile = Join-Path $LogFolder $LogFileName
    $NetworkLogFile = Join-Path $NetworkLogFolder $LogFileName

    # 檢查網路位置是否已存在日誌文件
    if (Test-Path $NetworkLogFile) {
        Write-Host "日誌文件 $NetworkLogFile 已存在，表示此操作已執行過。腳本將退出。"
        exit
    }

    $FileName = "${env:COMPUTERNAME}x64.printerExport"
    $SourceFile = Join-Path $PrinterExportFileLocation $FileName
    $TempFile = Join-Path $env:TEMP $FileName

    Write-Log "開始執行印表機匯入作業"

    try {
        Write-Log "正在複製文件到臨時目錄"
        Copy-Item -Path $SourceFile -Destination $TempFile -Force
        Write-Log "文件複製完成"

        Write-Log "開始匯入印表機設定"
        $PrintbrmPath = Join-Path $env:SystemRoot "system32\Spool\Tools\Printbrm.exe"
        $PrintbrmOutput = & $PrintbrmPath -R -F "$TempFile" 2>&1
        Write-Log "Printbrm.exe 輸出: $PrintbrmOutput"

        Write-Log "設置印表機權限"
        $DefaultPermission = (Get-Printer -Name 'Microsoft Print to PDF' -Full).PermissionSDDL
        Get-Printer | ForEach-Object { 
            Set-Printer -Name $_.Name -PermissionSDDL $DefaultPermission 
            Write-Log "已設置 $($_.Name) 的權限"
        }

        Write-Log "移動原始文件到備份位置"
        Move-Item -Path $SourceFile -Destination $BackupLocation -Force
        Write-Log "文件備份完成"

        Write-Log "印表機匯入作業完成"
    }
    catch {
        Write-Log "發生錯誤: $_"
    }
    finally {
        if (Test-Path $TempFile) {
            Remove-Item $TempFile -Force
            Write-Log "臨時文件已刪除"
        }

        # 移動日誌文件到網絡位置
        Move-Item -Path $LogFile -Destination $NetworkLogFile -Force
        Write-Host "日誌文件已移動到 $NetworkLogFile"
    }