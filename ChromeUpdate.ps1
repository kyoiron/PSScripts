﻿# 定義基礎路徑
$BASE_NETWORK_PATH = "\\172.29.205.114\Public\sources\audit"
$BASE_MSI_FOLDER = "\\172.29.205.114\loginscript\Update"

# 取得產品名稱
$productName = "Google Chrome"

# 動態定義常數
$NETWORK_LOG_PATH = Join-Path $BASE_NETWORK_PATH $productName
$LOCAL_LOG_PATH = "$env:systemdrive\temp"
$TEMP_DIR = "$env:systemdrive\temp\${productName}Update"
$INSTALL_IF_NOT_INSTALLED = $true

# 根據系統架構選擇合適的 MSI 檔
$msiFileName = if ([Environment]::Is64BitOperatingSystem) {
    "googlechromestandaloneenterprise64.msi"
} else {
    "googlechromestandaloneenterprise.msi"
}
$msiPath = Join-Path $BASE_MSI_FOLDER "Chrome\$msiFileName"

# 取得 MSI 檔詮釋資料
function Get-FileMetaData {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline)][Object] $File,
        [switch] $Signature
    )
    
    Process {
        foreach ($F in $File) {
            if ($F -isnot [System.IO.FileInfo]) {
                if ($F -is [string]) {
                    $F = Get-Item -Path $F
                } else {
                    Write-Log "Get-FileMetaData - 只支援檔案。跳過 $F。"
                    continue
                }
            }

            $shellApplication = New-Object -ComObject Shell.Application
            $shellFolder = $shellApplication.Namespace($F.Directory.FullName)
            $shellFile = $shellFolder.ParseName($F.Name)

            $metaDataObject = [ordered] @{}
            0..400 | ForEach-Object {
                $propertyName = $shellFolder.GetDetailsOf($null, $_)
                if ($propertyName -and $propertyName -notin 'Attributes', 'Folder', 'Type', 'SpaceFree', 'TotalSize', 'SpaceUsed') {
                    $propertyValue = $shellFolder.GetDetailsOf($shellFile, $_)
                    if ($propertyValue) {
                        $metaDataObject[$propertyName] = $propertyValue
                    }
                }
            }

            # 添加基本檔案屬性
            $metaDataObject["Attributes"] = $F.Attributes
            $metaDataObject['IsReadOnly'] = $F.IsReadOnly
            $metaDataObject['IsHidden'] = $F.Attributes -like '*Hidden*'
            $metaDataObject['IsSystem'] = $F.Attributes -like '*System*'
            $metaDataObject['File'] = $F.FullName

            # 如果需要，添加數字簽名信息
            if ($Signature) {
                $digitalSignature = Get-AuthenticodeSignature -FilePath $F.FullName
                $metaDataObject['SignatureCertificateSubject'] = $digitalSignature.SignerCertificate.Subject
                $metaDataObject['SignatureCertificateIssuer'] = $digitalSignature.SignerCertificate.Issuer
                $metaDataObject['SignatureCertificateSerialNumber'] = $digitalSignature.SignerCertificate.SerialNumber
                $metaDataObject['SignatureCertificateNotBefore'] = $digitalSignature.SignerCertificate.NotBefore
                $metaDataObject['SignatureCertificateNotAfter'] = $digitalSignature.SignerCertificate.NotAfter
                $metaDataObject['SignatureCertificateThumbprint'] = $digitalSignature.SignerCertificate.Thumbprint
                $metaDataObject['SignatureStatus'] = $digitalSignature.Status
                $metaDataObject['IsOSBinary'] = $digitalSignature.IsOSBinary
            }

            [PSCustomObject] $metaDataObject
        }
    }
}

$msiMetadata = Get-FileMetaData -File $msiPath
if (-not $msiMetadata) { 
    Write-Error "錯誤: 無法取得 MSI 檔的詮釋資料。"
    exit 1
}

$msiVersion = [version]($msiMetadata.註解 -split "Copyright")[0].Trim()

# 更新 $LOCAL_LOG_FILE 定義，使用 .log 副檔名
$LOCAL_LOG_FILE = Join-Path $LOCAL_LOG_PATH "$env:COMPUTERNAME`_$productName`_$msiVersion.log"
$NETWORK_LOG_FILE = Join-Path $NETWORK_LOG_PATH "$env:COMPUTERNAME`_$productName`_$msiVersion.log"

# 建立紀錄函數
function Write-Log {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    # 直接寫入新的紀錄訊息，覆蓋原有內容
    Set-Content -Path $LOCAL_LOG_FILE -Value $logMessage -Encoding UTF8
    
    Write-Host $logMessage
}

# 上傳紀錄到 NAS 紀錄存放區
function Upload-Log {
    if (Test-Path $LOCAL_LOG_FILE) {
        try {
            if (-not (Test-Path $NETWORK_LOG_PATH)) {
                New-Item -ItemType Directory -Path $NETWORK_LOG_PATH -Force | Out-Null
            }
            Copy-Item -Path $LOCAL_LOG_FILE -Destination $NETWORK_LOG_FILE -Force
            Write-Host "紀錄檔已成功上傳到 NAS 紀錄存放區。"
        } catch {
            Write-Host "警告: 無法上傳紀錄檔到 NAS 紀錄存放區。錯誤: $_"
        }
    } else {
        Write-Host "警告: 本機紀錄檔不存在。"
    }
}

function Get-ProductInstallation {
    param (
        [string]$ProductName
    )
    $regPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    
    $installations = foreach ($path in $regPaths) {
        if (Test-Path $path) {
            Get-ItemProperty $path | Where-Object { $_.DisplayName -eq $ProductName }
        }
    }
    
    $installations | Sort-Object -Property DisplayVersion -Descending | Select-Object -First 1
}

function Install-Product {
    param (
        [string]$MsiPath,
        [string]$ProductName,
        [string]$CustomLogName = ""
    )
    
    $tempLogName = "temp_install_log.txt"
    $tempLogPath = Join-Path $TEMP_DIR $tempLogName
    $arguments = "/i `"$MsiPath`" /qn ALLUSERS=1 /L*VX `"$tempLogPath`""
    
    Write-Log "停止 $ProductName 進程..."
    Stop-Process -Name ($ProductName.Split(" ")[0]).ToLower() -Force -ErrorAction SilentlyContinue
    
    Write-Log "開始安裝 $ProductName..."
    $process = Start-Process "msiexec" -ArgumentList $arguments -Wait -PassThru
    
    if ($process.ExitCode -eq 0) {
        Write-Log "$ProductName 安裝成功完成。"
        
        # 獲取安裝後的實際版本
        $installedProduct = Get-ProductInstallation -ProductName $ProductName
        if ($installedProduct) {
            $installedVersion = $installedProduct.DisplayVersion
            Write-Log "安裝的 $ProductName 版本: $installedVersion"
            
            # 根據實際安裝的版本創建最終的紀錄檔名，使用 .txt 副檔名
            if ([string]::IsNullOrEmpty($CustomLogName)) {
                $finalLogName = "{0}_{1}_{2}.txt" -f $env:COMPUTERNAME, $ProductName.Replace(" ", ""), $installedVersion
            } else {
                $finalLogName = $CustomLogName -replace '(?<=_)[\d.]+(?=\.txt$)', $installedVersion
            }
            
            $finalLogPath = Join-Path $TEMP_DIR $finalLogName
            Rename-Item -Path $tempLogPath -NewName $finalLogName
            
            Write-Log "上傳安裝紀錄到 NAS 紀錄存放區..."
            try {
                $networkLogPath = Join-Path $NETWORK_LOG_PATH $finalLogName
                Copy-Item $finalLogPath $networkLogPath -Force
                Write-Log "安裝紀錄已成功上傳到 NAS 紀錄存放區。"
                
                # 保留本地副本並記錄位置
                Write-Log "安裝紀錄的本機副本保留在: $finalLogPath"
            } catch {
                Write-Log "警告: 無法上傳安裝紀錄到 NAS 紀錄存放區。錯誤: $_"
                Write-Log "安裝紀錄保留在本機: $finalLogPath"
            }
        } else {
            Write-Log "警告: 無法獲取安裝後的產品版本。"
        }
    } else {
        Write-Log "$ProductName 安裝失敗。退出碼: $($process.ExitCode)"
        
        # 即使安裝失敗，也嘗試上傳日誌
        $failedLogName = "{0}_{1}_InstallFailed_{2}.txt" -f $env:COMPUTERNAME, $ProductName.Replace(" ", ""), (Get-Date -Format "yyyyMMdd_HHmmss")
        $failedLogPath = Join-Path $TEMP_DIR $failedLogName
        Rename-Item -Path $tempLogPath -NewName $failedLogName
        
        Write-Log "上傳失敗的安裝紀錄到 NAS 紀錄存放區..."
        try {
            $networkFailedLogPath = Join-Path $NETWORK_LOG_PATH $failedLogName
            Copy-Item $failedLogPath $networkFailedLogPath -Force
            Write-Log "失敗的安裝紀錄已成功上傳到 NAS 紀錄存放區。"
            Write-Log "失敗的安裝紀錄的本機副本保留在: $failedLogPath"
        } catch {
            Write-Log "警告: 無法上傳失敗的安裝紀錄到 NAS 紀錄存放區。錯誤: $_"
            Write-Log "失敗的安裝紀錄保留在本機: $failedLogPath"
        }
    }
    
    # 不再刪除本機紀錄檔，因為我們現在保留它們
}

# 主程序
Write-Log "開始 $productName $msiVersion 更新腳本..."

Write-Log "檢查當前安裝的 $productName 版本..."
$productInstalled = Get-ProductInstallation -ProductName $productName
if (-not $productInstalled -and -not $INSTALL_IF_NOT_INSTALLED) { 
    Write-Log "$productName 未安裝且不需要安裝。"
    Upload-Log
    exit 0 
}

Write-Log "MSI 版本: $msiVersion"
if ($productInstalled) {
    Write-Log "已安裝版本: $($productInstalled.DisplayVersion)"
}

if ($productInstalled -and [version]$productInstalled.DisplayVersion -ge $msiVersion) {
    Write-Log "已安裝的 $productName 版本不低於 MSI 版本。無需更新。"
    Upload-Log
    exit 0
}

Write-Log "準備更新 $productName..."
if (-not (Test-Path $TEMP_DIR)) { 
    Write-Log "建立暫存目錄: $TEMP_DIR"
    New-Item -ItemType Directory -Path $TEMP_DIR -Force | Out-Null 
}

Write-Log "複製 MSI 檔到暫存目錄..."
Copy-Item $msiPath $TEMP_DIR

Write-Log "開始 $productName 安裝/更新過程..."
$customLogName = "${env:COMPUTERNAME}_${productName}_VERSION.txt"
Install-Product -MsiPath (Join-Path $TEMP_DIR (Split-Path $msiPath -Leaf)) -ProductName $productName -CustomLogName $customLogName

Write-Log "清理暫存檔..."
Remove-Item $TEMP_DIR -Recurse -Force

Write-Log "$productName $msiVersion 更新腳本完成。"
Upload-Log