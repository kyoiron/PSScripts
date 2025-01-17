# 變數定義
$fingerPrintPC = @("TND-STOF-113","TND-STOF-112","TND-RMSE-141","TND-RMSE-142","TND-GCSE-143","TND-GCSE-144","TND-GCSE-145")
$pfxFilePath = "\\172.29.205.114\loginscript\Update\PIDDLL\NeoFaceCert.pfx"
$thumbprint = "9E8F00E882B44F9C955B555B2D7FA4EB3FBA94F3"
$password = ConvertTo-SecureString "1qaz@WSX" -AsPlainText -Force
$certPathPersonal = "Cert:\CurrentUser\My"
$logContent = ""
$logFileName = "$env:COMPUTERNAME-CertificateImportResult.log"
$localLogPath = "$env:SystemDrive\temp\$logFileName"
$networkLogPath = "\\172.29.205.114\Public\sources\audit\PIDDLL\$logFileName"

# 函數定義
function Write-Log {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    $script:logContent += "$logMessage`n"
    Write-Host $logMessage
}

# 主要程式邏輯
if ($fingerPrintPC -contains $env:COMPUTERNAME) {
    try {
        $isInstalledCertificatePersonal = Get-ChildItem -Path $certPathPersonal | Where-Object { $_.Thumbprint -eq $thumbprint }
        
        if (-not $isInstalledCertificatePersonal) {
            $certificatePersonal = Import-PfxCertificate -FilePath $pfxFilePath -Password $password -CertStoreLocation $certPathPersonal -Exportable
            Write-Log "Certificate imported successfully."
        } else {
            Write-Log "Certificate already exists."
        }
    } catch {
        Write-Log "An error occurred: $_"
    }
} else {
    Write-Log "This computer is not in the list of finger print PCs."
}

# 保存日誌
$logContent | Out-File -FilePath $localLogPath -Encoding UTF8
Copy-Item -Path $localLogPath -Destination $networkLogPath -Force

Write-Log "Log file saved locally at $localLogPath and copied to $networkLogPath"