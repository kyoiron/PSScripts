# 設定變數
$fingerPrintPC = @("TND-STOF-113","TND-STOF-112","TND-RMSE-141","TND-RMSE-142","TND-GCSE-143","TND-CENTRAL-144","TND-GCSE-145","TND-RMSE-028","TND-RMSE-172","TND-RMSE-049","TND-RMSE-043","TND-WOME-142","TND-RMSE-047")
$certPath = "\\172.29.205.114\loginscript\Update\PIDDLL\NeoFaceCert.pfx"
$certPass = ConvertTo-SecureString -String "1qaz@WSX" -Force -AsPlainText
$certStore = "Cert:\CurrentUser\My"
$expectedThumbprint = "9E8F00E882B44F9C955B555B2D7FA4EB3FBA94F3"
$logFileName = "$env:COMPUTERNAME`_FingerprintSystemSetup_CurrentUser.log"
$localLogPath = "$env:SystemDrive\temp\$logFileName"
$networkLogPath = "\\172.29.205.114\Public\sources\audit\PIDDLL\$logFileName"
$sourcePath = "\\172.29.205.114\loginscript\Update\PIDDLL"
$destinationPath = "D:\"
$shortcutName = "指紋系統"
$shortcutTarget = "D:\PIDDLLV2\PIDDLL.exe"

# 檢查當前電腦名稱
if ($fingerPrintPC -notcontains $env:COMPUTERNAME) {
    Write-Host "此腳本不適用於當前電腦 ($env:COMPUTERNAME)。目標指紋系統電腦為 $($fingerPrintPC -join ', ')。"
    exit
}

# 函數：寫入日誌
function Write-Log {
    param ([string]$Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Add-Content -Path $localLogPath -Value $logMessage
    Add-Content -Path $networkLogPath -Value $logMessage
    Write-Host $logMessage
}

# 函數：清理舊的日誌檔案
function Clean-OldLogs {
    # 清理本機舊日誌
    if (Test-Path $localLogPath) {
        Remove-Item -Path $localLogPath -Force
    }
    # 清理網絡共享上的舊日誌
    if (Test-Path $networkLogPath) {
        Remove-Item -Path $networkLogPath -Force
    }
}

# 函數：檢查和建立資料夾
function Check-And-Create-Folders {
    $folders = @("CPID_Client", "PIDDLLV2")
    foreach ($folder in $folders) {
        $path = Join-Path -Path $destinationPath -ChildPath $folder
        if (!(Test-Path $path)) {
            Write-Log "建立資料夾: $path"
            New-Item -ItemType Directory -Path $path | Out-Null
            Copy-Item -Path "$sourcePath\*" -Destination $path -Recurse -Force
            Write-Log "從 $sourcePath 複製檔案到 $path"
        } else {
            Write-Log "資料夾已存在: $path"
        }
    }
}

# 函數：建立或更新桌面捷徑
function Create-Shortcut {
    $shortcutPath = "$env:USERPROFILE\Desktop\$shortcutName.lnk"
    
    # 檢查捷徑是否已存在
    if (Test-Path $shortcutPath) {
        Write-Log "桌面捷徑已存在: $shortcutName"
    } else {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = $shortcutTarget
        $Shortcut.Save()
        Write-Log "建立新的桌面捷徑: $shortcutName"
    }
}

# 清理舊的日誌檔案
Clean-OldLogs

# 建立本機日誌目錄（如果不存在）
if (!(Test-Path "$env:SystemDrive\temp")) {
    New-Item -ItemType Directory -Path "$env:SystemDrive\temp" | Out-Null
}

# 記錄執行本程式的使用者帳號
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
Write-Log "程式執行使用者: $currentUser"
Write-Log "開始執行腳本於指紋系統電腦: $env:COMPUTERNAME"

# 檢查和建立資料夾
Check-And-Create-Folders

# 建立桌面捷徑
#Create-Shortcut

# 匯入憑證
try {
    $importedCert = Import-PfxCertificate -FilePath $certPath -CertStoreLocation $certStore -Password $certPass -Exportable
    Write-Log "憑證已成功匯入。"
} catch {
    Write-Log "匯入憑證時發生錯誤: $_"
    exit
}

# 驗證憑證是否正確匯入
$cert = Get-ChildItem -Path $certStore | Where-Object { $_.Thumbprint -eq $expectedThumbprint }

if ($cert) {
    Write-Log "憑證驗證成功。找到匹配的 thumbprint: $($cert.Thumbprint)"
} else {
    Write-Log "憑證驗證失敗。未找到具有 thumbprint $expectedThumbprint 的憑證。"
}

# 顯示匯入的憑證詳細資訊
if ($importedCert) {
    Write-Log "匯入的憑證詳細資訊:"
    Write-Log "主體: $($importedCert.Subject)"
    Write-Log "發行者: $($importedCert.Issuer)"
    Write-Log "有效期自: $($importedCert.NotBefore)"
    Write-Log "有效期至: $($importedCert.NotAfter)"
    Write-Log "Thumbprint: $($importedCert.Thumbprint)"
}

Write-Log "腳本執行完成"