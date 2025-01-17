# 指紋系統設置腳本 (FingerPrintSystemSetup.ps1)

# 自訂變數
$FingerPrint_PC = @("TND-STOF-113","TND-STOF-112","TND-RMSE-141","TND-RMSE-142","TND-GCSE-143","TND-CENTRAL-144","TND-RMSE-028","TND-RMSE-172","TND-RMSE-049","TND-RMSE-043","TND-RMSE-047","TND-STOF-117")
$pfxFilePath = "\\172.29.205.114\loginscript\Update\PIDDLL\NeoFaceCert.pfx"
$thumbprint = "9E8F00E882B44F9C955B555B2D7FA4EB3FBA94F3"
$password = ConvertTo-SecureString "1qaz@WSX" -AsPlainText -Force
$certPathPersonalLocal = "Cert:\LocalMachine\My"
$certPathPersonalUser = "Cert:\CurrentUser\My"
$certPathRoot = "Cert:\LocalMachine\Root"
$NAS_Path = "\\172.29.205.114\loginscript\Update\PIDDLL"
$PC_Path = "D:\"
$folders = @("CPID_Client", "PIDDLLV2")

# 更新紀錄檔路徑，包含電腦名稱
$logFileName = "$env:COMPUTERNAME`_FingerPrintSystemSetup_LocalMachine.log"
$logPath = "$env:SystemDrive\temp\$logFileName"
$remoteLogPath = "\\172.29.205.114\Public\sources\audit\PIDDLL\$logFileName"

# 清空現有的日誌檔案
"" | Set-Content -Path $logPath
"" | Set-Content -Path $remoteLogPath

# 函數：寫入日誌
function Write-Log {
    param (
        [string]$Message
    )
    $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$logTimestamp - $Message"
    Add-Content -Path $logPath -Value $logMessage
    Add-Content -Path $remoteLogPath -Value $logMessage
    Write-Host $logMessage
}

# 函數：檢查憑證安裝狀態
function Check-CertificateInstallation {
    param (
        [string]$CertPath,
        [string]$StoreName
    )
    
    $cert = Get-ChildItem -Path $CertPath | Where-Object { $_.Thumbprint -eq $thumbprint }
    if ($cert) {
        Write-Log "憑證檢查結果: 已安裝在 $StoreName 存放區"
        Write-Log "憑證詳細資訊:"
        Write-Log "主體: $($cert.Subject)"
        Write-Log "發行者: $($cert.Issuer)"
        Write-Log "有效期自: $($cert.NotBefore)"
        Write-Log "有效期至: $($cert.NotAfter)"
        Write-Log "Thumbprint: $($cert.Thumbprint)"
        return $true
    } else {
        Write-Log "憑證檢查結果: 未在 $StoreName 存放區找到指定的憑證"
        return $false
    }
}

# 函數：檢查並匯入憑證
function Import-CertificateIfNeeded {
    param (
        [string]$CertPath,
        [string]$StoreName
    )
    
    if (-not (Check-CertificateInstallation -CertPath $CertPath -StoreName $StoreName)) {
        try {
            $certificate = Import-PfxCertificate -FilePath $pfxFilePath -Password $password -CertStoreLocation $CertPath -Exportable
            Write-Log "憑證已成功匯入到 $StoreName 存放區。"
            Write-Log "憑證驗證成功。找到匹配的 thumbprint: $thumbprint"
            Write-Log "匯入的憑證詳細資訊:"
            Write-Log "主體: $($certificate.Subject)"
            Write-Log "發行者: $($certificate.Issuer)"
            Write-Log "有效期自: $($certificate.NotBefore)"
            Write-Log "有效期至: $($certificate.NotAfter)"
            Write-Log "Thumbprint: $($certificate.Thumbprint)"
        }
        catch {
            Write-Log "匯入憑證到 $StoreName 存放區時發生錯誤: $_"
        }
    }
    else {
        Write-Log "憑證已存在於 $StoreName 存放區，無需重新匯入"
    }
}

# 主程序
try {
    Write-Log "開始執行腳本於指紋系統電腦: $env:COMPUTERNAME"
    if ($FingerPrint_PC -contains $env:COMPUTERNAME) {
        # 檢查並匯入憑證
        Import-CertificateIfNeeded -CertPath $certPathPersonalLocal -StoreName "本機個人"
        Import-CertificateIfNeeded -CertPath $certPathRoot -StoreName "受信任的根憑證授權單位"

        # 執行 IB_Driver 腳本
        try {
            & "$env:SystemDrive\temp\IB_Driver.ps1"
            Write-Log "成功執行 IB_Driver.ps1"
        }
        catch {
            Write-Log "執行 IB_Driver.ps1 時發生錯誤: $_"
        }

        # 複製 CPID_Client 和 PIDDLLV2
        foreach ($folder in $folders) {
            $sourcePath = Join-Path $NAS_Path $folder
            $destinationPath = Join-Path $PC_Path $folder
            
            Write-Log "開始同步資料夾: $folder"
            Write-Log "來源路徑: $sourcePath"
            Write-Log "目標路徑: $destinationPath"
        
            if (Test-Path $sourcePath) {
                # 執行 robocopy 並直接捕獲輸出
                $robocopyArgs = @(
                    $sourcePath
                    $destinationPath
                    "/E"      # 複製所有子目錄（包括空目錄）
                    "/XO"     # 排除目標位置中較新的檔案
                    "/XX"     # 排除額外的檔案
                    "/NP"     # 不顯示進度
                    "/NS"     # 不顯示檔案大小
                    "/NC"     # 不顯示檔案類別
                    "/NDL"    # 不顯示目錄名稱
                    "/njh"    # 不顯示工作標頭
                    "/njs"    # 不顯示工作摘要
                )
                
                $output = & robocopy $robocopyArgs | Where-Object { $_ -ne "" }
                
                # 記錄複製的檔案
                if ($output) {
                    Write-Log "複製的檔案清單："
                    foreach ($line in $output) {
                        Write-Log "  $line"
                    }
                } else {
                    Write-Log "沒有需要複製的新檔案"
                }
        
                # 檢查 robocopy 的返回碼並記錄結果
                $exitCode = $LASTEXITCODE
                switch ($exitCode) {
                    0 { Write-Log "資料夾 $folder 同步完成：沒有檔案需要複製" }
                    1 { Write-Log "資料夾 $folder 同步完成：已成功複製一個或多個檔案" }
                    2 { Write-Log "資料夾 $folder 同步完成：發現一些多餘的檔案或資料夾，但未執行刪除" }
                    { $_ -ge 8 } { Write-Log "資料夾 $folder 同步時發生錯誤，錯誤碼：$exitCode" }
                    default { Write-Log "資料夾 $folder 同步完成，返回碼：$exitCode" }
                }
            }
            else {
                Write-Log "錯誤：來源資料夾不存在 - $sourcePath"
            }
        
            Write-Log "完成資料夾同步: $folder"
            Write-Log "----------------------------------------"
        }

        # 建立桌面捷徑
        $shortcutPath = "$env:PUBLIC\Desktop\指紋系統.lnk"
        if (-not (Test-Path $shortcutPath) -and (Test-Path "D:\PIDDLLV2\PIDDLL.exe")) {
            try {
                $shell = New-Object -ComObject WScript.Shell
                $shortcut = $shell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = "D:\PIDDLLV2\PIDDLL.exe"
                $shortcut.Save()
                Write-Log "成功建立指紋系統桌面捷徑"
            }
            catch {
                Write-Log "建立指紋系統桌面捷徑時發生錯誤: $_"
            }
        }
        else {
            Write-Log "指紋系統桌面捷徑已存在或目標程序不存在"
        }
    }
    else {
        Write-Log "此電腦不在指紋系統設置清單中，跳過設置"
    }
}
catch {
    Write-Log "指紋系統設置過程中發生未預期的錯誤: $_"
}
finally {
    Write-Log "腳本執行完成"
}