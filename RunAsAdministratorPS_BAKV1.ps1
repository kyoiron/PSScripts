# 設置紀錄文件路徑
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$localLogPath = "$env:SystemDrive\temp\$env:ComputerName-RunAsAdministratorPS_$timestamp.log"
$networkBasePath = "\\172.29.205.114\Public\sources\audit"
$networkLogFolder = Join-Path $networkBasePath "RunAsAdministratorPS"
$networkLogPath = Join-Path $networkLogFolder "$env:ComputerName-RunAsAdministratorPS_$timestamp.log"
# 設置保留紀錄檔的天數
$LogRetentionDays = 30

# 函數：寫入紀錄
function Write-Log {
    param (
        [string]$Message
    )
    $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$logTimestamp] $Message"
    Add-Content -Path $localLogPath -Value $logMessage
    Write-Host $logMessage
}
# 函數：清理舊的紀錄檔
function Clean-OldLogs {
    param (
        [string]$LogFolder,
        [int]$RetentionDays
    )
    
    $cutOffDate = (Get-Date).AddDays(-$RetentionDays)
    
    Get-ChildItem -Path $LogFolder -Filter "*RunAsAdministratorPS*.log" | 
    Where-Object { $_.LastWriteTime -lt $cutOffDate } | 
    ForEach-Object {
        Write-Log "正在刪除舊的紀錄檔: $($_.FullName)"
        Remove-Item $_.FullName -Force
    }
}
# 函數：確保資料夾存在
function Ensure-FolderExists {
    param (
        [string]$FolderPath
    )
    if (!(Test-Path -Path $FolderPath)) {
        try {
            New-Item -Path $FolderPath -ItemType Directory -Force -ErrorAction Stop
            Write-Log "已建立資料夾: $FolderPath"
        }
        catch {
            Write-Log "無法建立資料夾 $FolderPath : $_"
            return $false
        }
    }
    return $true
}

# 函數：獲取 Symantec Endpoint Protection 路徑
function Get-SymantecPath {
    $paths = @("${env:ProgramFiles(x86)}\Symantec\Symantec Endpoint Protection\Smc.exe",
               "${env:ProgramFiles}\Symantec\Symantec Endpoint Protection\Smc.exe")
    foreach ($path in $paths) {
        if (Test-Path $path) { return $path }
    }
    Write-Log "找不到 Symantec Endpoint Protection。將跳過 Symantec 相關操作。"
    return $null
}

# 函數：解碼 Symantec 密碼
function Get-DecodedSymantecPassword {
    $encodedPassword = "c3ltYW50ZWM="
    $decodedBytes = [System.Convert]::FromBase64String($encodedPassword)
    return [System.Text.Encoding]::UTF8.GetString($decodedBytes)
}

# 函數：獲取 Bitlocker 恢復金鑰
function Get-BitlockerRecoveryKey {
    $encodedKey = "MTAyMTU3LTQwODg3MC00NTU3MzAtNDYzMTU1LTE0OTM1OC0yOTY5NTYtNzAwNzExLTA0NTU3Mw=="
    $decodedBytes = [System.Convert]::FromBase64String($encodedKey)
    return [System.Text.Encoding]::UTF8.GetString($decodedBytes)
}

# 函數：執行更新腳本並處理錯誤
function Execute-UpdateScript {
    param (
        [string]$ScriptPath,
        [int]$Timeout = 180  # 預設超時時間為3分鐘
    )

    Write-Log "開始執行更新腳本: $ScriptPath"
    
    $job = Start-Job -ScriptBlock {
        param($path)
        & $path
    } -ArgumentList $ScriptPath

    $completed = $job | Wait-Job -Timeout $Timeout
    if ($completed -eq $null) {
        Write-Log "警告: 腳本 $ScriptPath 執行超時（${Timeout}秒）"
        
        # 檢查是否有正在運行的 msiexec 進程
        $msiProcesses = Get-Process | Where-Object { $_.Name -eq "msiexec" }
        if ($msiProcesses) {
            Write-Log "檢測到正在運行的 MSI 安裝進程。不會強制終止這些進程，但會停止 PowerShell 作業。"
            foreach ($proc in $msiProcesses) {
                Write-Log "正在運行的 MSI 進程 ID: $($proc.Id)"
            }
        }

        Stop-Job $job
        Write-Log "PowerShell 作業已停止，但 MSI 安裝可能仍在進行。"
    }
    else {
        $result = Receive-Job $job
        Write-Log "腳本 $ScriptPath 執行完成。輸出: $result"
    }
    Remove-Job $job -Force

    # 檢查是否有殘留的 msiexec 進程
    $remainingMsiProcesses = Get-Process | Where-Object { $_.Name -eq "msiexec" }
    if ($remainingMsiProcesses) {
        Write-Log "警告: 檢測到仍在運行的 MSI 安裝進程。這些進程將繼續在背景運行。"
        foreach ($proc in $remainingMsiProcesses) {
            Write-Log "運行中的 MSI 進程 ID: $($proc.Id)"
        }
    }

    # 立即更新網絡日誌
    Update-NetworkLog
}

# 函數：更新網絡日誌
function Update-NetworkLog {
    if (Ensure-FolderExists -FolderPath $networkLogFolder) {
        try {
            Copy-Item -Path $localLogPath -Destination $networkLogPath -Force -ErrorAction Stop
            Write-Log "紀錄已更新到網絡位置: $networkLogPath"
        }
        catch {
            Write-Log "無法更新紀錄到網絡位置: $_"
        }
    }
}

# 主要執行邏輯
try {
    Write-Log "開始執行 RunAsAdministratorPS 腳本"

    # 清理本機舊紀錄檔
    Clean-OldLogs -LogFolder "$env:SystemDrive\temp" -RetentionDays $LogRetentionDays
    
    # 清理網路位置舊紀錄檔
    if (Test-Path $networkLogFolder) {
        Clean-OldLogs -LogFolder $networkLogFolder -RetentionDays $LogRetentionDays
    }

    # 強制對時
    $NTPServer = "172.29.204.63"
    Write-Log "正在進行時間同步，NTP服務器: $NTPServer"
    w32tm /config /manualpeerlist:"$NTPServer" /syncfromflags:manual /update | Out-Null
    Write-Log "時間同步完成"

    # 將tnduser加入本機管理者帳號
    Write-Log "正在將tnduser加入本機管理者群組"
    Add-LocalGroupMember -Group "Administrators" -Member "$env:COMPUTERNAME\tnduser" -ErrorAction SilentlyContinue
    Write-Log "tnduser已加入本機管理者群組"

    # 啟用tndadmin帳戶
    Write-Log "正在啟用tndadmin帳戶"
    net user "tndadmin" /active:yes
    Write-Log "tndadmin帳戶已啟用"

    # 獲取tnduser狀態並保存
    Write-Log "正在獲取tnduser狀態"
    $tnduserStatusPath = "$env:SystemDrive\temp\${env:computername}_tnduserStatus.txt"
    Get-LocalUser -Name tnduser | Select-Object * | Out-File $tnduserStatusPath
    if (Test-Path $tnduserStatusPath) {
        Copy-Item $tnduserStatusPath -Destination "$networkBasePath\tnduser" -Force
        Write-Log "tnduser狀態已保存並複製到網絡位置"
    }

    # 解鎖Bitlocker槽
    $BitlockerRecoveryKey = Get-BitlockerRecoveryKey
    if ($env:BitlockerDataDrive -ne $null) {
        Write-Log "正在解鎖Bitlocker槽: $env:BitlockerDataDrive"
        manage-bde -unlock $env:BitlockerDataDrive -RecoveryPassword $BitlockerRecoveryKey
        Write-Log "Bitlocker槽解鎖完成"
    }
    $BitlockerRecoveryKey = $null  # 清除金鑰變數

    # 檢查並啟用UAC
    Write-Log "正在檢查UAC狀態"
    if ((Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA -eq 0) {
        Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name EnableLUA -Value 1
        Write-Log "UAC已啟用"
    } else {
        Write-Log "UAC已處於啟用狀態"
    }
    #筆硯簽章工具及列印工具安裝程式
        $EicPC_Reinstall_starup=@("TND-BUSE-072")
        if($EicPC_Reinstall_starup -contains $env:COMPUTERNAME){
            Write-Log "進行列印工具重新安裝(啟動區版)"
            & $env:SystemDrive\temp\eic.ps1 -ForceReinstall:$true -InstallNonStartup:$false -InstallScope Both
            Write-Log "完成印工具重新安裝(啟動區版)"
        } 
        $EicPC_Reinstall_nonstarup=@()
        if($EicPC_Reinstall_nonstarup -contains $env:COMPUTERNAME){
            Write-Log "進行列印工具重新安裝"
            & $env:SystemDrive\temp\eic.ps1 -ForceReinstall:$true -InstallScope EicPrint
            Write-Log "完成印工具重新安裝"
        } 
    #指紋機
        # 在主腳本的適當位置添加以下代碼

        Write-Log "開始執行指紋系統設置"
        $fingerPrintSetupScript = "$env:SystemDrive\temp\FingerPrintSystemSetup.ps1"
        if (Test-Path $fingerPrintSetupScript) {
            try {
                & $fingerPrintSetupScript
                Write-Log "指紋系統設置腳本執行完成"
            }
            catch {
                Write-Log "執行指紋系統設置腳本時發生錯誤: $_"
            }
        }
        else {
            Write-Log "找不到指紋系統設置腳本: $fingerPrintSetupScript"
        }

    # 重新匯入印表機
        $ReImportPrinterPC = @("TND-STOF-113","TND-ARCH-082","TND-ARCH-081","TND-1EES-083","TND-3EES-084","TND-2EES-088","TND-GCSE-034","TND-GCSE-073","TND-ASSE-032","TND-ACOF-040","TND-ACOF-020")
        if ($ReImportPrinterPC -contains $env:COMPUTERNAME) {& "$env:SystemDrive\temp\ReImportPrinter.ps1"}
    # 印表機更名
    Write-Log "正在更新印表機名稱"
    Get-printer | Where-Object { $_.Name -like ("*" + [regex]::escape(']')) } | 
    ForEach-Object { 
        $newName = ($_.Name -replace [regex]::escape('['), '【') -replace [regex]::escape(']'), '】'
        Rename-Printer -name $_.Name -NewName $newName
        Write-Log "印表機已更名: $($_.Name) -> $newName"
    }

    # 印表機備份程式
    Write-Log "正在執行印表機備份程序"
    & "$env:SystemDrive\temp\PrinterBackup.ps1"
    Write-Log "印表機備份完成"

    # 印表機權限設置
    Write-Log "正在設置印表機權限"
    $Temp_PermissionSDDL = "G:SYD:(A;;LCSWSDRCWDWO;;;WD)(A;OIIO;RPWPSDRCWDWO;;;WD)(A;;SWRC;;;AC)(A;CIIO;RC;;;AC)(A;OIIO;RPWPSDRCWDWO;;;AC)(A;;LCSWSDRCWDWO;;;CO)(A;OIIO;RPWPSDRCWDWO;;;CO)(A;OIIO;RPWPSDRCWDWO;;;BA)(A;;LCSWSDRCWDWO;;;BA)"
    Get-printer | ForEach-Object { 
        Set-Printer $_.Name -PermissionSDDL $Temp_PermissionSDDL
        Write-Log "已設置印表機權限: $($_.Name)"
    }

    # 檢查並啟動跨平台工具
    Write-Log "正在檢查跨平台工具"
    $ChkSrv = Get-Process -Name ChkSrv -ErrorAction SilentlyContinue
    if ((!$ChkSrv) -and (Test-Path -Path "${env:ProgramFiles(x86)}\HiPKILocalSignServer\ChkSrv.exe")) {
        Start-Process -FilePath "${env:ProgramFiles(x86)}\HiPKILocalSignServer\ChkSrv.exe"
        Write-Log "已啟動跨平台工具"
    } else {
        Write-Log "跨平台工具已在運行或未安裝"
    }

    # 列出所有正在運行的程序
    <#
    Write-Log "正在列出所有正在運行的程序"
    $runningProcesses = Get-Process | Select-Object ProcessName, Id, CPU, WorkingSet, Description
    $runningProcessesLog = $runningProcesses | Format-Table -AutoSize | Out-String
    Write-Log "正在運行的程序列表:"
    Write-Log $runningProcessesLog
    #>
    
    # 執行各種更新腳本
    $updateScripts = @(
        "ChromeUpdate.ps1",
        "AdobeReaderUpdateV2.ps1",
        "7zipUpdateV3.ps1",
        "GeasBatchsign.ps1",
        "MariadbConnectorOdbc.ps1",
        "ThreatSonarPCv3.ps1",
        "WM7AssetCluster.ps1",
        "FileZillaUpdate.ps1",
        "XnViewUpdate.ps1",
        "K-LiteMegaCodecPackUpdate.ps1",
        "PDFXChangeEditorUpdate.ps1",
        "CMEXFontClientUpdate.ps1",
        "CheckBoshiamyTIP.ps1",
        "CheckGoing.ps1",
        "CheckLINE.ps1"
    )

    foreach ($script in $updateScripts) {
        $scriptPath = "$env:SystemDrive\temp\$script"
        if (Test-Path $scriptPath) {
            Execute-UpdateScript -ScriptPath $scriptPath
        } else {
            Write-Log "找不到更新腳本: $script"
        }
        # 每執行完一個腳本後更新網絡日誌
        Update-NetworkLog
    }

    # 確認SEP啟動並進行防毒更新
    $SymantecPath = Get-SymantecPath
    if (!(Get-Process -Name ccSvcHst -ErrorAction SilentlyContinue)) {
        Write-Log "正在啟動Symantec Endpoint Protection"
        $symantecPassword = Get-DecodedSymantecPassword
        Start-Process -FilePath $SymantecPath -ArgumentList "-p `"$symantecPassword`" -start" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
        $symantecPassword = $null
        Write-Log "Symantec Endpoint Protection已啟動"
    }

    # 進行Windows Update
    Write-Log "正在檢查PSWindowsUpdate模組"
    if ((Get-Module -ListAvailable -Name PSWindowsUpdate) -eq $null) {
        Write-Log "正在安裝PSWindowsUpdate模組"
        $PSWindowsUpdate_Path = "\\172.29.205.114\loginscript\PSWindowsUpdate"
        $PSModule_Path = "$Env:ProgramFiles\WindowsPowerShell\Modules\PSWindowsUpdate"    
        robocopy $PSWindowsUpdate_Path $PSModule_Path "/e /XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        Import-Module PSWindowsUpdate
        Write-Log "PSWindowsUpdate模組安裝完成"
    }

    $LogPath = "$networkBasePath\WSUS"
    $temp = "$env:SystemDrive\temp"
    $ServiceID_WSUS = (Get-WUServiceManager | Where-Object { $_.Name -like "Windows Server Update Service" }).ServiceID

    Write-Log "正在獲取Windows更新歷史紀錄"
    Get-WUHistory | Format-Table -AutoSize | Out-File "$temp\${env:computername}_WindowsUpdate_History.txt" -Force
    Robocopy $temp $LogPath "${env:computername}_WindowsUpdate.txt" "${env:computername}_WindowsUpdate_History.txt" " /XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    Write-Log "Windows更新歷史紀錄已保存"

    # 移除Java
    Write-Log "正在執行Java移除腳本"
    & "$env:SystemDrive\temp\UninstallJAVA.ps1"
    Write-Log "Java移除腳本執行完成"
    
    
    # 進行Win1021H2評估
    <#
    Write-Log "正在執行Win10_21H2評估腳本"
    if($env:computername -ne "TND-STOF-113"){& "$env:SystemDrive\temp\\CompactWin10_21H2.ps1"}
    Write-Log "Win10_21H2評估腳本執行完成"
    #>
    
    # 進行Win1021H2評估
    Write-Log "正在執行Win10_21H2評估及執行升級安裝腳本"
    & "$env:SystemDrive\temp\Windows10-2021Upgrade.ps1"
    Write-Log "Win10_21H2評估及執行升級安裝腳本執行完成"
    
    Write-Log "RunAsAdministratorPS 腳本執行完成"
}
catch {
    Write-Log "發生錯誤: $_"
}
finally {
    # 最後一次更新網絡日誌
    Update-NetworkLog
    Write-Log "本機紀錄檔位置: $localLogPath"
    Write-Log "網路紀錄檔位置: $networkLogPath"
    Write-Log "本次執行的紀錄檔已保留。舊紀錄檔（超過 $LogRetentionDays 天）已被清理。"
}