# 配置
$JavaExcludePC = @("TND-GASE-061")
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$Global:PatternJava32 = "Java ([0-9]|[0-9][0-9]) Update ([0-9]|[0-9][0-9]|[0-9][0-9][0-9])$"
$Global:PatternJava64 = "Java ([0-9]|[0-9][0-9]) Update ([0-9]|[0-9][0-9]|[0-9][0-9][0-9]) \(64-bit\)$" 
$RegUninstallPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")

# 取得當前日期時間
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"

# 設置本機暫存紀錄檔目錄
$localTempPath = "$env:systemdrive\temp"
if (-not (Test-Path $localTempPath)) { 
    New-Item -ItemType Directory $localTempPath -Force | Out-Null 
}

# 設置紀錄檔
$transcriptLogFile = Join-Path $localTempPath "$env:COMPUTERNAME`_Java移除_系統紀錄_$currentDateTime.txt"
$customLogFile = Join-Path $localTempPath "$env:COMPUTERNAME`_Java移除_自定義紀錄_$currentDateTime.txt"

# 啟動系統記錄
Start-Transcript -Path $transcriptLogFile -Force

# 輔助函數
function Write-CustomLog {
    param (
        [string]$Message
    )
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Write-Host $logMessage
    Add-Content -Path $customLogFile -Value $logMessage
}

function Get-JavaInstalls {
    param (
        [string]$Pattern
    )
    $installedJava = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $installedJava += Get-ItemProperty $Path | Where-Object{$_.DisplayName -match $Pattern}
        }
    }
    return $installedJava | Sort-Object -Property Version -Descending
}

function Uninstall-Java {
    param (
        [array]$JavaInstalls
    )
    foreach($install in $JavaInstalls) {
        $uninstall = ($install.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()
        $individualLogFile = "$localTempPath\${env:Computername}_$($install.DisplayName)_移除_$currentDateTime.txt"
        Write-CustomLog "正在移除 $($install.DisplayName)"
        Start-Process "msiexec.exe" -ArgumentList "/X $uninstall /quiet /l*vx ""$individualLogFile""" -Wait
        Write-CustomLog "$($install.DisplayName) 移除完成"
        
        # 將個別移除日誌內容加入到自定義紀錄檔中
        Write-CustomLog "-------------------- $($install.DisplayName) 移除日誌 --------------------"
        Get-Content $individualLogFile | ForEach-Object { Write-CustomLog $_ }
        Write-CustomLog "-------------------- 移除日誌結束 --------------------"
    }
}

function Get-CMEXFontClientVersion {
    $cmexClient = Get-ItemProperty $RegUninstallPaths | Where-Object { $_.DisplayName -like "中推會用戶端更新工具*" } | Select-Object -First 1
    if ($cmexClient) {
        return [version]$cmexClient.DisplayVersion
    }
    return $null
}

# 主要執行
Write-CustomLog "開始執行Java移除腳本"

if($JavaExcludePC.Contains($env:Computername)) { 
    Write-CustomLog "此電腦被排除在Java移除操作之外。"
    Stop-Transcript
    exit 
}

# 檢查中推會用戶端更新工具版本
$cmexVersion = Get-CMEXFontClientVersion
if ($cmexVersion -eq $null) {
    Write-CustomLog "未找到中推會用戶端更新工具。Java移除操作將不會執行。"
    Stop-Transcript
    exit
}

if ($cmexVersion -lt [version]"2.7.1.0") {
    Write-CustomLog "中推會用戶端更新工具版本 ($cmexVersion) 小於 2.7.1.0。Java移除操作將不會執行。"
    Stop-Transcript
    exit
}

Write-CustomLog "中推會用戶端更新工具版本 ($cmexVersion) 大於或等於 2.7.1.0。繼續執行..."

# 獲取所有已安裝的Java
$Java_32_installeds = Get-JavaInstalls -Pattern $Global:PatternJava32
$Java_64_installeds = Get-JavaInstalls -Pattern $Global:PatternJava64

$totalJavaInstalls = $Java_32_installeds.Count + $Java_64_installeds.Count

if ($totalJavaInstalls -eq 0) {
    Write-CustomLog "未檢測到已安裝的Java。無需進行移除操作。"
    Stop-Transcript
    
    # 紀錄檔管理
    $Log_Folder_Path = Join-Path $Log_Path "Java"
    if(!(Test-Path -Path $Log_Folder_Path)) { 
        New-Item -ItemType Directory -Path $Log_Folder_Path -Force | Out-Null
    }

    Write-Host "正在將紀錄檔複製到網路儲存區..."
    $LogFiles = @($transcriptLogFile, $customLogFile)
    foreach ($file in $LogFiles) {
        $destinationPath = Join-Path $Log_Folder_Path (Split-Path $file -Leaf)
        Copy-Item -Path $file -Destination $destinationPath -Force
        Write-Host "已複製紀錄檔: $(Split-Path $file -Leaf)"
    }

    Write-Host "腳本執行完成。紀錄檔已上傳到網路儲存區。"
    exit
}

Write-CustomLog "發現 $($Java_32_installeds.Count) 個32位Java版本和 $($Java_64_installeds.Count) 個64位Java版本"

# 移除所有Java版本
Write-CustomLog "開始移除32位Java版本..."
Uninstall-Java -JavaInstalls $Java_32_installeds

Write-CustomLog "開始移除64位Java版本..."
Uninstall-Java -JavaInstalls $Java_64_installeds

# 停止記錄
Stop-Transcript

# 紀錄檔管理
$Log_Folder_Path = Join-Path $Log_Path "Java"
if(!(Test-Path -Path $Log_Folder_Path)) { 
    New-Item -ItemType Directory -Path $Log_Folder_Path -Force | Out-Null
}

Write-Host "正在將紀錄檔複製到網路儲存區..."
$LogFiles = @($transcriptLogFile, $customLogFile)
foreach ($file in $LogFiles) {
    $destinationPath = Join-Path $Log_Folder_Path (Split-Path $file -Leaf)
    Copy-Item -Path $file -Destination $destinationPath -Force
    Write-Host "已複製紀錄檔: $(Split-Path $file -Leaf)"
}

Write-Host "Java移除過程已完成。所有紀錄檔已上傳到網路儲存區。"