# 定義函數來獲取文件詳細信息
function Get-FileDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = Split-Path $FilePath
        $file = Split-Path $FilePath -Leaf
        $shellfolder = $shell.Namespace($folder)
        $shellfile = $shellfolder.ParseName($file)
        
        $fieldIndices = @{}
        for ($i = 0; $i -le 266; $i++) {
            $fieldName = $shellfolder.GetDetailsOf($null, $i)
            if ($fieldName -eq "檔案描述" -or $fieldName -eq "檔案版本") {
                $fieldIndices[$fieldName] = $i
            }
            if ($fieldIndices.Count -eq 2) { break }
        }
        
        return @{
            FileDescription = $shellfolder.GetDetailsOf($shellfile, $fieldIndices["檔案描述"])
            FileVersion = $shellfolder.GetDetailsOf($shellfile, $fieldIndices["檔案版本"])
        }
    }
    catch {
        Write-Error "讀取檔案資訊時發生錯誤: $_"
        return $null
    }
    finally {
        if ($shell) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        }
    }
}

# 定義日誌函數
function Write-Log {
    param (
        [string]$Message
    )
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Add-Content -Path $logFile -Value $logMessage
    Write-Host $logMessage
}

# 配置變量
$CDCPKI_Path = "\\172.29.205.114\loginscript\Update\CDCPKI"
$Install_IF_NOT_Installed = $true
$tempPath = "C:\temp"
$logFile = Join-Path $tempPath "CDCPKI_Install_Log.txt"
$auditPath = "\\172.29.205.114\Public\sources\audit\CDCServiSignAdapterSetup"

# 創建日誌目錄
if (-not (Test-Path $tempPath)) {
    New-Item -Path $tempPath -ItemType Directory | Out-Null
}

Write-Log "開始 CDCPKI 安裝程序"

# 獲取最新的 CDCPKI 安裝文件
$CDCPKI_EXE = Get-ChildItem -Path "$CDCPKI_Path\*.exe" | 
              Sort-Object -Property VersionInfo.FileVersionRaw -Descending | 
              Select-Object -First 1

if (-not $CDCPKI_EXE) {
    Write-Log "錯誤: 找不到 CDCPKI 安裝文件"
    exit
}

$CDCPKI_EXE_Path = $CDCPKI_EXE.FullName
$fileDetails = Get-FileDetails -FilePath $CDCPKI_EXE_Path

if (-not $fileDetails) {
    Write-Log "錯誤: 無法獲取文件詳細信息"
    exit
}

$CDCPKI_EXE_FileDescription = $fileDetails.FileDescription.TrimEnd(".exe")
$CDCPKI_EXE_FileVersion = [version]$fileDetails.FileVersion

Write-Log "找到 CDCPKI 安裝文件: $CDCPKI_EXE_Path"
Write-Log "版本: $CDCPKI_EXE_FileVersion"

# 檢查是否已安裝
$RegUninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

$installed = Get-ItemProperty $RegUninstallPaths | 
             Where-Object { $_.DisplayName -match $CDCPKI_EXE_FileDescription } |
             Sort-Object -Property DisplayVersion -Descending |
             Select-Object -First 1

if ($installed) {
    Write-Log "發現已安裝的版本: $($installed.DisplayVersion)"
    if ([version]$CDCPKI_EXE_FileVersion -le [version]$installed.DisplayVersion) {
        Write-Log "已安裝的版本是最新的，無需更新"
        exit
    }
    else {
        Write-Log "開始移除舊版本"
        $uninstallString = $installed.UninstallString
        if ($uninstallString) {
            #Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString /S" -Wait
            & $uninstallString /S
            Write-Log "舊版本移除完成"
        }
        else {
            Write-Log "警告: 無法找到解除安裝字串，跳過移除步驟"
        }
    }
}
elseif (-not $Install_IF_NOT_Installed) {
    Write-Log "CDCPKI 未安裝，且未設置自動安裝。"
    exit
}

# 準備安裝
$EXE_FileName = $CDCPKI_EXE.Name

# 複製安裝文件到臨時目錄
Copy-Item -Path $CDCPKI_EXE_Path -Destination $tempPath -Force
Write-Log "安裝文件已複製到臨時目錄"

# 解除文件阻止並安裝
$tempFilePath = Join-Path $tempPath $EXE_FileName
Unblock-File $tempFilePath
Write-Log "開始安裝 CDCPKI"
& ($tempFilePath) /S

# 驗證安裝
$newInstalled = Get-ItemProperty $RegUninstallPaths | 
                Where-Object { $_.DisplayName -match $CDCPKI_EXE_FileDescription } |
                Sort-Object -Property DisplayVersion -Descending |
                Select-Object -First 1

if ($newInstalled -and [version]$newInstalled.DisplayVersion -eq $CDCPKI_EXE_FileVersion) {
    Write-Log "CDCPKI 安裝成功，版本: $($newInstalled.DisplayVersion)"
}
else {
    Write-Log "錯誤: CDCPKI 安裝失敗或版本不匹配"
}

# 清理臨時文件
Remove-Item $tempFilePath -Force
Write-Log "臨時安裝文件已刪除"

# 複製日誌到審計目錄
$auditLogName = "$env:COMPUTERNAME`_CDCPKI_$CDCPKI_EXE_FileVersion.txt"
$auditLogPath = Join-Path $auditPath $auditLogName
Copy-Item -Path $logFile -Destination $auditLogPath -Force
Write-Log "日誌檔案已複製到審計目錄: $auditLogPath"

Write-Log "CDCPKI 安裝程序完成"