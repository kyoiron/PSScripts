# 設定日誌路徑
$Log_Path = "\\172.29.205.114\Public\sources\audit\SEP_Install_Log"
$ComputerName = $env:COMPUTERNAME
$LogFile = Join-Path $Log_Path ("{0}_SEP_Install_Log.txt" -f $ComputerName)

# 確保日誌目錄存在
if (-not (Test-Path $Log_Path)) {
    New-Item -ItemType Directory -Force -Path $Log_Path | Out-Null
}

# 日誌函數
function Write-Log {
    param([string]$Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Add-Content -Path $LogFile -Value $logMessage
    Write-Host $logMessage
}

# 清理舊日誌內容
Clear-Content -Path $LogFile -ErrorAction SilentlyContinue

Write-Log "開始執行 SEP 安裝腳本"

# 尋找 7-Zip 路徑的函數
function Find-7Zip {
    $possiblePaths = @(
        "${env:ProgramFiles}\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
        (Get-ItemProperty -Path "HKLM:\SOFTWARE\7-Zip" -ErrorAction SilentlyContinue)."Path" + "7z.exe",
        (Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\7-Zip" -ErrorAction SilentlyContinue)."Path" + "7z.exe"
    )

    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    return $null
}

$Sep_Registry = @(
    "HKLM:\software\symantec\symantec endpoint protection\smc",
    "HKLM:\software\wow6432node\symantec\symantec endpoint protection\smc"
)
$Sep_NeedReboot_Registry = 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\RebootMgr'
$SEP_Path = "\\172.29.205.114\loginscript\Update\SEP"
$CacheFilePath = "$SEP_Path\SEP_VersionCache.json"

# 動態選擇最新的SEP安裝檔
$SEP_File = Get-ChildItem -Path "$SEP_Path\*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($SEP_File) {
    # 尋找 7-Zip 路徑
    $7zipPath = Find-7Zip
    if ($7zipPath) {
        $SEP_FileName = $SEP_File.Name
        $SEP_FilePath = $SEP_File.FullName

        # 比較版本號的函數
        function Compare-Versions {
            param (
                [string]$Version1,
                [string]$Version2
            )
            $v1 = [version]$Version1
            $v2 = [version]$Version2
            
            if ($v1 -lt $v2) { return -1 }
            elseif ($v1 -gt $v2) { return 1 }
            else { return 0 }
        }

        # 從 JSON 檔案讀取快取資訊
        function Get-CacheInfo {
            if (Test-Path $CacheFilePath) {
                $cacheContent = Get-Content $CacheFilePath | ConvertFrom-Json
                return $cacheContent
            }
            return $null
        }

        # 更新快取資訊
        function Update-CacheInfo {
            param (
                [string]$Version,
                [DateTime]$LastWriteTime,
                [string]$FileName
            )
            $cacheInfo = @{
                Version = $Version
                Timestamp = (Get-Date).ToString("o")
                LastWriteTime = $LastWriteTime.ToString("o")
                FileName = $FileName
            }
            $cacheInfo | ConvertTo-Json | Set-Content $CacheFilePath
        }

        # 從 Setup.exe 獲取產品版本
        function Get-SetupExeVersion {
            param (
                [string]$SetupPath
            )
            $versionInfo = (Get-Item $SetupPath).VersionInfo
            return $versionInfo.ProductVersion
        }

        # 主要執行邏輯
        $cacheInfo = Get-CacheInfo
        $fileLastWriteTime = $SEP_File.LastWriteTime

        if ($null -eq $cacheInfo -or 
            $cacheInfo.FileName -ne $SEP_FileName -or
            [DateTime]::Parse($cacheInfo.LastWriteTime) -ne $fileLastWriteTime) {
            
            Write-Log "檢測到 SEP 檔案有更新或快取不存在，開始解壓縮並更新快取。"
            
            # 創建臨時資料夾在 $env:systemdrive\temp 下
            $tempFolder = Join-Path "$env:systemdrive\temp" "SEP_Temp_$(Get-Random)"
            New-Item -ItemType Directory -Force -Path $tempFolder | Out-Null
            
            try {
                # 使用 7-Zip 解壓縮 Setup.exe
                $setupExePath = Join-Path $tempFolder "Setup.exe"
                & $7zipPath e $SEP_FilePath "-o$tempFolder" "Setup.exe" -r -y | Out-Null

                if (Test-Path $setupExePath) {
                    $newVersion = Get-SetupExeVersion -SetupPath $setupExePath
                    Update-CacheInfo -Version $newVersion -LastWriteTime $fileLastWriteTime -FileName $SEP_FileName
                    Write-Log "更新快取版本為：$newVersion"
                } else {
                    Write-Log "錯誤：在解壓縮的檔案中找不到 Setup.exe"
                }
            } catch {
                Write-Log "錯誤：解壓縮或讀取版本時發生錯誤：$_"
            } finally {
                # 清理臨時檔案
                Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
            }
        } else {
            Write-Log "SEP 檔案未變更，使用快取的版本資訊。"
        }

        # 檢查並比較版本
        $cacheInfo = Get-CacheInfo  # 重新讀取可能更新的快取資訊
        if ($cacheInfo) {
            $MinimumVersion = $cacheInfo.Version
            $CurrentVersion = $null

            foreach ($regPath in $Sep_Registry) {
                $version = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).ProductVersion
                if ($version) {
                    $CurrentVersion = $version
                    break
                }
            }
            
            if ($null -eq $CurrentVersion) {
                Write-Log "錯誤：無法從註冊表讀取當前 SEP 版本。"
            } elseif ((Compare-Versions -Version1 $CurrentVersion -Version2 $MinimumVersion) -lt 0 -and 
                      !(Test-Path -Path $Sep_NeedReboot_Registry)) {
                Write-Log "當前版本 $CurrentVersion 低於最小要求版本 $MinimumVersion，開始執行更新程序。"
                
                # 創建本地臨時目錄在 $env:systemdrive\temp 下
                $localTempDir = Join-Path "$env:systemdrive\temp" "SEP_Install_$(Get-Random)"
                New-Item -ItemType Directory -Force -Path $localTempDir | Out-Null

                try {
                    # 複製安裝檔到本地臨時目錄
                    $localFilePath = Join-Path $localTempDir $SEP_FileName
                    Write-Log "正在複製 SEP 安裝檔到本地臨時目錄..."
                    Copy-Item -Path $SEP_FilePath -Destination $localFilePath -Force

                    # 執行本地複製的 SEP 安裝檔
                    Write-Log "開始執行 SEP 安裝檔：$localFilePath"
                    # 注意：這裡假設 SEP 安裝檔支持靜默安裝。您可能需要添加適當的命令行參數。
                    Unblock-File $localFilePath
                    $process = Start-Process -FilePath $localFilePath -Wait -PassThru -ErrorAction Stop
                    if ($process.ExitCode -eq 0) {
                        Write-Log "SEP 更新成功完成。"
                    } else {
                        Write-Log "錯誤：SEP 更新過程返回錯誤代碼：$($process.ExitCode)"
                    }
                } catch {
                    Write-Log "錯誤：執行 SEP 更新時發生錯誤：$_"
                } finally {
                    # 清理本地臨時檔案
                    Remove-Item -Path $localTempDir -Recurse -Force -ErrorAction SilentlyContinue
                }
            } else {
                Write-Log "當前版本 $CurrentVersion 已符合或高於最小要求版本 $MinimumVersion，或系統需要重新啟動。無需更新。"
            }
        } else {
            Write-Log "錯誤：無法讀取版本資訊。"
        }
    } else {
        Write-Log "錯誤：無法找到 7-Zip。請確保 7-Zip 已安裝。"
    }
} else {
    Write-Log "錯誤：在 $SEP_Path 中找不到任何 .exe 檔案"
}

Write-Log "SEP 安裝腳本執行完畢"