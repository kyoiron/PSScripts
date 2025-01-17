# 獲取當前用戶
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$userAccount = $currentUser.Name.Split('\')[1]

# 指定允許創建捷徑的帳號列表
$allowedAccounts = @("tnd6083","g4091","tnd5065","tnd6016","tnd5072","tnd5067","tnd6088","yaofu","tnd5040","tnd5081","tnd6084","tnd5037","tnd5131","tnd5020","kyoiron","tnd5057","tnd6034","tnd6073")  # 請將這裡的用戶名替換為實際允許的帳號

# 檢查當前用戶是否在允許的帳號列表中
$isAllowed = $allowedAccounts -contains $userAccount

if ($isAllowed) {
    # 設定捷徑路徑
    $desktopPath = [System.Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path -Path $desktopPath -ChildPath "移機暫存資料夾.lnk"
    $targetPath = "\\172.29.205.114\data\$userAccount"

    # 創建捷徑
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = $targetPath
    $Shortcut.Save()

    Write-Host "捷徑「移機暫存資料夾」已成功創建在桌面上。"
} else {
    Write-Host "當前用戶不在允許創建捷徑的帳號列表中，不創建捷徑。"
}