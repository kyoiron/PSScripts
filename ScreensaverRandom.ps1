# 宣告一個字串陣列，字串長度勿超過全形15字元
$IT_Security_Slogan =  @("密碼換新，程式更新，下載當心", "帳密勿張貼電腦或螢幕週邊可及處", "帳密及自然人憑證不得與他人共用","帳密收妥勿外流，禁讓同學觸電腦","資 訊 安 全 人 人 有 責","電子郵件傳送機敏檔案應加密","機敏檔案應置於加密磁區（D槽）")

# 生成一個隨機索引
$randomIndex = Get-Random -Minimum 0 -Maximum $IT_Security_Slogan.Length

# 使用隨機索引從陣列中取得一個項目
$randomItem = $IT_Security_Slogan[$randomIndex]

#螢幕保護程式ssText3d之登錄檔路徑
$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Screensavers\ssText3d"


#Write-Host "隨機選擇的項目：" $newValue
if (Test-Path -Path $registryPath) {
    # 如果項目存在，則使用 Set-ItemProperty 修改 DisplayString 項目的值
    Set-ItemProperty -Path $registryPath -Name "DisplayString" -Value $IT_Security_Slogan[$randomIndex]
} 