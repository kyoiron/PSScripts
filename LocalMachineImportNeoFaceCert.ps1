# 指定 PFX 憑證檔案的路徑
$pfxFilePath = "\\172.29.205.114\loginscript\Update\PIDDLL\NeoFaceCert.pfx"

# 提示用戶輸入 PFX 憑證的密碼
$password = ConvertTo-SecureString "1qaz@WSX" -AsPlainText -Force

# 匯入 PFX 憑證到 "個人" 證書存儲區
$certPathPersonal = "Cert:\LocalMachine\My"
$certificatePersonal = Import-PfxCertificate -FilePath $pfxFilePath -Password $password -CertStoreLocation $certPathPersonal -Exportable

# 匯入 PFX 憑證到 "受信任的根憑證授權單位" 證書存儲區
$certPathRoot = "Cert:\LocalMachine\Root"
$certificateRoot = Import-PfxCertificate -FilePath $pfxFilePath -Password $password -CertStoreLocation $certPathRoot -Exportable

# 顯示匯入的憑證資訊
#$certificatePersonal
#$certificateRoot



