# 函數：獲取 Symantec Endpoint Protection 路徑
function Get-SymantecPath {
    $paths = @("${env:ProgramFiles(x86)}\Symantec\Symantec Endpoint Protection\Smc.exe",
               "${env:ProgramFiles}\Symantec\Symantec Endpoint Protection\Smc.exe")
    foreach ($path in $paths) {
        if (Test-Path $path) { return $path }
    }
    Write-Warning "找不到 Symantec Endpoint Protection。將跳過 Symantec 相關操作。"
    return $null
}
$SymantecPath = Get-SymantecPath

# 函數：解碼 Symantec 密碼
function Get-DecodedSymantecPassword {
    $encodedPassword = "c3ltYW50ZWM="  # 這是 密碼 的 Base64 編碼
    $decodedBytes = [System.Convert]::FromBase64String($encodedPassword)
    return [System.Text.Encoding]::UTF8.GetString($decodedBytes)
}

#強制對時
    #RODC(機關時間伺服器)之IP，如機關未自架可指向法務部172.31.1.2
    $NTPServer="172.29.204.63"
    w32tm /config /manualpeerlist:"$NTPServer" /syncfromflags:manual /update | Out-Null
    
#將tnduser加入本機管理者帳號
Add-LocalGroupMember -Group "Administrators" -Member "$env:COMPUTERNAME\tnduser" -ErrorAction SilentlyContinue
#將tnduser加入本機管理者帳號
net user "tndadmin" /active:yes
Get-LocalUser -Name tnduser | Select-Object * | Out-File $env:SystemDrive\temp\${env:computername}_tnduserStatus.txt
if(test-path("$env:SystemDrive\temp\${env:computername}_tnduserStatus.txt")){Copy-Item "$env:SystemDrive\temp\${env:computername}_tnduserStatus.txt" -Destination  "\\172.29.205.114\Public\sources\audit\tnduser" -Force}
#解鎖Bitlocker槽
$BitlockerRecoveryKey='102157-408870-455730-463155-149358-296956-700711-045573'
#$DontUnlockPrinterPC = @("TND-ICPSC-051","TND-ICPSC-052","TND-ICPSC-053")
if($env:BitlockerDataDrive -ne $null){
    manage-bde -unlock $env:BitlockerDataDrive -RecoveryPassword $BitlockerRecoveryKey
    #manage-bde -off $env:BitlockerDataDrive
}

<#
#清空temp資料夾
Remove-Item -Path $env:systemdrive\temp\* -Force -Exclude RunAsAdministratorPS.ps1 -ErrorAction SilentlyContinue
#>
#檢查UAC有沒有開啟，沒開啟則開啟。
if((Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA -eq 0){
    Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name EnableLUA -Value 1
}

#印表機更名：將中括號[,]替換成【】，因為中括號容易造成字串字元判斷困難
    Get-printer | Where-Object{$_.Name -like  ("*"+[regex]::escape(']'))} |Where-Object{ Rename-Printer -name $_.Name -NewName ((($_.Name -replace  [regex]::escape('['),'【') -replace  [regex]::escape(']'),'】'))}
#印表機備份程式
    powershell  "$env:SystemDrive\temp\PrinterBackup.ps1"
#印表機權限
    $Temp_PermissionSDDL="G:SYD:(A;;LCSWSDRCWDWO;;;WD)(A;OIIO;RPWPSDRCWDWO;;;WD)(A;;SWRC;;;AC)(A;CIIO;RC;;;AC)(A;OIIO;RPWPSDRCWDWO;;;AC)(A;;LCSWSDRCWDWO;;;CO)(A;OIIO;RPWPSDRCWDWO;;;CO)(A;OIIO;RPWPSDRCWDWO;;;BA)(A;;LCSWSDRCWDWO;;;BA)"
    Get-printer | ForEach-Object{ Set-Printer $_.Name -PermissionSDDL $Temp_PermissionSDDL}


#檢查跨平台工具有無執行，有安裝而無執行的話，就執行
    $ChkSrv = Get-Process -Name ChkSrv -ErrorAction SilentlyContinue
    if((!$ChkSrv) -and (Test-Path -Path ${env:ProgramFiles(x86)}\HiPKILocalSignServer\ChkSrv.exe)){Start-Process -FilePath "${env:ProgramFiles(x86)}\HiPKILocalSignServer\ChkSrv.exe"}

#HiCOS更新
    #不更新Hicos的電腦名稱

    #powershell  "$env:SystemDrive\temp\HiCOS_Remove.ps1"    
    <#
    $NotUpdateHicos_ComputerName=@()    
    if(!$NotUpdateHicos_ComputerName.Contains($env:computername)){
        powershell  "$env:SystemDrive\temp\HiCOS_Update.ps1"        
    }
    #>
#if($env:computername -eq "TND-ASSE-025"){powershell  "$env:SystemDrive\temp\HiCOS_UpdateV2.ps1" 0 }
#if($env:computername -eq "TND-ASSE-025"){powershell  "$env:SystemDrive\temp\HiCOS_UpdateV2.ps1" "1"}
#if($env:computername -eq "TND-STOF-112"){powershell  "$env:SystemDrive\temp\HiCOS_UpdateV2.ps1" "1"}

<#$Windows_Update_Repair_PC = @("TND-COU-048","TND-CENTRAL-144")
if((Get-WUInstallerStatus -Silent) -or $Windows_Update_Repair_PC.Contains($env:computername)){
    Reset-WUComponents -Verbose
}#>

#Chrome更新
    powershell "$env:SystemDrive\temp\ChromeUpdate.ps1"
#Adobe Reader更新
    powershell "$env:SystemDrive\temp\AdobeReaderUpdateV2.ps1"
#Java更新
    #powershell "$env:SystemDrive\temp\JavaUpdate.ps1"
#7zip更新
    powershell "$env:SystemDrive\temp\7zipUpdateV3.ps1"
#經費結報系統元件GeasBatchsign更新
    #$GeasBatchsign_repair_PC=@('TND-ACOF-040')
    #$GeasBatchsign_repair_PC=@('TND-STOF-112')
    $GeasBatchsign_repair_PC=@()
    if($GeasBatchsign_repair_PC.Contains($env:computername)){
        $GeasBatchsigns_Path = "\\172.29.205.114\loginscript\Update\GeasBatchsign"
        $Log_Path = "\\172.29.205.114\Public\sources\audit"
        $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
        $GeasBatchsign2_installeds = $null
        $GeasBatchsign_installeds = $null
        foreach ($Path in $RegUninstallPaths) {
            if (Test-Path $Path) {
                $GeasBatchsign2_installeds += @(Get-ItemProperty $Path | Where-Object{$_.DisplayName -like "背景簽章服務2"})
                $GeasBatchsign_installeds += @(Get-ItemProperty $Path | Where-Object{$_.DisplayName -eq "背景簽章服務CS"})
            }
        }
        #移除所有已安裝的2版程式
        if($GeasBatchsign2_installeds){
            Taskkill /f /im javaw.exe
            foreach($exe in $GeasBatchsign2_installeds){        
                $uninstall = ($exe.UninstallString  -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()
                $LogName = $env:Computername + "_"+ $exe.DisplayName+"_"+ $exe.DisplayVersion + ".txt"
                $LogFile = $env:systemdrive+"\temp\" + $LogName            
                start-process "msiexec.exe" -arg "/X $uninstall /quiet /passive /norestart /log ""$LogFile""" -Wait -WindowStyle Hidden               
                $Log_Folder_Path = $Log_Path + "\"+ $exe.DisplayName
                if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
                if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null} 
        
            }            
        }
        if($GeasBatchsign_installeds){
            Taskkill /f /im javaw.exe
            foreach($exe in $GeasBatchsign_installeds){        
                $uninstall = ($exe.UninstallString  -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()
                $LogName = $env:Computername + "_"+ $exe.DisplayName+"_"+ $exe.DisplayVersion + "_Remove.txt"
                $LogFile = $env:systemdrive+"\temp\" + $LogName            
                start-process "msiexec.exe" -arg "/X $uninstall /quiet /passive /norestart /log ""$LogFile""" -Wait -WindowStyle Hidden                
                $Log_Folder_Path = $Log_Path + "\"+ $exe.DisplayName
                if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
                if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null} 
        
            }
        }
        $userFolders = Get-ChildItem -Path "$env:systemdrive\Users" -Directory
        foreach ($userFolder in $userFolders) {
            $pathToSearch = Join-Path -Path $userFolder.FullName -ChildPath "AppData\Local\BatchSignCS"
            $pathToSearch2 = Join-Path -Path $userFolder.FullName -ChildPath "AppData\Local\Temp"
            if (Test-Path $pathToSearch -PathType Container){
                Remove-Item -Path $pathToSearch -Recurse -Force -ErrorAction SilentlyContinue        
            }            
        }                
    }
    powershell "$env:SystemDrive\temp\GeasBatchsign.ps1"
#MariaDB ODBC Driver 64-bit更新
    powershell "$env:SystemDrive\temp\MariadbConnectorOdbc.ps1"
#ThreatSonarPC檢測
    powershell "$env:SystemDrive\temp\ThreatSonarPCv3.ps1"
#安裝VANs軟體
    powershell "$env:SystemDrive\temp\WM7AssetCluster.ps1"
#安裝新版軟體FileZilla
    powershell "$env:SystemDrive\temp\FileZillaUpdate.ps1"
#安裝新版軟體XnView
    powershell "$env:SystemDrive\temp\XnViewUpdate.ps1"
#安裝新版軟體K-LiteMegaCodecPack
    powershell "$env:SystemDrive\temp\K-LiteMegaCodecPackUpdate.ps1"
#安裝新版軟體PDF-Xchange Editor 
    powershell "$env:SystemDrive\temp\PDFXChangeEditorUpdate.ps1" 
#安裝新版中推會用戶端更新工具軟體
    powershell "$env:SystemDrive\temp\CMEXFontClientUpdate.ps1"
#刪除不要的軟體
    #powershell "$env:SystemDrive\temp\UninstallSoftware.ps1"
#PC基本資料蒐集
 #   powershell "$env:SystemDrive\temp\PCChecker.ps1"
#檢查嘸蝦米輸入法安裝狀況
   powershell "$env:SystemDrive\temp\CheckBoshiamyTIP.ps1" 
#檢查自然輸入法安裝狀況
    powershell "$env:SystemDrive\temp\CheckGoing.ps1"
#檢查電腦板LINE安裝狀況
    powershell "$env:SystemDrive\temp\CheckLINE.ps1"


#檢查JDK安裝狀況
    if($env:computername -eq "TND-PEOF-015"){ 
        if(Test-Path -Path "$env:SystemDrive\temp\JDKUpdate.ps1"){     
            powershell "$env:SystemDrive\temp\JDKUpdate.ps1"
        }
    }

#更新NewClient版本
<#
if($evn:computername -eq "TND-GASE-061"){
    $SourcePath = "D:\NewClient"
    $TargetPath = "\\172.29.205.114\Public\sources\audit\NewClient\backup"
    $IsEmpty = (Get-ChildItem -Path $TargetPath).Count -eq 0
    if($IsEmpty){
        # 備份 NEWCLIENT_lib 資料夾
            Copy-Item -Path "$SourcePath\NEWCLIENT_lib" -Destination $TargetPath -Recurse -ErrorAction SilentlyContinue  
        # 備份 NEWCLIENT.jar 檔案
            Copy-Item -Path "$SourcePath\NEWCLIENT.jar" -Destination $TargetPath -ErrorAction SilentlyContinue
        # 將 NewClient340 資料夾內的檔案覆蓋到 D:\NewClient 資料夾中
            Copy-Item -Path "\\172.29.205.114\Public\sources\audit\NewClient\NewClient340" -Destination $SourcePath -Recurse -Force -ErrorAction SilentlyContinue
    }
}
#>

#指定電腦（們）匯入特定電腦的印表機封裝檔
#要匯入的電腦
#$InstallPrinterPC=@("TND-SASE-016","TND-SASE-089","TND-SASE-091","TND-SASE-095","TND-SASE-102","TND-SASE-107")
#$InstallPrinterPC=@("TND-STOF-136")
<#
$InstallPrinterPC=@("TND-CPSC-064")
#指定哪個電腦的匯出當匯入範本
$ImportFromComputername = "TND-CPSC-064"
#>
#$InstallPrinterPC=@("TND-CPSC-064")
#指定哪個電腦的匯出當匯入範本
#$ImportFromComputername = "TND-STOF-113"
#$pattern = '^TND-CRS-\d{3}$'
#$InstallPrinterPC = @()
#$InstallPrinterPC.Contains($env:computername)
if($env:computername -match $pattern){
    $PrinterExportFileLocation = "\\172.29.205.114\mig\Printer"
    $File_Name = $ImportFromComputername +"x64.printerExport"
    $File_FullName = $PrinterExportFileLocation + "\" + $File_Name
    if(Test-Path $File_FullName){
        robocopy $PrinterExportFileLocation "$env:systemdrive\temp" $File_Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        $tempFile = "$env:systemdrive\temp\" + $File_Name
        start-Process cmd.exe -Verb RunAs -Args '/c',"${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe -R -F $tempFile -O FORCE" -Wait
        Get-printer | ForEach-Object{Set-Printer $_.Name -PermissionSDDL ((Get-Printer -Name 'Microsoft Print to PDF' -full).PermissionSDDL)}
        Remove-Item $tempFile -Force
        remove-item -Path ($PrinterExportFileLocation+"\"+$env:computername+"x64.printerExport") -Force
        #powershell "$env:SystemDrive\temp\SpecificPrintersImport.ps1"
    }
}


#$Rebuid_EICSignTSR_PC=@("TND-BUSE-072")
$Rebuid_EICSignTSR_PC=@()
if($Rebuid_EICSignTSR_PC.Contains($env:computername)){   
    get-item -Path "$env:PUBLIC\EICSignTSR\EicSignTSR.ini"
    Stop-Process -Name EicSignTSR -Force -ErrorAction Continue
    Remove-Item -Path "$env:PUBLIC\EICSignTSR\EicSignTSR.ini" -Force
    Start-Process "$env:SystemDrive\eic\EICSignTSR\EicSignTSR.exe"
    Start-Sleep -s 3 
    Stop-Process -Name EicSignTSR -Force
}

#指紋機
$FingerPrint_PC =@("TND-STOF-113")
if($FingerPrint_PC.Contains($env:computername)){
    
    $pfxFilePath = "\\172.29.205.114\loginscript\Update\PIDDLL\NeoFaceCert.pfx"
    $thumbprint = "9E8F00E882B44F9C955B555B2D7FA4EB3FBA94F3"
    $password = ConvertTo-SecureString "1qaz@WSX" -AsPlainText -Force
    # "個人"證書存儲區
        $certPathPersonal = "Cert:\LocalMachine\My"
        $IsInstalledCertificatePersonal = (Get-ChildItem -Path $certPathPersonal | Where-Object { $_.Thumbprint -eq $thumbprint }) -ne $null
    #"受信任的根憑證授權單位"證書存儲區
        $certPathRoot = "Cert:\LocalMachine\Root"
        $IsInstalledCertificateRoot = ((Get-ChildItem -Path $certPathRoot | Where-Object {$_.Thumbprint -eq $thumbprint})) -ne $null
    # 匯入 PFX 憑證到 "個人"證書存儲區
    if(!$IsInstalledCertificatePersonal){    
        $certificatePersonal = Import-PfxCertificate -FilePath $pfxFilePath -Password $password -CertStoreLocation $certPathPersonal -Exportable
    }
    # 匯入 PFX 憑證到 "受信任的根憑證授權單位" 證書存儲區
    if(!$IsInstalledCertificatePersonal ){    
        $certificateRoot = Import-PfxCertificate -FilePath $pfxFilePath -Password $password -CertStoreLocation $certPathRoot -Exportable
    }    
    powershell "$env:SystemDrive\temp\IB_Driver.ps1"
    $NAS_CPID_Client = "\\172.29.205.114\loginscript\Update\PIDDLL\CPID_Client"
    $NAS_PIDDLLV2 = "\\172.29.205.114\loginscript\Update\PIDDLL\PIDDLLV2"
    $PC_CPID_Client = "D:\CPID_Client"
    $PC_PIDDLLV2 = "D:\PIDDLLV2"
    if(!(Test-Path $PC_CPID_Client -PathType Container)){robocopy $NAS_CPID_Client $PC_CPID_Client "/E".Split(' ') | Out-Null}
    if(!(Test-Path $PC_PIDDLLV2 -PathType Container)){robocopy $NAS_PIDDLLV2 $PC_PIDDLLV2 "/E".Split(' ') | Out-Null}

    if((!(Test-Path "$env:PUBLIC\Desktop\指紋系統.lnk")) -and (Test-Path -Path "D:\PIDDLLV2" -PathType Container)-and(Test-Path -Path "D:\CPID_Client" -PathType Container) -and (Test-Path -Path "D:\PIDDLLV2\PIDDLL.exe")){
        $targetPath = "D:\PIDDLLV2\PIDDLL.exe"
        $shortcutPath = "$env:PUBLIC\Desktop\指紋系統.lnk"
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetPath
        $shortcut.Save()    
    }

}

<#
#指紋機驅動程式安裝
$FingerPrint_PC =@("TND-COU-141","TND-RMSE-142"."TND-GCSE-143","TND-CENTRAL-144","TND-GCSE-146","TND-GCSE-147","TND-STOF-112","TND-STOF-113")
#$FingerPrint_PC =@("TND-STOF-113")
#$FingerPrint_PC =@()
if($FingerPrint_PC.Contains($env:computername)){   
    $RegUninstallPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $Driver_RET = "\\172.29.205.114\loginscript\Update\Driver_RET"
    $Log_Path = "\\172.29.205.114\Public\sources\audit"    
    $DigitalPersona_installeds = Get-ItemProperty $RegUninstallPath | Where-Object{$_.DisplayName -match "DigitalPersona SDK Runtime"}               
    $Log_Folder_Path = $Log_Path +"\"+ "DigitalPersona SDK Runtime"
    if(($DigitalPersona_installeds -eq $null) -or ([Version]$DigitalPersona_installeds.DisplayVersion -lt  [Version]3.4.0.127)){
        robocopy $Driver_RET "$env:systemdrive\temp"  /E /XO /NJH /NJS /NDL /NC /NS | Out-Null

        #start-process $env:systemdrive\temp\RET\InstallOnly.bat
        $arguments="/s /v""REBOOT=ReallySuppress /qn /l*v $env:systemdrive\temp\$env:Computername" + "_ururte_install.log"""        
        start-process ($env:systemdrive+"\temp\RET\Setup.exe") -arg $arguments -wait -WindowStyle Hidden                
    }
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    $LogPattern="${env:Computername}"+"_ururte_install.log"
    if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
}
#>
<#
$FingerPrint_PC=@()
if($FingerPrint_PC.Contains($env:computername)){
    if((Get-WindowsOptionalFeature -Online -FeatureName TelnetClient).State -eq "Disabled"){
        Enable-WindowsOptionalFeature -Online -FeatureName TelnetClient
    }
}
#>

#同電腦名稱印表機匯入作業
<#
$ImportPrinterPC=@("TND-PEOF-030")
if ($ImportPrinterPC.Contains($env:computername)){    
    $PrinterExportFileLocation = "\\172.29.205.114\mig\Printer"
    $File_Name = $env:computername+"x64.printerExport"
    $File_FullName = $PrinterExportFileLocation + "\" + $File_Name
    robocopy $PrinterExportFileLocation "$env:systemdrive\temp" $File_Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    $tempFile = "$env:systemdrive\temp\" + $File_Name
    start-Process cmd.exe -Verb RunAs -Args '/c',"${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe -R -F $tempFile" -Wait
    Get-printer | ForEach-Object{ Set-Printer $_.Name -PermissionSDDL ((Get-Printer -Name 'Microsoft Print to PDF' -full).PermissionSDDL)}
    Remove-Item $tempFile
    Move-Item -Path $File_FullName -Destination "\\172.29.205.114\mig\Printer_BACKUP" -Force
}
#>


#更新SEP版本
<#
$Sep_Registry = "HKLM:\software\wow6432node\symantec\symantec endpoint protection\smc"
$Sep_NeedReboot_Registry = 'HKLM:\\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\RebootMgr'
if(((Get-ItemProperty -Path $Sep_Registry).ProductVersion -ne '14.3.558.0000') -and (!(Test-Path -Path  $Sep_NeedReboot_Registry))){
    powershell "$env:SystemDrive\temp\SEP_AutoUpdate.ps1"
}
#>

if (Test-Path "$env:SystemDrive\temp\SEP_AutoUpdate.ps1") {
    & "$env:SystemDrive\temp\SEP_AutoUpdate.ps1"
}

#確認SEP啟動並進行防毒更新病毒碼及政策
    if(!(Get-Process -Name ccSvcHst -ErrorAction SilentlyContinue)){
        $symantecPassword = Get-DecodedSymantecPassword
        Start-Process -FilePath $SymantecPath -ArgumentList "-p `"$symantecPassword`" -start" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
        $symantecPassword = $null  # 清除密碼變數
    }

    #Start-Process -FilePath "${env:ProgramFiles(x86)}\Symantec\Symantec Endpoint Protection\Smc.exe" -ArgumentList ' -updateconfig' -WindowStyle Hidden
    #Start-Process -FilePath "${env:ProgramFiles(x86)}\Symantec\Symantec Endpoint Protection\SepLiveUpdate.exe" -WindowStyle Hidden
#進行WindowsUpdate
    #如果沒有PSWindowsUpdate模組則安裝
    if((Get-Module -ListAvailable -Name PSWindowsUpdate) -eq $null){
        $PSWindowsUpdate_Path = "\\172.29.205.114\loginscript\PSWindowsUpdate"
        $PSModule_Path = "$Env:ProgramFiles\WindowsPowerShell\Modules\PSWindowsUpdate"    
        robocopy $PSWindowsUpdate_Path $PSModule_Path "/e /XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        Import-Module PSWindowsUpdate             
    }

    $LogPath = "\\172.29.205.114\Public\sources\audit\WSUS"
    $temp = "$env:SystemDrive\temp"
    $ServiceID_WindowsUpdate = (Get-WUServiceManager | Where-Object{$_.Name -like "Windows Update"}).ServiceID
    $ServiceID_WSUS = (Get-WUServiceManager | Where-Object{$_.Name -like "Windows Server Update Service"}).ServiceID
    #$PSWUSettings = @{SmtpServer="smtp.moj.gov.tw";From="tndi@mail.moj.gov.tw";To="kyoiron@mail.moj.gov.tw";Port=25}
    <# 暫停 觀察
    start-job -ScriptBlock {
        Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WSUS -ScheduleJob (Get-Date).AddMinutes(30)  -Verbose
        #Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot  -ServiceID $ServiceID_WSUS -ScheduleJob | Out-File "$env:SystemDrive\temp\${env:computername}_WindowsUpdate.txt" -Force -Append  
    }
    #>
    Get-WUHistory | Format-Table -AutoSize | Out-File "$temp\${env:computername}_WindowsUpdate_History.txt" -Force
    Robocopy $temp $LogPath "${env:computername}_WindowsUpdate.txt" "${env:computername}_WindowsUpdate_History.txt" " /XO /NJH /NJS /NDL /NC /NS".Split(' ')| Out-Null

#Java更新
    #powershell "$env:SystemDrive\temp\JavaUpdate.ps1"
#移除Java
    powershell "$env:SystemDrive\temp\UninstallJAVA.ps1"
