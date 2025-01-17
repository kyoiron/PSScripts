param (
    [ValidateSet("0", "1")]
    [string] $mode = 0
    #模式 0：檢查有無異常，有需要在裝；
    #模式 1：強製重新移除在安裝
)
if ($mode -ne "0" -and $mode -ne "1"){$mode = "0"}

#放置Hicos EXE的位置
    $HiCOSs_Path = "\\172.29.205.114\loginscript\Update\HiCOS"
#放置安裝LOG檔的網路位置
    $Log_Path = "\\172.29.205.114\Public\sources\audit"
#取得EXE檔版本
    $HiCOS_EXE = (Get-ChildItem -Path ($HiCOSs_Path+"\*.exe") | Where-Object{$_.VersionInfo.ProductName -eq "HiCOS PKI Smart Card"} | Sort-Object)
#取得EXE檔完整路徑
    $global:HiCOS_EXE_Path = $HiCOS_EXE.FullName
#預設Hicos程式安裝位置
    $global:HiCOS_Folder = "${env:ProgramFiles(x86)}\Chunghwa Telecom\HiCOS PKI Smart Card"
    $global:HiCOS_StartMenuFolder = "${env:ProgramData}\Microsoft\Windows\Start Menu\Programs\HiCOS PKI Smart Card"
    $HiCOS_Displayname = "HiCOS PKI Smart Card"
#預設跨平台元件(HiPKILocalSignServer)安裝位置
    $global:HiPKILocalSignServer_Folder = "${env:ProgramFiles(x86)}\HiPKILocalSignServer"
    $global:HiPKILocalSignServer_StartMenuFolder = "${env:ProgramData}\Microsoft\Windows\Start Menu\Programs\跨平台網頁元件"
    $HiPKILocalSignServer_Displayname = "跨平台網頁元件*"


#EXE檔跨平台元件(HiPKILocalSignServer)版本正規表示式
    $HiPKILocalSignServer_Version_Pattern = '_________\d+\.\d+\.\d+\.\d+\.exe'
#EXE檔跨平台元件(HiPKILocalSignServer)版本
    $HiCOS_EXE_HiPKILocalSignServer_Version = $null
#登入檔路徑
$global:RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')

function Uninstall_Hicos_EXE(){

}

function Uninstall_Hicos($HiCOS_installeds) {
    #關閉HiCOS PKI Smart Card Token Utility程式
        #&"${env:ProgramFiles(x86)}\HiPKILocalSignServer\ChkSrv.exe" "/stop"                    
         &"taskkill.exe" /f /im TokenUtility.exe 

    # 在這裡執行對輸入字串的處理
        
        foreach($item in $HiCOS_installeds){
            $uninstall_Char = ($item.UninstallString -split "  ")
            if($uninstall_Char[0] -eq ""){$uninstall_EXE="$env:systemdrive\temp\HiCOS_Client.exe"}else{$uninstall_EXE = $uninstall_Char[0].Trim()}
            $LogFile= "$env:systemdrive\temp\"+$env:Computername + "_HiCOS_Uninstall_"+ $item.DisplayVersion + ".txt"
            $arguments = " /uninstall /passive /quiet /norestart /log " + $LogFile                
            start-process $uninstall_EXE  -arg $arguments -Wait -WindowStyle Hidden
        }
    #確認還有無存在
        foreach ($Path in $RegUninstallPaths) {
            if (Test-Path $Path) {
                $HiCOS_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $HiCOS_Displayname}
            }
        }
    #仍存在則強制移除
    if($HiCOS_installeds -ne $null){
        foreach($item in $HiCOS_installeds){
            Uninstall_Hicos_Force $item
        }
        
    } 

}

function Uninstall_Hicos_Force($HiCOS_installed) {
    Remove-Item -Path $HiCOS_Folder -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $HiCOS_StartMenuFolder -Recurse -Force -ErrorAction SilentlyContinue    
    Remove-Item -Path "HKEY_CURRENT_USER\Software\Microsoft\HiCOS PKI Smart Card" -Force -Confirm:$false
    Remove-Item -Path "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Chunghwa TeleCom\HiCOS PKI Smart Card" -Force -Confirm:$false 
    Remove-Item -Path ((Convert-Path $HiCOS_installed.PSParentPath)+"HiCOS PKI Smart Card") -Force -Confirm:$false 
    Remove-Item -Path (Convert-Path $HiCOS_installed.PSPath) -Force -Confirm:$false 
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Name "HiCOS Token Utility" -Force -ErrorAction SilentlyContinue    
}

function Uninstall_HiPKILocalSignServer($HiPKILocalSignServer_installeds) {
    #關閉跨平台元件服務
        #&"${env:ProgramFiles(x86)}\HiPKILocalSignServer\ChkSrv.exe" "/stop"                    
         &"taskkill.exe" /f /im ChkSrv.exe 
         &"taskkill.exe" /f /im CheckServer.exe
         &"taskkill.exe" /f /im chtnode.exe
         &"taskkill.exe" /f /im node.exe
    # 在這裡執行對輸入字串的處理
    foreach($item in $HiPKILocalSignServer_installeds){                                                                
        $uninstall_Char_HiPKI = ($item.UninstallString  -split "  ")
        $LogFile_HiPKI = "$env:systemdrive\temp\"+$env:Computername + "_HiPKILocalSignServer_Uninstall_"+ $item.DisplayVersion + ".txt"
        $arguments_HiPKI = " /SILENT /log=" + $LogFile_HiPKI
        if(!(Test-Path -Path $uninstall_Char_HiPKI[0])){
            $uninstall_EXE="$env:systemdrive\temp\HiCOS_Client.exe"
            $LogFile= "$env:systemdrive\temp\"+$env:Computername + "_HiCOS_Uninstall_"+ $item.DisplayVersion + ".txt"
            $arguments = " /uninstall /passive /quiet /norestart /log " + $LogFile                
            start-process $uninstall_EXE  -arg $arguments -Wait -WindowStyle Hidden
        }
        start-process $uninstall_Char_HiPKI[0] -arg $arguments_HiPKI -Wait -WindowStyle Hidden
    }
    #確認還有無存在
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $HiPKILocalSignServer_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $HiPKILocalSignServer_Displayname}
        }
    }
    #仍存在則強制移除
    foreach($item in $HiPKILocalSignServer_installeds){
        if($HiPKILocalSignServer_installeds -ne $null){
            Uninstall_HiPKILocalSignServer_Force $item
        }
    }
}

function Uninstall_HiPKILocalSignServer_Force($HiPKILocalSignServer_installed){
    Remove-Item -Path $HiPKILocalSignServer_Folder -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $HiPKILocalSignServer_StartMenuFolder -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path (Convert-Path $HiPKILocalSignServer_installed.PSPath) -Force -Confirm:$false 
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Name "跨平台網頁元件" -Force -ErrorAction SilentlyContinue
}


#確認新版EXE檔存在的話，則進行程式
    if($HiCOS_EXE_Path){
        #保留一份HiCOS_Client.exe在本機電腦C曹temp資料夾
            robocopy "$HiCOSs_Path" "$env:systemdrive\temp" "HiCOS_Client.exe" "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
            if(Test-Path -Path "$env:systemdrive\temp\HiCOS_Client.exe"){$HiCOS_EXE_Path="$env:systemdrive\temp\HiCOS_Client.exe"}
        #取得EXE檔中的Hicos的程式名稱及版本
            $HiCOS_EXE_ProductName = (Get-ItemProperty $HiCOS_EXE_Path).VersionInfo.ProductName
            $HiCOS_EXE_ProductVersion = (Get-ItemProperty $HiCOS_EXE_Path).VersionInfo.ProductVersion
            <#
                PSPath            : Microsoft.PowerShell.Core\FileSystem::C:\temp\HiCOS_Client.exe
                PSParentPath      : Microsoft.PowerShell.Core\FileSystem::C:\temp
                PSChildName       : HiCOS_Client.exe
                PSDrive           : C
                PSProvider        : Microsoft.PowerShell.Core\FileSystem
                Mode              : -a----
                VersionInfo       : File:             C:\temp\HiCOS_Client.exe
                                    InternalName:     setup
                                    OriginalFilename: HiCOS_Client.exe
                                    FileVersion:      3.0.3.62814
                                    FileDescription:  HiCOS PKI Smart Card
                                    Product:          HiCOS PKI Smart Card
                                    ProductVersion:   3.0.3.62814
                                    Debug:            False
                                    Patched:          False
                                    PreRelease:       False
                                    PrivateBuild:     False
                                    SpecialBuild:     False
                                    Language:         英文 (美國)
                                
                BaseName          : HiCOS_Client
                Target            : {}
                LinkType          : 
                Name              : HiCOS_Client.exe
                Length            : 20344600
                DirectoryName     : C:\temp
                Directory         : C:\temp
                IsReadOnly        : False
                Exists            : True
                FullName          : C:\temp\HiCOS_Client.exe
                Extension         : .exe
                CreationTime      : 2020/8/12 上午 10:06:22
                CreationTimeUtc   : 2020/8/12 上午 02:06:22
                LastAccessTime    : 2020/8/12 上午 10:06:22
                LastAccessTimeUtc : 2020/8/12 上午 02:06:22
                LastWriteTime     : 2020/8/12 上午 09:29:56
                LastWriteTimeUtc  : 2020/8/12 上午 01:29:56
                Attributes        : Archive
            #>
        #取得EXE檔中的跨平台元件(HiPKILocalSignServer)版本
            $Layout_Log_FilePath = "${env:temp}\layoutlog.txt"
            &$HiCOS_EXE_Path  /layout /passive /quiet /norestart /log $Layout_Log_FilePath -wait
            $matchingLine = Get-Content -Path $Layout_Log_FilePath | Select-String -Pattern $HiPKILocalSignServer_Version_Pattern         
            if ($matchingLine) {
                $matchedString = [Regex]::Match($matchingLine.Line, $HiPKILocalSignServer_Version_pattern).Value
                $matchedString = $matchedString -replace '^_________', ''
                $matchedString = $matchedString -replace '\.exe$', ''
                $HiCOS_EXE_HiPKILocalSignServer_Version = $matchedString
            }       
        #找尋電腦中有無安裝Hicos 以及 有無安裝「跨平台網頁元件」
            foreach ($Path in $RegUninstallPaths) {
                if (Test-Path $Path) {
                    $HiCOS_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $HiCOS_EXE_ProductName} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
                    $HiPKILocalSignServer_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $HiPKILocalSignServer_Displayname}
                }
            }
            if(($HiPKILocalSignServer_installeds -ne $null) -and (Test-Path -Path "$HiPKILocalSignServer_Folder\ChkSrv.exe")){
                $HiPKILocalSignServer_Exist=$true
                $HiPKILocalSignServer_installed = $HiPKILocalSignServer_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
            }else{ 
                #有安裝跨平台但是安裝不完整的狀況，則全部移除
                if(($HiPKILocalSignServer_installeds -ne $null) -xor (Test-Path -Path "$HiPKILocalSignServer_Folder\ChkSrv.exe")){                   
                    Uninstall_HiPKILocalSignServer $HiPKILocalSignServer_installeds
                }
                $HiPKILocalSignServer_Exist=$false
            }
            if(($HiCOS_installeds -ne $null) -and (Test-Path -Path "$HiCOS_Folder\TokenUtility.exe")){                
                $HiCOS_Exist=$true
                $HiCOS_installed = $HiCOS_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
            }else{
                 #有安裝hicos但是安裝不完整的狀況，則全部移除
                if(($HiCOS_installeds -ne $null) -xor (Test-Path -Path "$HiCOS_Folder\TokenUtility.exe")){
                   Uninstall_Hicos $HiCOS_installeds
                }
                $HiCOS_Exist=$false
            }
        #重複安裝或者強製安裝(mode為1時)狀況則都移除   
            if((($HiCOS_installeds|Measure-Object).count -ge 2) -or ($mode -eq "1")){
                Uninstall_Hicos $HiCOS_installeds 
                $HiCOS_Exist=$false
            }
            if((($HiPKILocalSignServer_installeds|Measure-Object).count -ge 2) -or ($mode -eq "1")){
                Uninstall_HiPKILocalSignServer $HiPKILocalSignServer_installeds
                $HiPKILocalSignServer_Exist=$false
            }
        
        #版本檢查
            #兩者皆存在的狀況下，才比對版本
            if($HiCOS_Exist -and $HiPKILocalSignServer_Exist){
                if(([version]$HiCOS_installed.BundleVersion -ge [version]$HiCOS_EXE_ProductVersion) -and ([version]$HiPKILocalSignServer_installed.DisplayVersion -ge [version]$HiCOS_EXE_HiPKILocalSignServer_Version) -and $HiPKILocalSignServer_Exist){exit}
            }elseif($HiCOS_Exist){
                Uninstall_Hicos $HiCOS_installeds
            }elseif($HiPKILocalSignServer_Exist){
                Uninstall_HiPKILocalSignServer $HiPKILocalSignServer_installeds
            }
                #安裝EXE  
                $LogName = $env:Computername + "_HiCOS_"+ $HiCOS_EXE_ProductVersion + ".txt"
                #$arguments = " /install  /passive /quiet /norestart /log $env:systemdrive\temp\$LogName"
                $arguments_install = " /install /passive /quiet /norestart /log $env:systemdrive\temp\$LogName"
                #robocopy $HiCOSs_Path "$env:systemdrive\temp" $HiCOS_EXE.Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
                unblock-file ($env:systemdrive+"\temp\"+$HiCOS_EXE.Name)
                #start-process ($env:systemdrive+"\temp\"+$HiCOS_EXE.Name) -arg $arguments_install -WindowStyle Hidden
                $job  = Start-Job -ScriptBlock {
                    param ([string]$exe,[string]$arg1)
                    & ($env:systemdrive+"\temp\"+$exe)  /install /passive /quiet /norestart /log $arg1
                } -ArgumentList $HiCOS_EXE.Name,"$env:systemdrive\temp\$LogName"
                Wait-Job -id $job.Id -Timeout 60
                Remove-Job -id $job.Id

                #Start-Sleep -s 5
                $Log_Folder_Path = $Log_Path +"\"+ $HiCOS_EXE_ProductName
                $LogPattern = $env:Computername + "_HiCOS_*.txt"
                if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
                if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}                                         
    }    



<#
    foreach ($Path in $RegUninstallPaths) {
                if (Test-Path $Path) {
                    $HiCOS_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $HiCOS_EXE_ProductName} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
                    $HiPKILocalSignServer_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $HiPKILocalSignServer_Displayname}
                }
            }
    Uninstall_HiPKILocalSignServer $HiPKILocalSignServer_installeds
    Uninstall_Hicos $HiCOS_installeds
    & D:\HiCOS_Client.exe /install /passive /quiet /norestart


#>