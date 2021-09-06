#$pdis_EXE_Path = "\\172.29.205.114\loginscript\Update\pdis"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$RegUninstallPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
#$pdis_EXE = Get-ChildItem -Path ($pdis_EXE_Path+"\pdis_*.exe") | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1
#$pdis_EXE_Version =  $pdis_EXE.BaseName.TrimStart("pdis_v")
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$Web_pdis = Invoke-WebRequest -UseBasicParsing -Uri "https://pdis.moj.gov.tw/U100/U101-1.aspx" -WebSession $session
$pdis_web_EXE = ($Web_pdis.Links.FindById('ctl00_ContentPlaceHolder1_HyperLink_EXE').href).trimstart("../version/")
$pdis_web_version = [int]$pdis_web_EXE.TrimStart("pdis_v").TrimEnd(".exe")

foreach ($Path in $RegUninstallPaths) {
    if (Test-Path $Path) {
        $pdis_installed = Get-ItemProperty $Path | Where-Object{$_.DisplayName -match "公職人員財產申報系統"} 
    }
}
if(Test-Path($pdis_installeds.InstallLocation + "Ins_Apply.ver")){
    $pdis_installed_Version = [int]((Get-Content ($pdis_installeds.InstallLocation + "Ins_Apply.ver") -ErrorAction Continue).TrimStart("v") )      
}
if(($pdis_web_version -gt $pdis_installed_Version)-or(!$pdis_installed)){
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        Invoke-WebRequest -UseBasicParsing -Uri "https://pdis.moj.gov.tw/version/pdis_v1658.exe" -WebSession $session  -OutFile "$env:systemdrive\temp\$pdis_web_EXE"    
        #robocopy $pdis_EXE_Path "$env:systemdrive\temp" ""$pdis_EXE.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
        unblock-file ($env:systemdrive+"\temp\"+ $pdis_web_EXE)
        $pdis_EXE = Get-ItemProperty -Path ($env:systemdrive+"\temp\"+ $pdis_web_EXE)
        $pdis_EXE_Version =  $pdis_EXE.BaseName.TrimStart("pdis_v")
        $LogName = $env:Computername + "_"+$pdis_EXE.VersionInfo.ProductName.trim()+"_v"+ $pdis_EXE_Version + ".txt"
        $arguments = "/VERYSILENT /NORESTART /ALLUSERS /SP- /LOG=""$env:systemdrive\temp\$LogName"""
        start-process ($env:systemdrive+"\temp\"+ $pdis_web_EXE) -arg $arguments        
        do{
            Start-Sleep -s 1
        }while( !(Get-Content -Path "$env:systemdrive\temp\$LogName" | Where-Object {$_.Contains("Log closed.")}) )
        #設定捷徑
        if(!(test-path -path "$env:PUBLIC\Desktop\公職人員財產申報.lnk")){
           robocopy "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\公職人員財產申報系統" "$env:PUBLIC\Desktop" "公職人員財產申報.lnk" /XO /NJH /NJS /NDL /NC /NS
        }
        #回傳安裝LOG檔
        $Log_Folder_Path = $Log_Path +"\"+ "pdis"
        if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
        if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName /XO /NJH /NJS /NDL /NC /NS }
}    
<#
if($pdis_EXE.FullName){
    $pdis_EXE_ProductName = $pdis_EXE.VersionInfo.Product
    $pdis_EXE_ProductVersion = $pdis_EXE.VersionInfo.ProductVersion

    <#
        PSPath            : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\pdis\pdis_v1658.exe
        PSParentPath      : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\pdis
        PSChildName       : pdis_v1658.exe
        PSProvider        : Microsoft.PowerShell.Core\FileSystem
        PSIsContainer     : False
        Mode              : -a----
        VersionInfo       : File:             \\172.29.205.114\loginscript\Update\pdis\pdis_v1658.exe
                            InternalName:     
                            OriginalFilename: 
                            FileVersion:                          
                            FileDescription:  公職人員財產申報系統 Setup                                            
                            Product:          公職人員財產申報系統                                                  
                            ProductVersion:                                                     
                            Debug:            False
                            Patched:          False
                            PreRelease:       False
                            PrivateBuild:     False
                            SpecialBuild:     False
                            Language:         中性語言
                            
        BaseName          : pdis_v1658
        Target            : 
        LinkType          : 
        Name              : pdis_v1658.exe
        Length            : 9195614
        DirectoryName     : \\172.29.205.114\loginscript\Update\pdis
        Directory         : \\172.29.205.114\loginscript\Update\pdis
        IsReadOnly        : False
        Exists            : True
        FullName          : \\172.29.205.114\loginscript\Update\pdis\pdis_v1658.exe
        Extension         : .exe
        CreationTime      : 2021/9/6 上午 10:02:25
        CreationTimeUtc   : 2021/9/6 上午 02:02:25
        LastAccessTime    : 2021/9/6 上午 10:30:38
        LastAccessTimeUtc : 2021/9/6 上午 02:30:38
        LastWriteTime     : 2021/9/6 上午 09:39:49
        LastWriteTimeUtc  : 2021/9/6 上午 01:39:49
        Attributes        : Archive
    
    #$pdis_EXE_installeds = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $pdis_installed = Get-ItemProperty $Path | Where-Object{$_.DisplayName -match "公職人員財產申報系統"} 
        }
    }
    if(Test-Path($pdis_installeds.InstallLocation + "Ins_Apply.ver")){
        $pdis_installed_Version = [int]((Get-Content ($pdis_installeds.InstallLocation + "Ins_Apply.ver") -ErrorAction Continue).TrimStart("v") )      
    }
    if(($pdis_EXE_Version -ge $pdis_installed_Version)-or(!$pdis_installed)){
        robocopy $pdis_EXE_Path "$env:systemdrive\temp" ""$pdis_EXE.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
        unblock-file ($env:systemdrive+"\temp\"+ $pdis_EXE.Name)            
        $LogName = $env:Computername + "_"+$pdis_EXE.VersionInfo.ProductName.trim()+"_"+ $pdis_EXE_Version + ".txt"
        $arguments = "/VERYSILENT /NORESTART ALLUSERS /SUPPRESSMSGBOXES /LOG=""$env:systemdrive\temp\$LogName"""
        start-process ($env:systemdrive+"\temp\"+ $pdis_EXE.Name) -arg $arguments -wait
    }    
}
#>