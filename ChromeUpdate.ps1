function Get-FileMetaData {
    <#
    .SYNOPSIS
    Small function that gets metadata information from file providing similar output to what Explorer shows when viewing file

    .DESCRIPTION
    Small function that gets metadata information from file providing similar output to what Explorer shows when viewing file

    .PARAMETER File
    FileName or FileObject

    .EXAMPLE
    Get-ChildItem -Path $Env:USERPROFILE\Desktop -Force | Get-FileMetaData | Out-HtmlView -ScrollX -Filtering -AllProperties

    .EXAMPLE
    Get-ChildItem -Path $Env:USERPROFILE\Desktop -Force | Where-Object { $_.Attributes -like '*Hidden*' } | Get-FileMetaData | Out-HtmlView -ScrollX -Filtering -AllProperties

    .NOTES
    #>
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline)][Object] $File,
        [switch] $Signature
    )
    Process {
        foreach ($F in $File) {
            $MetaDataObject = [ordered] @{}
            if ($F -is [string]) {
                $FileInformation = Get-ItemProperty -Path $F
            } elseif ($F -is [System.IO.DirectoryInfo]) {
                #Write-Warning "Get-FileMetaData - Directories are not supported. Skipping $F."
                continue
            } elseif ($F -is [System.IO.FileInfo]) {
                $FileInformation = $F
            } else {
                Write-Warning "Get-FileMetaData - Only files are supported. Skipping $F."
                continue
            }
            $ShellApplication = New-Object -ComObject Shell.Application
            $ShellFolder = $ShellApplication.Namespace($FileInformation.Directory.FullName)
            $ShellFile = $ShellFolder.ParseName($FileInformation.Name)
            $MetaDataProperties = [ordered] @{}
            0..400 | ForEach-Object -Process {
                $DataValue = $ShellFolder.GetDetailsOf($null, $_)
                $PropertyValue = (Get-Culture).TextInfo.ToTitleCase($DataValue.Trim()).Replace(' ', '')
                if ($PropertyValue -ne '') {
                    $MetaDataProperties["$_"] = $PropertyValue
                }
            }
            foreach ($Key in $MetaDataProperties.Keys) {
                $Property = $MetaDataProperties[$Key]
                $Value = $ShellFolder.GetDetailsOf($ShellFile, [int] $Key)
                if ($Property -in 'Attributes', 'Folder', 'Type', 'SpaceFree', 'TotalSize', 'SpaceUsed') {
                    continue
                }
                If (($null -ne $Value) -and ($Value -ne '')) {
                    $MetaDataObject["$Property"] = $Value
                }
            }
            if ($FileInformation.VersionInfo) {
                $SplitInfo = ([string] $FileInformation.VersionInfo).Split([char]13)
                foreach ($Item in $SplitInfo) {
                    $Property = $Item.Split(":").Trim()
                    if ($Property[0] -and $Property[1] -ne '') {
                        $MetaDataObject["$($Property[0])"] = $Property[1]
                    }
                }
            }
            $MetaDataObject["Attributes"] = $FileInformation.Attributes
            $MetaDataObject['IsReadOnly'] = $FileInformation.IsReadOnly
            $MetaDataObject['IsHidden'] = $FileInformation.Attributes -like '*Hidden*'
            $MetaDataObject['IsSystem'] = $FileInformation.Attributes -like '*System*'
            if ($Signature) {
                $DigitalSignature = Get-AuthenticodeSignature -FilePath $FileInformation.Fullname
                $MetaDataObject['SignatureCertificateSubject'] = $DigitalSignature.SignerCertificate.Subject
                $MetaDataObject['SignatureCertificateIssuer'] = $DigitalSignature.SignerCertificate.Issuer
                $MetaDataObject['SignatureCertificateSerialNumber'] = $DigitalSignature.SignerCertificate.SerialNumber
                $MetaDataObject['SignatureCertificateNotBefore'] = $DigitalSignature.SignerCertificate.NotBefore
                $MetaDataObject['SignatureCertificateNotAfter'] = $DigitalSignature.SignerCertificate.NotAfter
                $MetaDataObject['SignatureCertificateThumbprint'] = $DigitalSignature.SignerCertificate.Thumbprint
                $MetaDataObject['SignatureStatus'] = $DigitalSignature.Status
                $MetaDataObject['IsOSBinary'] = $DigitalSignature.IsOSBinary
            }
            [PSCustomObject] $MetaDataObject
        }
    }
}
#LOG檔NAS存放路徑
$Log_Path = "\\172.29.205.114\Public\sources\audit"
#MSI檔安裝路徑
$Chrome_MSI_Folder="\\172.29.205.114\loginscript\Update\Chrome"
#沒裝的是否要裝$true或$false
$Install_IF_NOT_Installed = $true
if([System.Environment]::Is64BitOperatingSystem){
    $Chrome_MIS_FileMetaData = Get-FileMetaData -File ($Chrome_MSI_Folder+"\GoogleChromeStandaloneEnterprise64.msi") 
    #$Chrome_MIS =  Get-MsiInformation -Path ($Chrome_MSI_Folder+"\GoogleChromeStandaloneEnterprise64.msi") 
}else{
    $Chrome_MIS_FileMetaData = Get-FileMetaData -File ($Chrome_MSI_Folder+"\GoogleChromeStandaloneEnterprise.msi") 
    #$Chrome_MIS =  Get-MsiInformation -Path ($Chrome_MSI_Folder+"\GoogleChromeStandaloneEnterprise.msi") 
}
<#
    Get-MsiInformation chrome msi
    File            : \\172.29.205.114\loginscript\Update\Chrome\GoogleChromeStandaloneEnterprise64.msi
    ProductCode     : {B01A8859-9D45-3472-AD5D-0FB367564035}
    Manufacturer    : Google LLC
    ProductName     : Google Chrome
    ProductVersion  : 68.21.49235
    ProductLanguage : 1033


    Get-FileMetaData chrome msi
    名稱           : GoogleChromeStandaloneEnterprise64.msi
    大小           : 65.8 MB
    項目類型         : Windows Installer 封裝
    修改日期         : 2020/8/22 下午 06:05
    建立日期         : 2020/9/2 下午 05:18
    存取日期         : 2020/9/2 下午 05:21
    屬性           : A
    認知類型         : 應用程式
    擁有者          : MOJ\kyoiron
    種類           : 程式
    標籤           : Installer
    評等           : 未評等
    作者           : Google LLC
    標題           : Installation Database
    主旨           : Google Chrome Installer
    註解           : 85.0.4183.83 Copyright 2020 Google LLC
    程式名稱         : Windows Installer XML Toolset (3.8.1128.0)
    大小總計         : 0.99 TB
    電腦           : 172.29.205.114
    已建立的本文       : ‎2020/‎8/‎23 ‏‎上午 09:03
    上次儲存日期       : ‎2020/‎8/‎23 ‏‎上午 09:03
    頁數           : 200
    字數統計         : 2
    副檔名          : .msi
    檔名           : GoogleChromeStandaloneEnterprise64.msi
    可用空間         : 255 GB
    資料夾名稱        : Chrome
    資料夾路徑        : \\172.29.205.114\loginscript\Update\Chrome
    資料夾          : Chrome (\\172.29.205.114\loginscript\Update)
    路徑           : \\172.29.205.114\loginscript\Update\Chrome\GoogleChromeStandaloneEnterprise64.msi
    類型           : Windows Installer 封裝
    連結狀態         : UNRESOLVED
    已使用空間        : ‎74%
    共用狀態         : 不分享
    File         : \\172.29.205.114\loginscript\Update\Chrome\GoogleChromeStandaloneEnterprise64.msi
    Debug        : False
    Patched      : False
    PreRelease   : False
    PrivateBuild : False
    SpecialBuild : False
    Attributes   : Archive
    IsReadOnly   : False
    IsHidden     : False
    IsSystem     : False   
#>
if($Chrome_MIS_FileMetaData){
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $Chrome_installeds = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $Chrome_installeds += Get-ItemProperty $Path | Where-Object{$_.DisplayName -eq "Google Chrome"}  
        }
        
    }
    <#
        Get-ItemProperty Google Chrome
        AuthorizedCDFPrefix : 
        Comments            : 
        Contact             : 
        DisplayVersion      : 84.0.4147.135
        HelpLink            : 
        HelpTelephone       : 
        InstallDate         : 20200420
        InstallLocation     : 
        InstallSource       : C:\temp\
        ModifyPath          : MsiExec.exe /X{78831B61-87DE-3660-9687-A541FD017EA9}
        NoModify            : 1
        Publisher           : Google LLC
        Readme              : 
        Size                : 
        EstimatedSize       : 61126
        UninstallString     : MsiExec.exe /X{78831B61-87DE-3660-9687-A541FD017EA9}
        URLInfoAbout        : 
        URLUpdateInfo       : 
        VersionMajor        : 67
        VersionMinor        : 243
        WindowsInstaller    : 1
        Version             : 1139998833
        Language            : 1033
        DisplayName         : Google Chrome
        sEstimatedSize2     : 60817
        PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{78831B61-87DE-3660-9687-A541FD017EA9}
        PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName         : {78831B61-87DE-3660-9687-A541FD017EA9}
        PSDrive             : HKLM
        PSProvider          : Microsoft.PowerShell.Core\Registry
    #>
    $Chrome_installed = $Chrome_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    if(($Chrome_installed -eq $null) -and ($Install_IF_NOT_Installed -ne $true)){exit}
    $Chrome_MSI_Version = [version]($Chrome_MIS_FileMetaData.註解 -split "Copyright")[0].trim()
    if([version]$Chrome_installed.DisplayVersion -ge $Chrome_MSI_Version){exit}        
    $LogName = $env:Computername + "_"+ "Google Chrome" +"_"+ $Chrome_MSI_Version + ".txt"
    $Chrome_MIS_fileName = (Get-ChildItem -Path $Chrome_MIS_FileMetaData.File).Name
    $arguments = "/i $env:systemdrive\temp\$Chrome_MIS_fileName /qn /log ""$env:systemdrive\temp\" +  $LogName+""""
    robocopy $Chrome_MSI_Folder "$env:systemdrive\temp" $Chrome_MIS_fileName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    $Log_Folder_Path = $Log_Path +"\Google Chrome"
    start-process "msiexec" -arg $arguments -Wait
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
}