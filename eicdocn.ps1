$url_eicdocnAt19 = "http://172.31.3.55/kw/common/comp/eicdocn@19.cab"
$eic_NAS_Path = "\\172.29.205.114\loginscript\Update\eic"
if(Test-Path "$env:windir\SysWOW64"){$eic_PC_Path="$env:windir\SysWOW64"}else{$eic_PC_Path="$env:windir\system32"}
#$url_eicdocn = "http://172.29.204.56:8080/TBKN/home/eicdocn.cab"
#Invoke-WebRequest -Uri $url_eicdocnAt19 -OutFile "$env:SystemDrive\temp\eicdocn@19.cab" 
#Invoke-WebRequest -Uri $url_eicdocn -OutFile "$env:SystemDrive\temp\eicdocn.cab"
#DISM /Online /Add-Package /PackagePath:"$env:SystemDrive\temp\eicdocn@19.cab"
#DISM /Online /Add-Package /PackagePath:"$env:SystemDrive\temp\eicdocn.cab"
$eicdocnDLL_NAS = Get-ChildItem -Path ($eic_NAS_Path+"\eicdocn.dll")
$eicdocnDLL_PC = Get-ChildItem -Path ($eic_PC_Path+"\eicdocn.dll")
<#
    PSPath            : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\eic\eicdocn.dll
    PSParentPath      : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\eic
    PSChildName       : eicdocn.dll
    PSProvider        : Microsoft.PowerShell.Core\FileSystem
    PSIsContainer     : False
    Mode              : -a----
    VersionInfo       : File:             \\172.29.205.114\loginscript\Update\eic\eicdocn.dll
                        InternalName:     EicDocN.dll
                        OriginalFilename: EicDocN.dll
                        FileVersion:      3, 2, 0, 46
                        FileDescription:  EicDocN Module
                        Product:          EicDocN Module
                        ProductVersion:   3, 2, 0, 46
                        Debug:            False
                        Patched:          False
                        PreRelease:       False
                        PrivateBuild:     False
                        SpecialBuild:     False
                        Language:         中文 (繁體，台灣)
                        
    BaseName          : eicdocn
    Target            : 
    LinkType          : 
    Name              : eicdocn.dll
    Length            : 367288
    DirectoryName     : \\172.29.205.114\loginscript\Update\eic
    Directory         : \\172.29.205.114\loginscript\Update\eic
    IsReadOnly        : False
    Exists            : True
    FullName          : \\172.29.205.114\loginscript\Update\eic\eicdocn.dll
    Extension         : .dll
    CreationTime      : 2020/8/27 下午 03:45:01
    CreationTimeUtc   : 2020/8/27 上午 07:45:01
    LastAccessTime    : 2020/8/27 下午 03:45:01
    LastAccessTimeUtc : 2020/8/27 上午 07:45:01
    LastWriteTime     : 2020/1/31 下午 02:37:56
    LastWriteTimeUtc  : 2020/1/31 上午 06:37:56
    Attributes        : Archive
#>
if([version]$eicdocnDLL_NAS.VersionInfo.ProductVersionRaw  -gt [version]$eicdocnDLL_PC.VersionInfo.ProductVersionRaw){Copy-Item -Path $eicdocnDLL_NAS -Destination $eic_PC_Path -Force}
#if(Test-Path "$env:SystemDrive\temp\eicdocn@19.cab"){DISM /Online /Add-Package /PackagePath:"$env:SystemDrive\temp\eicdocn@19.cab"}