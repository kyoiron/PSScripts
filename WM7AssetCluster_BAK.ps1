#�ѼƳ]�w
    #�{�����|
        $WM7Asset_Path = "$env:SystemDrive\WM7Asset"
    #�������
        $WM7Asset_Newest_Setup_FileVersion = [version]"22.11.16.0"
        $Version_Old_NeedInstalled = $false
#�P�_�n���n�w�ˡ]sep��dat�F���ɴN�O�ݤT���ɮצ��L�s�b�^ 
<#
    �ˬdC�ѤUWM7Asset��Ƨ����L�H�U�ɮ�:
    WM7Asset.bat
    WM7Assetreport.xml
    WM7LiteGreen.exe

#>               
    $WM7LiteGreen_Installed_Version
    $isInstall =!((Test-Path $WM7Asset_Path\WM7Asset.bat) -and (Test-Path $WM7Asset_Path\WM7Assetreport.xml) -and (Test-Path $WM7Asset_Path\WM7LiteGreen.exe))
    $WM7LiteGreen_Installed_FileVersion = [version](get-item -Path "$WM7Asset_Path\WM7LiteGreen.exe").VersionInfo.FileVersion
    #�ˬd�������L�L��
    if($WM7LiteGreen_Installed_FileVersion -lt $WM7Asset_Newest_Setup_FileVersion){
        $Version_Old_NeedInstalled = $true
        $isInstall = $true
     }
    <#
        PSPath            : Microsoft.PowerShell.Core\FileSystem::C:\WM7Asset\WM7LiteGreen.exe
        PSParentPath      : Microsoft.PowerShell.Core\FileSystem::C:\WM7Asset
        PSChildName       : WM7LiteGreen.exe
        PSDrive           : C
        PSProvider        : Microsoft.PowerShell.Core\FileSystem
        PSIsContainer     : False
        Mode              : -a----
        VersionInfo       : File:             C:\WM7Asset\WM7LiteGreen.exe
                            InternalName:     WM7Lite.exe
                            OriginalFilename: WM7Lite.exe
                            FileVersion:      22.11.16.0
                            FileDescription:  SMR���ε{��
                            Product:          SMR���ε{��
                            ProductVersion:   22.10.15.0
                            Debug:            False
                            Patched:          False
                            PreRelease:       False
                            PrivateBuild:     False
                            SpecialBuild:     False
                            Language:         ���� (�c��A�x�W)
                    
        BaseName          : WM7LiteGreen
        Target            : {}
        LinkType          : 
        Name              : WM7LiteGreen.exe
        Length            : 1101288
        DirectoryName     : C:\WM7Asset
        Directory         : C:\WM7Asset
        IsReadOnly        : False
        Exists            : True
        FullName          : C:\WM7Asset\WM7LiteGreen.exe
        Extension         : .exe
        CreationTime      : 2023/1/6 �U�� 04:56:11
        CreationTimeUtc   : 2023/1/6 �W�� 08:56:11
        LastAccessTime    : 2023/1/6 �U�� 04:56:11
        LastAccessTimeUtc : 2023/1/6 �W�� 08:56:11
        LastWriteTime     : 2022/12/7 �W�� 11:16:26
        LastWriteTimeUtc  : 2022/12/7 �W�� 03:16:26
        Attributes        : Archive
    #>
#�p�G�T���ɮ׳����s�b�h����WM7AssetCluster.exe�i��w��
    if($isInstall){
        #�ˬd���L��Ƨ��A�p�L�h�إ�
            if(!(Test-Path $WM7Asset_Path)){New-Item -Path $WM7Asset_Path -ItemType Directory}
        #�hdownload.moj�U��WM7AssetCluster.exe
            $url_exe  = "http://download.moj/files/VANS/INTRA/WM7AssetCluster.exe"
            Start-Job -Name WebReq -ScriptBlock { param($p1, $p2)
                Invoke-WebRequest -Uri $p1 -OutFile $p2
            } -ArgumentList $url_exe,"$WM7Asset_Path\WM7AssetCluster.exe"
            Wait-Job -Name WebReq -Force
            Remove-Job -Name WebReq -Force
        #����w�U������WM7AssetCluster.exe�H�i��w��
            Start-Process -FilePath "$WM7Asset_Path\WM7AssetCluster.exe"  -Wait
        #�NWM7AssetCluster.exe�R��
            Remove-Item "$WM7Asset_Path\WM7AssetCluster.exe" -Force
    }