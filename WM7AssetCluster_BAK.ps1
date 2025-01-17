#參數設定
    #程式路徑
        $WM7Asset_Path = "$env:SystemDrive\WM7Asset"
    #版本比較
        $WM7Asset_Newest_Setup_FileVersion = [version]"22.11.16.0"
        $Version_Old_NeedInstalled = $false
#判斷要不要安裝（sep之dat政策檔就是看三個檔案有無存在） 
<#
    檢查C槽下WM7Asset資料夾有無以下檔案:
    WM7Asset.bat
    WM7Assetreport.xml
    WM7LiteGreen.exe

#>               
    $WM7LiteGreen_Installed_Version
    $isInstall =!((Test-Path $WM7Asset_Path\WM7Asset.bat) -and (Test-Path $WM7Asset_Path\WM7Assetreport.xml) -and (Test-Path $WM7Asset_Path\WM7LiteGreen.exe))
    $WM7LiteGreen_Installed_FileVersion = [version](get-item -Path "$WM7Asset_Path\WM7LiteGreen.exe").VersionInfo.FileVersion
    #檢查版本有無過舊
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
                            FileDescription:  SMR應用程式
                            Product:          SMR應用程式
                            ProductVersion:   22.10.15.0
                            Debug:            False
                            Patched:          False
                            PreRelease:       False
                            PrivateBuild:     False
                            SpecialBuild:     False
                            Language:         中文 (繁體，台灣)
                    
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
        CreationTime      : 2023/1/6 下午 04:56:11
        CreationTimeUtc   : 2023/1/6 上午 08:56:11
        LastAccessTime    : 2023/1/6 下午 04:56:11
        LastAccessTimeUtc : 2023/1/6 上午 08:56:11
        LastWriteTime     : 2022/12/7 上午 11:16:26
        LastWriteTimeUtc  : 2022/12/7 上午 03:16:26
        Attributes        : Archive
    #>
#如果三個檔案都不存在則執行WM7AssetCluster.exe進行安裝
    if($isInstall){
        #檢查有無資料夾，如無則建立
            if(!(Test-Path $WM7Asset_Path)){New-Item -Path $WM7Asset_Path -ItemType Directory}
        #去download.moj下載WM7AssetCluster.exe
            $url_exe  = "http://download.moj/files/VANS/INTRA/WM7AssetCluster.exe"
            Start-Job -Name WebReq -ScriptBlock { param($p1, $p2)
                Invoke-WebRequest -Uri $p1 -OutFile $p2
            } -ArgumentList $url_exe,"$WM7Asset_Path\WM7AssetCluster.exe"
            Wait-Job -Name WebReq -Force
            Remove-Job -Name WebReq -Force
        #執行已下載完的WM7AssetCluster.exe以進行安裝
            Start-Process -FilePath "$WM7Asset_Path\WM7AssetCluster.exe"  -Wait
        #將WM7AssetCluster.exe刪除
            Remove-Item "$WM7Asset_Path\WM7AssetCluster.exe" -Force
    }