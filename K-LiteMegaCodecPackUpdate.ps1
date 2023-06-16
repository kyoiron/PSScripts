$RemoveFirstPC=@()
$KLiteMegaCodecPacks_Path = "\\172.29.205.114\loginscript\Update\KLiteMegaCodecPack"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$KLiteMegaCodecPack_EXE = (Get-ChildItem -Path ($KLiteMegaCodecPacks_Path+"\*.exe") | Where-Object{$_.VersionInfo.ProductName.trim() -eq "K-Lite Mega Codec Pack"} | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1)
$Force_Install = $false

#EXE安裝檔命令列參數請至https://jrsoftware.org/ishelp/index.php?topic=setupcmdline
<#
    FileVersionRaw     : 17.4.0.0
    ProductVersionRaw  : 17.4.0.0
    Comments           : This installation was built with Inno Setup.
    CompanyName        : KLCP                                                        
    FileBuildPart      : 0
    FileDescription    : K-Lite Mega Codec Pack Setup                                
    FileMajorPart      : 17
    FileMinorPart      : 4
    FileName           : \\172.29.205.114\loginscript\Update\KLiteMegaCodecPack\K-Lite_Codec_Pack_1740_Mega.exe
    FilePrivatePart    : 0
    FileVersion        : 17.4.0.0            
    InternalName       : 
    IsDebug            : False
    IsPatched          : False
    IsPrivateBuild     : False
    IsPreRelease       : False
    IsSpecialBuild     : False
    Language           : 中性語言
    LegalCopyright     :                                                                                                     
    LegalTrademarks    : 
    OriginalFilename   :                                                   
    PrivateBuild       : 
    ProductBuildPart   : 0
    ProductMajorPart   : 17
    ProductMinorPart   : 4
    ProductName        : K-Lite Mega Codec Pack                                      
    ProductPrivatePart : 0
    ProductVersion     : 17.4.0                                            
    SpecialBuild       : 
 #>
$KLiteMegaCodecPack_EXE_Path = $KLiteMegaCodecPack_EXE.FullName
if($KLiteMegaCodecPack_EXE_Path){
    $KLiteMegaCodecPack_EXE_ProductName = (Get-ItemProperty $KLiteMegaCodecPack_EXE_Path).VersionInfo.ProductName.trim()
    $KLiteMegaCodecPack_EXE_ProductVersion = (Get-ItemProperty $KLiteMegaCodecPack_EXE_Path).VersionInfo.ProductVersion.trim()
    <#
        Inno Setup: Setup Version         : 6.0.4 (u)
        Inno Setup: App Path              : C:\Program Files (x86)\K-Lite Codec Pack
        InstallLocation                   : C:\Program Files (x86)\K-Lite Codec Pack\
        Inno Setup: Icon Group            : K-Lite Codec Pack
        Inno Setup: User                  : tndadmin
        Inno Setup: Setup Type            : std_play_encoding
        Inno Setup: Selected Components   : choice,choice\best,player,player\mpchc,player\mpchc\x64,video,video\microsoft,video\microsoft\vc1,video\microsoft\wmv,video\lav,video\lav\hevc,video\lav\h264,video\lav\
                                            mpeg4,video\lav\mpeg2,video\lav\other,audio,audio\microsoft,audio\microsoft\wma,audio\lav,audio\lav\ac3dts,audio\lav\truehd,audio\lav\aac,audio\lav\flac,audio\lav\mpeg,
                                            audio\lav\other,sourcefilter,sourcefilter\microsoft,sourcefilter\microsoft\avi,sourcefilter\microsoft\mpegps,sourcefilter\microsoft\wmv,sourcefilter\lav,sourcefilter\la
                                            v\matroska,sourcefilter\lav\mp4,sourcefilter\lav\mpegts,sourcefilter\lav\other,subtitles,subtitles\vsfilter,subtitles\xysubfilter,other,other\madvr,other\mpcvr,videovfw
                                            ,videovfw\x264,videovfw\xvid,videovfw\lagarith,tools,tools\codectweaktool,tools\mediainfo,shell,shell\icaros_thumbnail,shell\icaros_property,misc,misc\brokencodecs,misc
                                            \brokenfilters
        Inno Setup: Deselected Components : choice\note,choice\custom,choice\note2,player\mpchc\x86,video\microsoft\note,video\lav\note,video\lav\vc1,video\lav\wmv,video\ffproc,video\ffproc\manual,video\ffproc\mp
                                            c,video\ffdshow,video\ffdshow\note,video\ffdshow\h264,video\ffdshow\mpeg4,video\ffdshow\mpeg2,video\ffdshow\mpeg2\libavcodec,video\ffdshow\mpeg2\libmpeg2,video\ffdshow\
                                            vc1,video\ffdshow\other,video\ffdshow\raw,video\xvid,video\xvid\note,video\xvid\mpeg4,audio\microsoft\note,audio\lav\note,audio\lav\wma,audio\ffproc,audio\ffproc\manual
                                            ,audio\ffproc\mpc,audio\ffdshow,audio\ffdshow\note,audio\ffdshow\ac3dts,audio\ffdshow\aac,audio\ffdshow\flac,audio\ffdshow\mpeg,audio\ffdshow\other,audio\ffdshow\pcm,au
                                            dio\ac3filter,audio\ac3filter\note,audio\ac3filter\ac3dts,audio\ac3filter\aac,audio\ac3filter\flac,audio\ac3filter\mpeg,audio\ac3filter\pcm,sourcefilter\lav\avi,sourcef
                                            ilter\lav\mpegps,sourcefilter\lav\wmv,sourcefilter\haali,sourcefilter\haali\note,sourcefilter\haali\matroska,sourcefilter\haali\mp4,sourcefilter\haali\mpegts,sourcefilt
                                            er\dcbass,sourcefilter\dcbass\note,sourcefilter\dcbass\shoutcast,sourcefilter\dcbass\optimfrog,sourcefilter\dcbass\tracker,other\haalirenderer,videovfw\note,videovfw\hu
                                            ffyuv,videovfw\ffdshow,audioacm,audioacm\note,audioacm\mp3lame,audioacm\ac3acm,tools\graphstudio,tools\haalimuxer,tools\vobsubstrip,tools\fourcc
        Inno Setup: Selected Tasks        : mpc_sendto,config_shortcuts,systray_lavsplitter,systray_lav,systray_madvr,wmp_reg_formats,adjust_preferred_decoders,mediainfo_contextmenu,h264mvc,checknews
        Inno Setup: Deselected Tasks      : reset_settings,mpc_desktop,ctt_desktop,mediainfo_sendto,lav_buffer_increase,use_lav_for_http,use_lav_for_https,no_thumb_overlays,icaros_cache,previewpane,update,update\
                                            d7,update\d14,update\d30,update\d90,update\minor
        Inno Setup: Language              : en
        DisplayName                       : K-Lite Mega Codec Pack 15.6.0
        DisplayIcon                       : C:\Program Files (x86)\K-Lite Codec Pack\unins000.exe
        UninstallString                   : "C:\Program Files (x86)\K-Lite Codec Pack\unins000.exe"
        QuietUninstallString              : "C:\Program Files (x86)\K-Lite Codec Pack\unins000.exe" /SILENT
        DisplayVersion                    : 15.6.0
        Publisher                         : KLCP
        NoModify                          : 1
        NoRepair                          : 1
        InstallDate                       : 20200729
        MajorVersion                      : 15
        MinorVersion                      : 6
        VersionMajor                      : 15
        VersionMinor                      : 6
        EstimatedSize                     : 177121
        sEstimatedSize2                   : 165525
        PSPath                            : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\KLiteCodecPack_is1
        PSParentPath                      : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName                       : KLiteCodecPack_is1
        PSDrive                           : HKLM
        PSProvider                        : Microsoft.PowerShell.Core\Registry      
    #>

    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $KLiteCodecPack_installeds += Get-ItemProperty $Path | Where-Object{($_.DisplayName -eq ($KLiteMegaCodecPack_EXE_ProductName+" "+$_.DisplayVersion)) -or ($_.DisplayName -like "K-Lite Codec Pack*" )} 
        }
    }
    if($KLiteCodecPack_installeds -or $Force_Install){
        if($RemoveFirstPC.Contains($env:Computername) -or (($KLiteCodecPack_installeds|Measure-Object).count -ge 2)){
            foreach($item in $KLiteCodecPack_installeds){
                $uninstall_Char = ($item.UninstallString -split "  ")
                #$LogFile = "$env:systemdrive\temp\"+$env:Computername + "_Uninstall_"+ $item.DisplayName.ToString() + ".txt"
                $LogFile = "$env:systemdrive\temp\"+$env:Computername + "_"+ $item.DisplayName.Replace($item.DisplayVersion,"").trim() +"_Uninstall_"+ $item.DisplayVersion + ".txt"
                if(test-path $LogFile){
                    $StartDate=(GET-DATE)
                    $EndDate=(Get-ItemProperty -Path $LogFile).LastWriteTime
                    $diff_Value = (NEW-TIMESPAN –Start $StartDate –End $EndDate).Days *-1
                    $diff_day = 3
                    $NEED_Remove = $false
                }else{
                    $NEED_Remove = $true
                }
                if( $NEED_Remove -or ($diff_Value -gt $diff_day)){
                    $arguments = " /VERYSILENT /NORESTART /LOG=""" + $LogFile+""""
                    start-process $uninstall_Char[0] -arg $arguments -Wait -WindowStyle Hidden 
                    $Force_Install = $true                                       
                }
             }
            $Log_Folder_Path = $Log_Path +"\"+ $KLiteMegaCodecPack_EXE_ProductName
            $LogPattern =$env:Computername + "_"+$KLiteMegaCodecPack_EXE_ProductName+"_*.txt"
            if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
            if(Test-Path -Path $LogFile){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
        }

        $KLiteCodecPack_installed = $KLiteCodecPack_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
        if(([version]$KLiteCodecPack_installed.DisplayVersion -ge [version]$KLiteMegaCodecPack_EXE_ProductVersion) -and ($Force_Install -ne $true)){exit}        
        $LogName = $env:Computername + "_"+$KLiteMegaCodecPack_EXE_ProductName+"_"+ $KLiteMegaCodecPack_EXE_ProductVersion + ".txt"
        $arguments = " /VERYSILENT /NORESTART /LOG=$env:systemdrive\temp\""$LogName"""
        robocopy $KLiteMegaCodecPacks_Path "$env:systemdrive\temp" $KLiteMegaCodecPack_EXE.Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        unblock-file ($env:systemdrive+"\temp\"+$KLiteMegaCodecPack_EXE.Name)
        start-process ($env:systemdrive+"\temp\"+$KLiteMegaCodecPack_EXE.Name) -arg $arguments -WindowStyle Hidden 
        Start-Sleep -s 15
        $Log_Folder_Path = $Log_Path +"\"+ $KLiteMegaCodecPack_EXE_ProductName
        $LogPattern =$env:Computername + "_"+$KLiteMegaCodecPack_EXE_ProductName+"_*.txt"
        if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
        if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}    
    }
}