$RemoveFirstPC=@()
$XnViews_Path = "\\172.29.205.114\loginscript\Update\XnView"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$XnView_EXE = (Get-ChildItem -Path ($XnViews_Path+"\*.exe") | Where-Object{$_.VersionInfo.ProductName.trim() -eq "XnView"} | Sort-Object)
$Force_Install = $false
#EXE安裝檔命令列參數請至https://jrsoftware.org/ishelp/index.php?topic=setupcmdline
<#
    FileVersionRaw     : 2.51.1.0
    ProductVersionRaw  : 2.51.1.0
    Comments           : This installation was built with Inno Setup.
    CompanyName        : Gougelet Pierre-e                                           
    FileBuildPart      : 1
    FileDescription    : XnView Setup                                                
    FileMajorPart      : 2
    FileMinorPart      : 51
    FileName           : \\172.29.205.114\loginscript\Update\XnView\XnView-win-full.exe
    FilePrivatePart    : 0
    FileVersion        : 2.51.1              
    InternalName       : 
    IsDebug            : False
    IsPatched          : False
    IsPrivateBuild     : False
    IsPreRelease       : False
    IsSpecialBuild     : False
    Language           : 中性語言
    LegalCopyright     : Copyright © 1991-2022 Pierre-e Gougelet                                                             
    LegalTrademarks    : 
    OriginalFilename   : 
    PrivateBuild       : 
    ProductBuildPart   : 1
    ProductMajorPart   : 2
    ProductMinorPart   : 51
    +ProductName        : XnView                                                      
    ProductPrivatePart : 0
    ProductVersion     : 2.51.1                                            
    SpecialBuild       : 
#>
$XnView_EXE_Path = $XnView_EXE.FullName
if($XnView_EXE_Path){
    $XnView_EXE_ProductName = (Get-ItemProperty $XnView_EXE_Path).VersionInfo.ProductName.trim()
    $XnView_EXE_ProductVersion = (Get-ItemProperty $XnView_EXE_Path).VersionInfo.ProductVersion.trim()
    <#
        Inno Setup: Setup Version         : 5.5.8 (a)
        Inno Setup: App Path              : C:\Program Files (x86)\XnView
        InstallLocation                   : C:\Program Files (x86)\XnView\
        Inno Setup: Icon Group            : XnView
        Inno Setup: User                  : tndadmin
        Inno Setup: Setup Type            : full
        Inno Setup: Selected Components   : mask,paint,xjp2,xmp,nero,mpeg,ftp,webp
        Inno Setup: Deselected Components : 
        Inno Setup: Selected Tasks        : 
        Inno Setup: Deselected Tasks      : desktopicon,quicklaunchicon
        Inno Setup: Language              : en
        DisplayName                       : XnView 2.51.1
        UninstallString                   : "C:\Program Files (x86)\XnView\unins000.exe"
        QuietUninstallString              : "C:\Program Files (x86)\XnView\unins000.exe" /SILENT
        DisplayVersion                    : 2.51.1
        Publisher                         : Gougelet Pierre-e
        URLInfoAbout                      : http://www.xnview.com
        NoModify                          : 1
        NoRepair                          : 1
        InstallDate                       : 20230109
        MajorVersion                      : 2
        MinorVersion                      : 51
        EstimatedSize                     : 21233
        sEstimatedSize2                   : 20649
        PSPath                            : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\XnView_is1
        PSParentPath                      : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName                       : XnView_is1
        PSDrive                           : HKLM
        PSProvider                        : Microsoft.PowerShell.Core\Registry        
    #>

    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $XnView_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -eq ($XnView_EXE_ProductName+" "+$_.DisplayVersion)} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
        }
    }
    if($XnView_installeds -or $Force_Install){
        if($RemoveFirstPC.Contains($env:Computername) -or (($XnView_installeds|Measure-Object).count -ge 2)){
            foreach($item in $XnView_installeds){
                $uninstall_Char = ($item.UninstallString -split "  ")
                $LogFile= "$env:systemdrive\temp\"+$env:Computername +"_"+ $item.DisplayName.Replace($item.DisplayVersion,"").trim() +"_Uninstall_"+ $item.DisplayVersion + ".txt"
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
                    $arguments = " /VERYSILENT /NORESTART /ALLUSERS /NOICONS /LOG=""" + $LogFile  + """"
                    start-process $uninstall_Char[0] -arg $arguments -Wait -WindowStyle Hidden   
                    $Force_Install = $true                 
                }
             }
            $Log_Folder_Path = $Log_Path +"\"+ $XnView_EXE_ProductName
            $LogPattern =$env:Computername + "_"+$XnView_EXE_ProductName+"_*.txt"
            if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
            if(Test-Path -Path $LogFile){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
        }

        $XnView_installed = $XnView_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
        if(([version]$XnView_installed.DisplayVersion -ge [version]$XnView_EXE_ProductVersion) -and ($Force_Install -ne $true)){exit}       
        $LogName = $env:Computername + "_" + $XnView_EXE_ProductName + "_" + $XnView_EXE_ProductVersion + ".txt"
        $arguments = " /VERYSILENT /NORESTART /ALLUSERS /NOICONS /FORCECLOSEAPPLICATIONS /LOGCLOSEAPPLICATIONS /LOG=$env:systemdrive\temp\""$LogName"""
        robocopy $XnViews_Path "$env:systemdrive\temp" $XnView_EXE.Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        unblock-file ($env:systemdrive+"\temp\"+$XnView_EXE.Name)
        start-process ($env:systemdrive+"\temp\"+$XnView_EXE.Name) -arg $arguments -WindowStyle Hidden 
        Start-Sleep -s 15
        $Log_Folder_Path = $Log_Path +"\"+ $XnView_EXE_ProductName
        $LogPattern =$env:Computername + "_"+$XnView_EXE_ProductName+"_*.txt"
        if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
        if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}    
    }
}