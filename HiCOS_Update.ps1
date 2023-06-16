$RemoveFirstPC=@()
$HiCOSs_Path = "\\172.29.205.114\loginscript\Update\HiCOS"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$HiCOS_EXE = (Get-ChildItem -Path ($HiCOSs_Path+"\*.exe") | Where-Object{$_.VersionInfo.ProductName -eq "HiCOS PKI Smart Card"} | Sort-Object)
$HiCOS_EXE_Path = $HiCOS_EXE.FullName
$HiPKILocalSignServer_Displayname = "跨平台網頁元件*"
if($HiCOS_EXE_Path){
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
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $HiCOS_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $HiCOS_EXE_ProductName} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
        }
    }
    if($HiCOS_installeds){
        if($RemoveFirstPC.Contains($env:Computername) -or (($HiCOS_installeds|Measure-Object).count -ge 2)){
            foreach($item in $HiCOS_installeds){
                $uninstall_Char = ($item.UninstallString -split "  ")
                $LogFile= "$env:systemdrive\temp\"+$env:Computername + "_HiCOS_Uninstall_"+ $item.DisplayVersion + ".txt"
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
                    $arguments = " /norestart /uninstall /quiet /log " + $LogFile
                    start-process $uninstall_Char[0] -arg $arguments -Wait -WindowStyle Hidden

                    foreach ($Path in $RegUninstallPaths) {
                        if (Test-Path $Path) {
                            $HiPKILocalSignServer_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -like $HiPKILocalSignServer_Displayname} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
                        }
                    }
                    if($HiPKILocalSignServer_installeds){
                         foreach($item in $HiPKILocalSignServer_installeds){                            
                            $uninstall_Char_HiPK = ($item.UninstallString -split "  ")
                            $LogFile_HiPK = "$env:systemdrive\temp\"+$env:Computername + "_HiPKILocalSignServer_Uninstall_"+ $item.DisplayVersion + ".txt"
                            $arguments_HiPK = " /SILENT /log" + $LogFile_HiPK
                             start-process $uninstall_Char_HiPK[0] -arg $arguments_HiPK -Wait -WindowStyle Hidden
                         }                         
                    }
                    $NeedRestart=$true                   
                }                      
            }
            $Log_Folder_Path = $Log_Path +"\"+ $HiCOS_EXE_ProductName
            $LogPattern =$env:Computername + "_HiCOS_*.txt"
            if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
            if(Test-Path -Path $LogFile){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
        }
        if($NeedRestart){Restart-Computer -Force}
    }
    $HiCOS_installed = $HiCOS_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    if(([version]$HiCOS_installed.BundleVersion -ge [version]$HiCOS_EXE_ProductVersion) -and ((Test-Path "${env:ProgramFiles(x86)}\Chunghwa Telecom\HiCOS PKI Smart Card\TokenUtility.exe") -eq $true)){exit}
    #if(([version]$HiCOS_installed.BundleVersion -ge [version]$HiCOS_EXE_ProductVersion) -and ($env:Computername -ne "TND-ASSE-021")){exit}
    $LogName = $env:Computername + "_HiCOS_"+ $HiCOS_EXE_ProductVersion + ".txt"
    $arguments = " /quiet /norestart /log $env:systemdrive\temp\$LogName"
    robocopy $HiCOSs_Path "$env:systemdrive\temp" $HiCOS_EXE.Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    unblock-file ($env:systemdrive+"\temp\"+$HiCOS_EXE.Name)
    start-process ($env:systemdrive+"\temp\"+$HiCOS_EXE.Name) -arg $arguments  -WindowStyle Hidden
    Start-Sleep -s 15
    $Log_Folder_Path = $Log_Path +"\"+ $HiCOS_EXE_ProductName
    $LogPattern =$env:Computername + "_HiCOS_*.txt"
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
}