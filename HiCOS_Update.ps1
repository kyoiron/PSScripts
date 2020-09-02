$RemoveFirstPC=@("TND-STOF-112","TND-5EES-068")
$HiCOSs_Path = "\\172.29.205.114\loginscript\Update\HiCOS"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$HiCOS_EXE = (Get-ChildItem -Path ($HiCOSs_Path+"\*.exe") | Where-Object{$_.VersionInfo.ProductName -eq "HiCOS PKI Smart Card"} | Sort-Object)
$HiCOS_EXE_Path = $HiCOS_EXE.FullName

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
        if($RemoveFirstPC.Contains($env:Computername)){
            foreach($item in $HiCOS_installeds){
                $uninstall_Char = ($item.UninstallString -split "  ")
                $LogFile= "$env:systemdrive\temp\"+$env:Computername + "_HiCOS_Uninstall_"+ $item.DisplayVersion + ".txt"
                if(!(test-Path -Path $LogFile)){
                    $arguments = " /norestart /uninstall /quiet /log " + $LogFile
                    start-process $uninstall_Char[0] -arg $arguments -Wait -WindowStyle Hidden
                    $NeedRestart=$true                   
                }               
            }
            $Log_Folder_Path = $Log_Path +"\"+ $HiCOS_EXE_ProductName
            $LogPattern =$env:Computername + "_HiCOS_*.txt"
            if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
            if(Test-Path -Path $LogFile){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern /XO /NJH /NJS /NDL /NC /NS}
        }
        if($NeedRestart){Restart-Computer -Force}
    }
    $HiCOS_installed = $HiCOS_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    if([version]$HiCOS_installed.BundleVersion -ge [version]$HiCOS_EXE_ProductVersion){exit}
    $LogName = $env:Computername + "_HiCOS_"+ $HiCOS_EXE_ProductVersion + ".txt"
    $arguments = "/quiet /norestart /log $env:systemdrive\temp\$LogName"
    robocopy $HiCOSs_Path "$env:systemdrive\temp" $HiCOS_EXE.Name /XO /NJH /NJS /NDL /NC /NS
    start-process ($env:systemdrive+"\temp\"+$HiCOS_EXE.Name) -arg $arguments -wait -WindowStyle Hidden  
    $Log_Folder_Path = $Log_Path +"\"+ $HiCOS_EXE_ProductName
    $LogPattern =$env:Computername + "_HiCOS_*.txt"
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    if(Test-Path -Path "$env:systemdrive\temp"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogPattern /XO /NJH /NJS /NDL /NC /NS }
}