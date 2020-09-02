function Get-MsiInformation
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                    PositionalBinding=$false,
                    ConfirmImpact='Medium')]
    [Alias("gmsi")]
    Param(
        [parameter(Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    HelpMessage = "Provide the path to an MSI")]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$Path,
        [parameter(Mandatory=$false)]
        [ValidateSet( "ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage" )]
        [string[]]$Property = ( "ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage" )
    )

    Begin
    {
        # Do nothing for prep
    }
    Process
    {
        
        ForEach ( $P in $Path )
        {
            if ($pscmdlet.ShouldProcess($P, "Get MSI Properties"))
            {            
                try
                {
                    Write-Verbose -Message "Resolving file information for $P"
                    $MsiFile = Get-Item -Path $P
                    Write-Verbose -Message "Executing on $P"
                    
                    # Read property from MSI database
                    $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
                    $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiFile.FullName, 0))
                    
                    # Build hashtable for retruned objects properties
                    $PSObjectPropHash = [ordered]@{File = $MsiFile.FullName}
                    ForEach ( $Prop in $Property )
                    {
                        Write-Verbose -Message "Enumerating Property: $Prop"
                        $Query = "SELECT Value FROM Property WHERE Property = '$( $Prop )'"
                        $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
                        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
                        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
                        $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
                        # Return the value to the Property Hash
                        $PSObjectPropHash.Add($Prop, $Value)
                    }
                    
                    # Build the Object to Return
                    $Object = @( New-Object -TypeName PSObject -Property $PSObjectPropHash )
                    
                    # Commit database and close view
                    $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
                    $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
                    $MSIDatabase = $null
                    $View = $null
                }
                catch
                {
                    Write-Error -Message $_.Exception.Message
                }
                finally
                {
                    Write-Output -InputObject @( $Object )
                }
            } # End of ShouldProcess If
        } # End For $P in $Path Loop

    }
    End
    {
        # Run garbage collection and release ComObject
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}

$DocNcomp42s_Path = "\\172.29.205.114\loginscript\Update\DocNcomp42"
$Log_Path = "\\172.29.205.114\Public\sources\audit"

$DocNcomp42 = Get-ChildItem -Path ($DocNcomp42s_Path+"\*.msi")|Sort-Object  -Property PSChildName -Descending | Select-Object -first 1
$DocNcomp42_EXE = Get-MsiInformation -Path  $DocNcomp42.FullName | Where-Object{$_.ProductCode -eq "{8E5272D5-3D03-4C96-B3EB-EF12B67A2F97}"}
if($DocNcomp42_EXE -ne $null){
    
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $DocNcomp42_installeds = Get-ItemProperty $Path | Where-Object{$_.PSChildName -eq $DocNcomp42_EXE.ProductCode} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
        }
    }
    $DocNcomp42_installed = $DocNcomp42_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    if([version]$DocNcomp42_installed.DisplayVersion -ge [version]$DocNcomp42_EXE.ProductVersion){exit}
    $LogName = $env:Computername + "_"+$DocNcomp42_EXE.ProductName +"_"+$DocNcomp42_EXE.ProductVersion + ".txt"
    $arguments = "/quiet /norestart /log $env:systemdrive\temp\" +  $LogName
    robocopy $DocNcomp42s_Path "$env:systemdrive\temp" ""$DocNcomp42.Name" /PURGE /XO /NJH /NJS /NDL /NC /NS".Split(' ')|Out-Null
    start-process ($env:systemdrive+"\temp\"+$DocNcomp42.Name) -arg $arguments -wait
    $Log_Folder_Path = $Log_Path +"\"+ $DocNcomp42_EXE.ProductName
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName /XO /NJH /NJS /NDL /NC /NS }
}

<#
    Get-MsiInformation msi檔所得資料
    File            : C:\temp\docNcomp42@0_13a46-moj.msi
    ProductCode     : {8E5272D5-3D03-4C96-B3EB-EF12B67A2F97}
    Manufacturer    : 傑印資訊
    ProductName     : 文書編輯-公文製作系統元件
    ProductVersion  : 4.2.0013
    ProductLanguage : 1028        
#>

<#
    Get-ChildItem msi檔所得資料
    PSPath            : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\DocNcomp42\docNcomp42@0_13a46-moj.msi
    PSParentPath      : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\DocNcomp42
    PSChildName       : docNcomp42@0_13a46-moj.msi
    PSProvider        : Microsoft.PowerShell.Core\FileSystem
    PSIsContainer     : False
    Mode              : -a----
    VersionInfo       : File:             \\1u,3g472.29.205.114\loginscript\Update\DocNcomp42\docNcomp42@0_13a46-moj.msi
                        InternalName:     
                        OriginalFilename: 
                        FileVersion:      
                        FileDescription:  
                        Product:          
                        ProductVersion:   
                        Debug:            False
                        Patched:          False
                        PreRelease:       False
                        PrivateBuild:     False
                        SpecialBuild:     False
                        Language:         
                        
    BaseName          : docNcomp42@0_13a46-moj
    Target            : 
    LinkType          : 
    Name              : docNcomp42@0_13a46-moj.msi
    Length            : 6029824
    DirectoryName     : \\172.29.205.114\loginscript\Update\DocNcomp42
    Directory         : \\172.29.205.114\loginscript\Update\DocNcomp42
    IsReadOnly        : False
    Exists            : True
    FullName          : \\172.29.205.114\loginscript\Update\DocNcomp42\docNcomp42@0_13a46-moj.msi
    Extension         : .msi
    CreationTime      : 2020/8/31 上午 11:57:23
    CreationTimeUtc   : 2020/8/31 上午 03:57:23
    LastAccessTime    : 2020/8/31 下午 12:01:23
    LastAccessTimeUtc : 2020/8/31 上午 04:01:23
    LastWriteTime     : 2020/2/24 下午 05:27:33
    LastWriteTimeUtc  : 2020/2/24 上午 09:27:33
    Attributes        : Archive    
#>


<#
    Get-ItemProperty $RegUninstallPaths 
    AuthorizedCDFPrefix : 
    Comments            : 公文製作系統元件 [4.2.0_13-moj]
    Contact             : 傑印資訊
    DisplayVersion      : 4.2.0013
    HelpLink            : 
    HelpTelephone       : 
    InstallDate         : 20200827
    InstallLocation     : 
    InstallSource       : C:\temp\
    ModifyPath          : MsiExec.exe /X{8E5272D5-3D03-4C96-B3EB-EF12B67A2F97}
    NoModify            : 1
    Publisher           : 傑印資訊
    Readme              : 
    Size                : 
    EstimatedSize       : 13609
    UninstallString     : MsiExec.exe /X{8E5272D5-3D03-4C96-B3EB-EF12B67A2F97}
    URLInfoAbout        : www.eic.com.tw
    URLUpdateInfo       : 
    VersionMajor        : 4
    VersionMinor        : 2
    WindowsInstaller    : 1
    Version             : 67239949
    Language            : 1028
    DisplayName         : 文書編輯-公文製作系統元件
    sEstimatedSize2     : 9088
    PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{8E5272D5-3D03-4C96-B3EB-EF12B67A2F97}
    PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
    PSChildName         : {8E5272D5-3D03-4C96-B3EB-EF12B67A2F97}
    PSDrive             : HKLM
    PSProvider          : Microsoft.PowerShell.Core\Registry
#>

