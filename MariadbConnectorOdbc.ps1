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
#LOG檔NAS存放路徑
$Log_Path = "\\172.29.205.114\Public\sources\audit"
#MSI檔安裝路徑
$MariadbConnectorOdbc_EXE_Folder="\\172.29.205.114\loginscript\Update\mariadb-connector-odbc"
#沒裝的是否要裝$true或$false
$Install_IF_NOT_Installed = $true

if([System.Environment]::Is64BitOperatingSystem){
    $MSI =  Get-MsiInformation -Path (Get-ChildItem -Path ($MariadbConnectorOdbc_EXE_Folder+"\*-win64.msi") | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1).fullname
}else{
    $MSI =  Get-MsiInformation -Path (Get-ChildItem -Path ($MariadbConnectorOdbc_EXE_Folder+"\*-win32.msi") | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1).fullname
}
<#
    Get-MsiInformation -Path ..\mariadb-connector-odbc-3.1.17-win64.msi
    
    File            : \\172.29.205.114\loginscript\Update\mariadb-connector-odbc\mariadb-connector-odbc-3.1.17-win64.msi
    ProductCode     : {2129726C-E25C-4197-8E6B-2E06175E8D0C}
    Manufacturer    : MariaDB
    ProductName     : MariaDB ODBC Driver 64-bit
    ProductVersion  : 3.1.17
    ProductLanguage : 1033
#>

if($MSI){
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $installeds = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $installeds += Get-ItemProperty $Path | Where-Object{$_.PSChildName -eq $MSI.ProductCode} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
        }        
    }
    <#
        AuthorizedCDFPrefix : 
        Comments            : 
        Contact             : 
        DisplayVersion      : 3.1.17
        HelpLink            : 
        HelpTelephone       : 
        InstallDate         : 20220927
        InstallLocation     : 
        InstallSource       : C:\Users\kyoiron\Downloads\
        ModifyPath          : MsiExec.exe /I{2129726C-E25C-4197-8E6B-2E06175E8D0C}
        Publisher           : MariaDB
        Readme              : 
        Size                : 
        EstimatedSize       : 53325
        UninstallString     : MsiExec.exe /I{2129726C-E25C-4197-8E6B-2E06175E8D0C}
        URLInfoAbout        : 
        URLUpdateInfo       : 
        VersionMajor        : 3
        VersionMinor        : 1
        WindowsInstaller    : 1
        Version             : 50397201
        Language            : 1033
        DisplayName         : MariaDB ODBC Driver 64-bit
        PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{2129726C-E25C-4197-8E6B-2E06175E8D0C}
        PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName         : {2129726C-E25C-4197-8E6B-2E06175E8D0C}
        PSDrive             : HKLM
        PSProvider          : Microsoft.PowerShell.Core\Registry
    #>
    $installed = $installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    <#以下狀況結束程式
        1.未安裝且未設定「沒裝要裝」
        2.有安裝但安裝檔之版本小於或等於已安裝的版本
    #>
    if(($installed -eq $null) -and ($Install_IF_NOT_Installed -ne $true)){exit}
    if($MSI){
        if(([version]$MSI.ProductVersion -le [version]$installed.DisplayVersion)){exit}
        $LogName = $env:Computername + "_"+$MSI.ProductName +"_"+$MSI.ProductVersion + ".txt"
        $MSI_fileName =  Split-Path -Path  $MSI.File -Leaf
        $arguments = "/i $env:systemdrive\temp\$MSI_fileName ALLUSERS=1 /quiet /l*vx ""$env:systemdrive\temp\" +  $LogName+""""
        robocopy $MariadbConnectorOdbc_EXE_Folder "$env:systemdrive\temp" (Get-ChildItem -Path $MSI.File).Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        do {
                $msiexecProcess = Get-Process -Name "msiexec.exe" -ErrorAction SilentlyContinue
                if ($msiexecProcess -ne $null) {
                    Start-Sleep -Seconds 1
                }
        } while ($msiexecProcess -ne $null)
        start-process "msiexec" -arg $arguments -Wait
        $Log_Folder_Path = $Log_Path + "\"+ $MSI.ProductName
        if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
        if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
    }
}