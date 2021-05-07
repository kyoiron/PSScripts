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
$Eic_Path = "\\172.29.205.114\loginscript\Update\eic"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$EicPrint_MSIs = Get-ChildItem -Path ($Eic_Path+"\EicPrint*.msi") 
$EICSignTSR_MSIs = Get-ChildItem -Path ($Eic_Path+"\EICSignTSR*.msi") 
$EicPrint_Newest_MSI = $EicPrint_MSIs | ForEach-Object { Get-MsiInformation -Path ($_.FullName ) } | Sort-Object -Descending ProductVersion | Select-Object -first 1
$EICSignTSR_Newest_MSI = $EICSignTSR_MSIs | ForEach-Object { Get-MsiInformation -Path ($_.FullName ) } | Sort-Object -Descending ProductVersion | Select-Object -first 1
if($EicPrint_Newest_MSI){
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    foreach ($Path in $RegUninstallPaths) {
        $EicPrint_installeds=@()
        if (Test-Path $Path) {
            $EicPrint_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -eq $EicPrint_Newest_MSI.ProductName}
        }
    }
    <#
        AuthorizedCDFPrefix : 
        Comments            : 法務部_筆硯列印工具[2021.4.20.1]
        Contact             : 傑印資訊
        DisplayVersion      : 1.0.0006
        HelpLink            : 
        HelpTelephone       : 
        InstallDate         : 20210506
        InstallLocation     : 
        InstallSource       : C:\temp\
        ModifyPath          : MsiExec.exe /I{E35F73D3-EB77-4AE3-A520-6197613AD93B}
        Publisher           : 傑印資訊
        Readme              : 
        Size                : 
        EstimatedSize       : 9360
        UninstallString     : MsiExec.exe /I{E35F73D3-EB77-4AE3-A520-6197613AD93B}
        URLInfoAbout        : 
        URLUpdateInfo       : 
        VersionMajor        : 1
        VersionMinor        : 0
        WindowsInstaller    : 1
        Version             : 16777222
        Language            : 1028
        DisplayName         : 筆硯列印工具
        PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{E35F73D3-EB77-4AE3-A520-6197613AD93B}
        PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName         : {E35F73D3-EB77-4AE3-A520-6197613AD93B}
        PSDrive             : HKLM
        PSProvider          : Microsoft.PowerShell.Core\Registry    
    #>

    $EicPrint_installed = $EicPrint_installeds| Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    #if([version]$EicPrint_installed.DisplayVersion -ge [version]$EicPrint_Newest_MSI.ProductVersion){exit}
    
    if($null -eq $EicPrint_installed){
        $LogName = $env:Computername + "_"+$EicPrint_Newest_MSI.ProductName +"_"+$EicPrint_Newest_MSI.ProductVersion + ".txt"
        $EicPrint_MIS_fileName = (Get-Item $EicPrint_Newest_MSI.File).Name 
        $arguments = "/i $env:systemdrive\temp\$EicPrint_MIS_fileName ALLUSERS=1 /qn /log ""$env:systemdrive\temp\" +  $LogName+""""
        robocopy $Eic_Path "$env:systemdrive\temp" (Get-ChildItem -Path $EicPrint_Newest_MSI.File).Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        start-process "msiexec" -arg $arguments -Wait
        $Log_Folder_Path = $Log_Path + "\"+ $EicPrint_Newest_MSI.ProductName
        if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
        if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
    }
}
if($EICSignTSR_Newest_MSI){
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $EICSignTSR_installeds = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {            
            $EICSignTSR_installeds += Get-ItemProperty $Path | Where-Object{$_.DisplayName -eq $EICSignTSR_Newest_MSI.ProductName}            
        }
    }
    <#
        AuthorizedCDFPrefix : 
        Comments            : 法務部_筆硯簽章工具[2021.4.15.1]
        Contact             : 傑印資訊
        DisplayVersion      : 1.0.0004
        HelpLink            : 
        HelpTelephone       : 
        InstallDate         : 20210506
        InstallLocation     : 
        InstallSource       : C:\temp\
        ModifyPath          : MsiExec.exe /I{9619B081-52EC-457C-9223-93CFC446D88E}
        Publisher           : 傑印資訊
        Readme              : 
        Size                : 
        EstimatedSize       : 7888
        UninstallString     : MsiExec.exe /I{9619B081-52EC-457C-9223-93CFC446D88E}
        URLInfoAbout        : 
        URLUpdateInfo       : 
        VersionMajor        : 1
        VersionMinor        : 0
        WindowsInstaller    : 1
        Version             : 16777220
        Language            : 1028
        DisplayName         : 筆硯簽章工具
        PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{9619B081-52EC-457C-9223-93CFC446D88E}
        PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName         : {9619B081-52EC-457C-9223-93CFC446D88E}
        PSDrive             : HKLM
        PSProvider          : Microsoft.PowerShell.Core\Registry
    #>
    $EICSignTSR_installed = $EICSignTSR_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    #if([version]$EICSignTSR_installed.DisplayVersion -ge [version]$EICSignTSR_Newest_MSI.ProductVersion){exit}    
    if($null -ne $EICSignTSR_installed){exit}
    $LogName = $env:Computername + "_"+$EICSignTSR_Newest_MSI.ProductName +"_"+$EICSignTSR_Newest_MSI.ProductVersion + ".txt"
    $EICSignTSR_MIS_fileName = (Get-Item $EICSignTSR_Newest_MSI.File).Name 
    $arguments = "/i $env:systemdrive\temp\$EICSignTSR_MIS_fileName ALLUSERS=1 /qn /log ""$env:systemdrive\temp\" +  $LogName+""""
    robocopy $Eic_Path "$env:systemdrive\temp" (Get-ChildItem -Path $EICSignTSR_Newest_MSI.File).Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    start-process "msiexec" -arg $arguments -Wait
    $Log_Folder_Path = $Log_Path + "\"+ $EICSignTSR_Newest_MSI.ProductName
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
}