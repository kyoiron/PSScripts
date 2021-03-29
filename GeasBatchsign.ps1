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
$GeasBatchsigns_Path = "\\172.29.205.114\loginscript\Update\GeasBatchsign"
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$GeasBatchsign_MSIs = Get-ChildItem -Path ($GeasBatchsigns_Path+"\*.msi") 
$GeasBatchsign_Newest_MSI = $GeasBatchsign_MSIs | ForEach-Object { Get-MsiInformation -Path ($_.FullName ) } | Sort-Object -Descending ProductVersion | Select-Object -first 1
if($GeasBatchsign_Newest_MSI){
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $GeasBatchsign_installeds = Get-ItemProperty $Path | Where-Object{$_.DisplayName -eq $GeasBatchsign_Newest_MSI.ProductName}
        }
    }
    <#
        AuthorizedCDFPrefix : 
        Comments            : 
        Contact             : 
        DisplayVersion      : 3.20.0000
        HelpLink            : 
        HelpTelephone       : 
        InstallDate         : 20200413
        InstallLocation     : C:\Program Files (x86)\WellChoose\BatchSignCS\
        InstallSource       : C:\Users\tndadmin\AppData\Local\Microsoft\Windows\INetCache\IE\PNW1K9GW\
        ModifyPath          : MsiExec.exe /X{862BF68D-0000-4DA4-B8FB-94E1CC7D0446}
        NoModify            : 1
        NoRepair            : 1
        Publisher           : Wellchoose Inc.
        Readme              : 
        Size                : 
        EstimatedSize       : 24147
        UninstallString     : MsiExec.exe /X{862BF68D-0000-4DA4-B8FB-94E1CC7D0446}
        URLInfoAbout        : http://www.WellchooseInc..com
        URLUpdateInfo       : 
        VersionMajor        : 3
        VersionMinor        : 20
        WindowsInstaller    : 1
        Version             : 51642368
        Language            : 1028
        DisplayName         : 背景簽章服務CS
        sEstimatedSize2     : 17737
        PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{862BF68D-0000-4DA4-B8FB-94E1CC7D0446}
        PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName         : {862BF68D-0000-4DA4-B8FB-94E1CC7D0446}
        PSDrive             : HKLM
        PSProvider          : Microsoft.PowerShell.Core\Registry
    #>
    $GeasBatchsign_installed = $GeasBatchsign_installeds| Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    #if([version]$GeasBatchsign_installed.DisplayVersion -ge [version]$GeasBatchsign_Newest_MSI.ProductVersion){exit}
    #有安裝，就直接離開。因為有裝2版仍可正常運作，如果裝3版則無法判斷次版本新舊，所以沒裝的再裝即可。
    if($null -ne $GeasBatchsign_installed){exit}
    $LogName = $env:Computername + "_"+$GeasBatchsign_Newest_MSI.ProductName +"_"+$GeasBatchsign_Newest_MSI.ProductVersion + ".txt"
    $GeasBatchsign_MIS_fileName = (Get-Item $GeasBatchsign_Newest_MSI.File).Name 
    $arguments = "/i $env:systemdrive\temp\$GeasBatchsign_MIS_fileName  /qn /log ""$env:systemdrive\temp\" +  $LogName+""""
    robocopy $GeasBatchsigns_Path "$env:systemdrive\temp" (Get-ChildItem -Path $GeasBatchsign_Newest_MSI.File).Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    start-process "msiexec" -arg $arguments -Wait
    $Log_Folder_Path = $Log_Path + "\"+ $GeasBatchsign_Newest_MSI.ProductName
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
}
