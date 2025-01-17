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
$MSI_Folder="\\172.29.205.114\loginscript\Update\PIDDLL\IB_Driver_3.9.3"
#沒裝的是否要裝$true或$false
$Install_IF_NOT_Installed = $true
#強迫安裝
$foreInstall = $false
$MSI =  Get-MsiInformation -Path (Get-ChildItem -Path "$MSI_Folder\IBScanDriver Setup.msi").FullName

$msiexecProcessName = "msiexec.exe"

<#
    File            : \\172.29.205.114\loginscript\Update\PIDDLL\IB_Driver_3.9.3\IBScanDriver Setup.msi
    ProductCode     : {FC93E3AD-91E9-4EBA-90C5-9DE774480AB1}
    Manufacturer    : Integrated Biometrics
    ProductName     : IBScanDriver
    ProductVersion  : 1.0.1
    ProductLanguage : 1033
#>
if($MSI){
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $installeds = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $installeds += Get-ItemProperty $Path | Where-Object{$_.PSChildName -match $MSI.ProductCode.Substring(1,19)} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
        }
    }
    <#
        AuthorizedCDFPrefix : 
        Comments            : 
        Contact             : 
        DisplayVersion      : 1.0.1
        HelpLink            : 
        HelpTelephone       : 
        InstallDate         : 20230918
        InstallLocation     : C:\Users\kyoiron\AppData\Roaming\Integrated Biometrics\IBScanDriver\
        InstallSource       : C:\Users\kyoiron\Desktop\IB_Driver_3.9.3\
        ModifyPath          : MsiExec.exe /I{FC93E3AD-91E9-4EBA-90C5-9DE774480AB1}
        Publisher           : Integrated Biometrics
        Readme              : 
        Size                : 
        EstimatedSize       : 10112
        UninstallString     : MsiExec.exe /I{FC93E3AD-91E9-4EBA-90C5-9DE774480AB1}
        URLInfoAbout        : http://www.IntegratedBiometrics.com
        URLUpdateInfo       : 
        VersionMajor        : 1
        VersionMinor        : 0
        WindowsInstaller    : 1
        Version             : 16777217
        Language            : 1033
        DisplayName         : IBScanDriver
        PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{FC93E3AD-91E9-4EBA-90C5-9DE774480AB1}
        PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName         : {FC93E3AD-91E9-4EBA-90C5-9DE774480AB1}
        PSDrive             : HKLM
        PSProvider          : Microsoft.PowerShell.Core\Registry
    #>  
    if(($installed -eq $null) -and ($Install_IF_NOT_Installed -ne $true)){exit}  
    if($installeds.Count -gt 1){
        $installeds | ForEach-Object{
            #$_.UninstallString.trim()
            #$uninstall = $_.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
            #$uninstall = $uninstall.Trim()                
            #start-process "msiexec.exe" -arg "/X $uninstall /qn /log ""$env:systemdrive\temp\$Uninstall_LogName""" -Wait -WindowStyle Hidden
            $Log_Folder_Path =  $Log_Path +"\"+ $_.DisplayName
            $Uninstall_LogName = $env:Computername + "_"+ $_.DisplayName +"_Remove" + ".txt" 
            if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}             
            if($_.UninstallString -match "msiexec.exe"){
                $uninstall = $_.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
                $uninstall = $uninstall.Trim()                
                start-process "msiexec.exe" -arg "/X $uninstall /qn /log ""$env:systemdrive\temp\$Uninstall_LogName""" -Wait -WindowStyle Hidden                                                                 
            }elseif($_.UninstallString -notmatch "/S" ){
                $uninstall = $_.UninstallString.Trim()             
                start-process -FilePath $uninstall -ArgumentList " /S"  -Wait -WindowStyle Hidden 
                "加入參數/S，嘗試移除：" + $_.DisplayName | Out-File  "$env:systemdrive\temp\$Uninstall_LogName"  
            }
            if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $Uninstall_LogName " /XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null} 
            $foreInstall = $true
        }
    }        
    $installed = $installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1    
    if($MSI){
        if($foreInstall -ne $true){
            if(([version]$MSI.ProductVersion -le [version]$installed.DisplayVersion)){exit}
        }
        $LogName = $env:Computername + "_"+$MSI.ProductName +"_"+$MSI.ProductVersion + ".txt"
        $MSI_fileName =  Split-Path -Path $MSI.File -Leaf
        $MSI_tempFolder = "$env:systemdrive\temp\"+$MSI.ProductName
        $arguments = "/i ""$MSI_tempFolder\$MSI_fileName"" ALLUSERS=1 /qn /log ""$env:systemdrive\temp\" +  $LogName+""""
        if(!(Test-Path -Path $MSI_tempFolder)){New-Item -ItemType Directory -Path MSI_tempFloder -Force}
        robocopy $MSI_Folder $MSI_tempFolder  "/E".Split(' ') | Out-Null        
        start-process "msiexec" -arg $arguments -Wait
        $Log_Folder_Path = $Log_Path + "\"+ $MSI.ProductName
        if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
        if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
        if(Test-Path -Path $MSI_tempFolder){Remove-Item  -Path $MSI_tempFolder -Recurse -Force}
    }
}
   