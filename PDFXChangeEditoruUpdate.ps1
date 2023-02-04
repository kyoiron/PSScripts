$RemoveFirstPC=@()
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
$MSI_Folder="\\172.29.205.114\loginscript\Update\PDF-XChangeEditor"
#沒裝的是否要裝$true或$false
$Install_IF_NOT_Installed = $true

#強迫安裝
$foreInstall = $false
#依作業系統版本選擇相對應的安裝檔（32bit或64bit）
if([System.Environment]::Is64BitOperatingSystem){
    $MSI =  Get-MsiInformation -Path (Get-ChildItem -Path ($MSI_Folder+"\*.x64.msi") | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1).fullname
}else{
    $MSI =  Get-MsiInformation -Path (Get-ChildItem -Path ($MSI_Folder+"\*.x86.msi") -Exclude "*.x64.msi" | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1).fullname
}
<#
    File            : \\172.29.205.114\loginscript\Update\PDF-XChangeEditor\EditorV9.x64.msi
    ProductCode     : {C41347AF-84A5-43C9-8212-5A97B2083ACB}
    Manufacturer    : Tracker Software Products (Canada) Ltd.
    ProductName     : PDF-XChange Editor
    ProductVersion  : 9.5.366.0
    ProductLanguage : 1033
#>
if($MSI){
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $installeds = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $installeds += Get-ItemProperty $Path | Where-Object{$_.displayname -match $MSI.ProductName } #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
        }
    }
    <#
        AuthorizedCDFPrefix : 
        Comments            : Tracker Software Products (Canada) Ltd.
        Contact             : http://www.tracker-software.com/contacts/
        DisplayVersion      : 8.0.340.0
        HelpLink            : http://help.tracker-software.com/
        HelpTelephone       : http://www.tracker-software.com/support/
        InstallDate         : 20200729
        InstallLocation     : C:\Program Files\Tracker Software\
        InstallSource       : C:\ProgramData\Tracker Software\TrackerUpdate\Download\
        ModifyPath          : MsiExec.exe /I{4960E5F0-12A4-44FE-8774-157C779C4384}
        Publisher           : Tracker Software Products (Canada) Ltd.
        Readme              : 
        Size                : 
        EstimatedSize       : 692323
        UninstallString     : MsiExec.exe /I{4960E5F0-12A4-44FE-8774-157C779C4384}
        URLInfoAbout        : http://www.tracker-software.com/
        URLUpdateInfo       : http://www.tracker-software.com/
        VersionMajor        : 8
        VersionMinor        : 0
        WindowsInstaller    : 1
        Version             : 134218068
        Language            : 0
        DisplayName         : PDF-XChange Editor
        sEstimatedSize2     : 512144
        PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{4960E5F0-12A4-44FE-8774-157C779C4384}
        PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName         : {4960E5F0-12A4-44FE-8774-157C779C4384}
        PSDrive             : HKLM
        PSProvider          : Microsoft.PowerShell.Core\Registry
    #>  
    if(($installed -eq $null) -and ($Install_IF_NOT_Installed -ne $true)){exit}  
    if(($installeds.Count -gt 1) -or $RemoveFirstPC.Contains($env:Computername)){
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
            if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $Uninstall_LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null} 
            $foreInstall = $true
        }
    }        
    $installed = $installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1    
    if($MSI){
        if($foreInstall -ne $true){
            if(([version]$MSI.ProductVersion -le [version]$installed.DisplayVersion)){exit}
        }
        $LogName = $env:Computername + "_"+$MSI.ProductName +"_"+$MSI.ProductVersion + ".txt"
        $MSI_fileName = (Get-Item $MSI.File).Name
        $arguments = "/i $env:systemdrive\temp\$MSI_fileName ALLUSERS=1 /qn /log ""$env:systemdrive\temp\" +  $LogName+"""" +"ADDLOCAL=""F_Viewer,F_Plugins,F_Plugin_SP,F_FileOpenPlugin,F_ReadOutLoudPlugin,F_OCRPlugin,F_OptimizerPlugin,F_BookmarksPlugin,F_PDFAPlugin,F_OFCPlugin,F_EOCRAPlugin,F_MDPlugin,F_3DPlugin,F_ColorPlugin,F_CSVPlugin,F_BrowserPlugins,F_NPPlugin,F_IEPlugin,F_SanitizePlugin,F_VLangs"" DESKTOP_SHORTCUTS=""0"" NOUPDATER=""1"" APP_LANG=""zh-tw"""
        #$arguments = "/i $env:systemdrive\temp\$MSI_fileName ALLUSERS=1 /qn /log ""$env:systemdrive\temp\" +  $LogName+"""" +"ADDLOCAL=""F_Viewer,F_Plugins,F_Plugin_SP,F_FileOpenPlugin,F_ReadOutLoudPlugin,F_OCRPlugin,F_OptimizerPlugin,F_BookmarksPlugin,F_PDFAPlugin,F_OFCPlugin,F_EOCRAPlugin,F_MDPlugin,F_3DPlugin,F_ColorPlugin,F_CSVPlugin,F_BrowserPlugins,F_NPPlugin,F_IEPlugin,F_SanitizePlugin,F_VLangs,F_ShellExt"" DESKTOP_SHORTCUTS=""0"" NOUPDATER=""1"" APP_LANG=""zh-tw"""
        robocopy $MSI_Folder "$env:systemdrive\temp" (Get-ChildItem -Path $MSI.File).Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        start-process "msiexec" -arg $arguments -Wait
        $Log_Folder_Path = $Log_Path + "\"+ $MSI.ProductName
        if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
        if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
    }

}