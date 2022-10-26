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
$MSI_Folder="\\172.29.205.114\loginscript\Update\MODA_ODF"
#沒裝的是否要裝$true或$false
$Install_IF_NOT_Installed = $true
#依作業系統版本選擇相對應的安裝檔（32bit或64bit）
if([System.Environment]::Is64BitOperatingSystem){
    $MSI =  Get-MsiInformation -Path (Get-ChildItem -Path "\\172.29.205.114\loginscript\Update\MODA_ODF" | Where-Object{ $_.Name -like "*x86_64*"} | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1).fullname
}else{
    $MSI =  Get-MsiInformation -Path (Get-ChildItem -Path "\\172.29.205.114\loginscript\Update\MODA_ODF" | Where-Object{ $_.Name -notlike "*x86_64*"} | Sort-Object -Property VersionInfo -Descending | Select-Object -first 1).fullname
}
<#
    File            : \\172.29.205.114\loginscript\Update\MODA_ODF\MODA ODF Application Tools_x86_64_3.5.2.msi
    ProductCode     : {90D1075E-2DFB-469E-BFBA-DAD7C0E5E57B}
    Manufacturer    : Ministry of Digital Affairs
    ProductName     : MODA ODF Application Tools 3.5.2
    ProductVersion  : 3.5.2
    ProductLanguage : 1033

    File            : \\172.29.205.114\loginscript\Update\NDC_ODF\NDC ODF Application Tools_x86_64-3.3.3.msi
    ProductCode     : {543327F0-F262-4490-B6B5-7BE571F24450}
    Manufacturer    : National Development Council
    ProductName     : NDC ODF Application Tools 3.3.3
    ProductVersion  : 3.3.3
    ProductLanguage : 1033
#>

if($MSI){
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $installeds = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $installeds += Get-ItemProperty $Path | Where-Object{($_.PSChildName -match $MSI.ProductCode)-or($_.PSChildName -match "{543327F0-F262-4490-B6B5-7BE571F24450}")} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
        }
    }
    <#
        AuthorizedCDFPrefix : 
        Comments            : NDC ODF Application Tools  (en-US,zh-TW)
        Contact             : LibreOffice 社群
        DisplayVersion      : 3.3.3
        HelpLink            : https://zh-tw.libreoffice.org/get-help
        HelpTelephone       : 
        InstallDate         : 20220221
        InstallLocation     : C:\Program Files\NDC ODF Application Tools 6\
        InstallSource       : C:\Users\kyoiron\Downloads\
        ModifyPath          : MsiExec.exe /I{543327F0-F262-4490-B6B5-7BE571F24450}
        Publisher           : National Development Council
        Readme              : 
        Size                : 
        EstimatedSize       : 1873088
        UninstallString     : MsiExec.exe /I{543327F0-F262-4490-B6B5-7BE571F24450}
        URLInfoAbout        : https://zh-tw.libreoffice.org/
        URLUpdateInfo       : https://zh-tw.libreoffice.org/download
        VersionMajor        : 3
        VersionMinor        : 3
        WindowsInstaller    : 1
        Version             : 50528259
        Language            : 1028
        DisplayName         : NDC ODF Application Tools 3.3.3
        sEstimatedSize2     : 936544
        PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{543327F0-F262-4490-B6B5-7BE571F24450}
        PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName         : {543327F0-F262-4490-B6B5-7BE571F24450}
        PSDrive             : HKLM
        PSProvider          : Microsoft.PowerShell.Core\Registry        
    #>
    $installed = $installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    if(($installed -eq $null) -and ($Install_IF_NOT_Installed -ne $true)){exit}
    if($MSI){
        if(([version]$MSI.ProductVersion -le [version]$installed.DisplayVersion)){exit}
        $LogName = $env:Computername + "_"+$MSI.ProductName +"_"+$MSI.ProductVersion + ".txt"
        $MSI_fileName = (Get-Item $MSI.File).Name
        $arguments = "/i ""$env:systemdrive\temp\$MSI_fileName"" ALLUSERS=1 /qn /norestart /log ""$env:systemdrive\temp\" +  $LogName+""""
        robocopy $MSI_Folder "$env:systemdrive\temp" (Get-ChildItem -Path $MSI.File).Name "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
        Stop-Process -Name FontClient -ErrorAction SilentlyContinue -Force
        start-process "msiexec" -arg $arguments -Wait
        $Log_Folder_Path = $Log_Path + "\"+ $MSI.ProductName
        if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
        if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null}
    }
    if((Get-Process -Name FontClient -ErrorAction SilentlyContinue)-eq $null){Start-Process "$env:systemdrive\CMEX_FontClient\FontClient.exe" -ArgumentList "-gui"}
}
