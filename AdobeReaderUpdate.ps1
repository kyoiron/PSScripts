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
$AdobeReader_EXE_Folder="\\172.29.205.114\loginscript\Update\AdobeReader"
#沒裝的是否要裝$true或$false
$Install_IF_NOT_Installed = $true

$MSP = Get-ChildItem -Path ($AdobeReader_EXE_Folder+"\*.msp") | Sort-Object  -Property Name -Descending | Select-Object -first 1
$AcroRead_MSI = Get-MsiInformation -Path ($AdobeReader_EXE_Folder+"\AcroRead.msi") 

<#
    Get-ChildItem MSP檔結果
    PSPath            : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\AdobeReader\AcroRdrDCUpd1800920044.msp
    PSParentPath      : Microsoft.PowerShell.Core\FileSystem::\\172.29.205.114\loginscript\Update\AdobeReader
    PSChildName       : AcroRdrDCUpd1800920044.msp
    PSProvider        : Microsoft.PowerShell.Core\FileSystem
    PSIsContainer     : False
    Mode              : -a----
    VersionInfo       : File:             \\172.29.205.114\loginscript\Update\AdobeReader\AcroRdrDCUpd1800920044.msp
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
                        
    BaseName          : AcroRdrDCUpd1800920044
    Target            : 
    LinkType          : 
    Name              : AcroRdrDCUpd1800920044.msp
    Length            : 102899712
    DirectoryName     : \\172.29.205.114\loginscript\Update\AdobeReader
    Directory         : \\172.29.205.114\loginscript\Update\AdobeReader
    IsReadOnly        : False
    Exists            : True
    FullName          : \\172.29.205.114\loginscript\Update\AdobeReader\AcroRdrDCUpd1800920044.msp
    Extension         : .msp
    CreationTime      : 2020/9/1 下午 03:41:31
    CreationTimeUtc   : 2020/9/1 上午 07:41:31
    LastAccessTime    : 2020/9/1 下午 03:41:52
    LastAccessTimeUtc : 2020/9/1 上午 07:41:52
    LastWriteTime     : 2017/11/5 上午 06:42:14
    LastWriteTimeUtc  : 2017/11/4 下午 10:42:14
    Attributes        : Archive    
#>

<#
    Get-MsiInformation  AcroRead.msi檔結果
    File            : \\172.29.205.114\loginscript\Update\AdobeReader\AcroRead.msi
    ProductCode     : {AC76BA86-7AD7-1028-7B44-AC0F074E4100}
    Manufacturer    : Adobe Systems Incorporated
    ProductName     : Adobe Acrobat Reader DC - Chinese Traditional
    ProductVersion  : 15.007.20033
    ProductLanguage : 1028
#>    
if($MSP -and $AcroRead_MSI){    
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $AdobeReader_installeds = @()
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $AdobeReader_installeds += Get-ItemProperty $Path | Where-Object{$_.PSChildName -eq $AcroRead_MSI.ProductCode} #| ForEach-Object{ Uninstall-MSI ($($_.UninstallString))}
        }
        
    }
    <#
        AuthorizedCDFPrefix : 
        Comments            :                           			
        Contact             : 客戶支援
        DisplayVersion      : 20.012.20043
        HelpLink            : http://www.adobe.com/tw/support/main.html
        HelpTelephone       : 
        InstallDate         : 20200820
        InstallLocation     : C:\Program Files (x86)\Adobe\Acrobat Reader DC\
        InstallSource       : C:\ProgramData\Adobe\Setup\{AC76BA86-7AD7-1028-7B44-AC0F074E4100}\
        ModifyPath          : MsiExec.exe /I{AC76BA86-7AD7-1028-7B44-AC0F074E4100}
        NoRepair            : 1
        Publisher           : Adobe Systems Incorporated
        Readme              : C:\Program Files (x86)\Adobe\Acrobat Reader DC\ReadmeCT.htm
        Size                : 
        EstimatedSize       : 569508
        UninstallString     : MsiExec.exe /I{AC76BA86-7AD7-1028-7B44-AC0F074E4100}
        URLInfoAbout        : http://www.adobe.com
        URLUpdateInfo       : http://helpx.adobe.com/tw/reader.html
        VersionMajor        : 20
        VersionMinor        : 12
        WindowsInstaller    : 1
        Version             : 336350795
        Language            : 1028
        DisplayName         : Adobe Acrobat Reader DC - Chinese Traditional
        sEstimatedSize2     : 284754
        PSPath              : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{AC76BA86-7AD7-1028-7B44-AC0F074E4100}
        PSParentPath        : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
        PSChildName         : {AC76BA86-7AD7-1028-7B44-AC0F074E4100}
        PSDrive             : HKLM
        PSProvider          : Microsoft.PowerShell.Core\Registry
        
    #>
    $AdobeReader_installed = $AdobeReader_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    if(($AdobeReader_installed -eq $null) -and ($Install_IF_NOT_Installed -ne $true)){exit}
    #$LogName = $env:Computername + "_"+ $AcroRead_MSI.ProductName +"_"+ $MSP_Version + ".txt" 
    #$Log_Folder_Path = $Log_Path +"\"+ $AcroRead_MSI.ProductName   
    
    if($AdobeReader_installed -ne $null){
        $MSP_Version_String = $MSP.BaseName.TrimStart("AcroRdrDCUpd")
        $MSP_Version = $MSP_Version_String.Substring(0,2)+"."+ $MSP_Version_String.Substring(2,3)+"."+$MSP_Version_String.Substring(5,5)         
        if([version]$AdobeReader_installed.DisplayVersion -ge [version]$MSP_Version){exit}
        $LogName = $env:Computername + "_"+ $AcroRead_MSI.ProductName +"_"+ $MSP_Version + ".txt" 
        $Log_Folder_Path = $Log_Path +"\"+ $AcroRead_MSI.ProductName  
        if((Test-Path ($AdobeReader_installed.InstallSource+"AcroRead.msi")) -and ($AdobeReader_installed.VersionMajor -eq ([version]$MSP_Version).Major)){
            $AcroRead_MSI_Path = $AdobeReader_installed.InstallSource+"AcroRead.msi"
        }else{
            robocopy $AdobeReader_EXE_Folder "$env:systemdrive\temp" "AcroRead.msi" "abcpy.ini" "Data1.cab" "setup.exe" "setup.ini"  /XO /NJH /NJS /NDL /NC /NS
            $AcroRead_MSI_Path = "$env:systemdrive\temp\AcroRead.msi"
            $Setup_Content = Get-Content -Path "$env:systemdrive\temp\setup.ini"
            if($Setup_Content -match 'PATCH=AcroRdrDCUpd') {
                ($Setup_Content -replace ($Setup_Content -match 'PATCH=AcroRdrDCUpd') , ("PATCH="+ $MSP.Name) ) | Set-Content  "$env:systemdrive\temp\setup.ini"  
            }
        }
        #Windows10才處理Win7暫時不管
        if([environment]::OSVersion.Version -match '10'){
        #小於現有主版本的安裝程式移除
            if(($AdobeReader_installed.VersionMajor -lt ([version]$MSP_Version).Major) <#-or {$env:Computername -like "TND-1EES-082"}#>){
                $Uninstall_LogName = $env:Computername + "_"+ $AdobeReader_installed.DisplayName +"_"+ $AdobeReader_installed.DisplayVersion+"_Remove" + ".txt"
                $uninstall = $AdobeReader_installed.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
                $uninstall = $uninstall.Trim()
                #Write "Uninstalling..."
                start-process "msiexec.exe" -arg "/X $uninstall /qn /log ""$env:systemdrive\temp\$Uninstall_LogName""" -Wait -WindowStyle Hidden
                if(Test-Path -Path "$env:systemdrive\temp\$Uninstall_LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $Uninstall_LogName /XO /NJH /NJS /NDL /NC /NS}
            }
        }                        
    }else{
        robocopy $AdobeReader_EXE_Folder "$env:systemdrive\temp" "AcroRead.msi" "abcpy.ini" "Data1.cab" "setup.exe" "setup.ini"  /XO /NJH /NJS /NDL /NC /NS
        $AcroRead_MSI_Path = "$env:systemdrive\temp\AcroRead.msi"
        $Setup_Content = Get-Content -Path "$env:systemdrive\temp\setup.ini"
        if($Setup_Content -match 'PATCH=AcroRdrDCUpd') {
            ($Setup_Content -replace ($Setup_Content -match 'PATCH=AcroRdrDCUpd') , ("PATCH="+ $MSP.Name) ) | Set-Content  "$env:systemdrive\temp\setup.ini"  
        }       
    }
    robocopy $AdobeReader_EXE_Folder "$env:systemdrive\temp" $MSP.Name /XO /NJH /NJS /NDL /NC /NS        
    $arguments = "/i $AcroRead_MSI_Path /update "+ $env:systemdrive+"\temp\" + $MSP.Name +" /qn /norestart /log ""$env:systemdrive\temp\$LogName"""    
    unblock-file ($env:systemdrive+"\temp\" + $MSP.Name)
    start-process "msiexec" -arg $arguments -Wait    
    if(!(Test-Path -Path $Log_Folder_Path)){New-Item -ItemType Directory -Path $Log_Folder_Path -Force}
    if(Test-Path -Path "$env:systemdrive\temp\$LogName"){robocopy "$env:systemdrive\temp" $Log_Folder_Path $LogName /XO /NJH /NJS /NDL /NC /NS}
}