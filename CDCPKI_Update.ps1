function Get-FileDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )

    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = Split-Path $FilePath
        $file = Split-Path $FilePath -Leaf
        $shellfolder = $shell.Namespace($folder)
        $shellfile = $shellfolder.ParseName($file)

        $fileDescription = $null
        $fileVersion = $null
        $foundFields = 0

        # 遍歷尋找「檔案描述」和「檔案版本」欄位
        for ($i = 0; $i -le 266 -and $foundFields -lt 2; $i++) {
            $fieldName = $shellfolder.GetDetailsOf($null, $i)
            if ($fieldName -eq "檔案描述") {
                $fileDescription = $shellfolder.GetDetailsOf($shellfile, $i)
                $foundFields++
            }
            elseif ($fieldName -eq "檔案版本") {
                $fileVersion = $shellfolder.GetDetailsOf($shellfile, $i)
                $foundFields++
            }
        }

        return @{
            FileDescription = $fileDescription
            FileVersion = $fileVersion
        }
    }
    catch {
        Write-Error "讀取檔案資訊時發生錯誤: $_"
        return $null
    }
    finally {
        if ($shell) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        }
    }
}
$CDCPKI_Path = "\\172.29.205.114\loginscript\Update\CDCPKI"
#LOG檔NAS存放路徑
#$Log_Path = "\\172.29.205.114\Public\sources\audit"
#沒裝的是否要裝$true或$false
$Install_IF_NOT_Installed = $true

$CDCPKI_EXE = (Get-ChildItem -Path ($CDCPKI_Path+"\*.exe") )   | Sort-Object -Property VersionInfo.FileVersionRaw -Descending | Select-Object -last 1

$CDCPKI_EXE_Path = $CDCPKI_EXE.FullName

if (Test-Path $CDCPKI_EXE_Path -ErrorAction SilentlyContinue) {
    $fileDetails = Get-FileDetails -FilePath $CDCPKI_EXE_Path
    $CDCPKI_EXE_FileDescription = $fileDetails["FileDescription"].TrimEnd(".exe") 
    $CDCPKI_EXE_FileVersion = [version]$fileDetails["FileVersion"]
    $CDCPKI_installeds = $null
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    foreach ($Path in $RegUninstallPaths) {
        if (Test-Path $Path) {
            $CDCPKI_installeds += @(Get-ItemProperty $Path | Where-Object{ $_.DisplayName -match ($CDCPKI_EXE_FileDescription)})
        }
    }
    $installed = $CDCPKI_installeds | Sort-Object -Property DisplayVersion -Descending | Select-Object -first 1
    if(($installed -eq $null) -and ($Install_IF_NOT_Installed -ne $true)){exit}
    if(([version]$CDCPKI_EXE_FileVersion -le [version]$installed.DisplayVersion)){exit}
    $LogName = $env:Computername + "_"+$CDCPKI_EXE_FileDescription +"_"+ $CDCPKI_EXE_FileVersion + ".txt"
    $EXE_FIleName = $CDCPKI_EXE.Name    
    robocopy $FileZillas_Path "$env:systemdrive\temp" $EXE_FIleName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    unblock-file ($env:systemdrive+"\temp\"+$EXE_FIleName)
    & ($env:systemdrive+"\temp\"+$EXE_FIleName) /S
} 