# �ҥ��Y��Ҧ��A�H���U�^�����~
Set-StrictMode -Version Latest

# �w�q���|�ܼ�
$systemDrive = $env:SystemDrive
$computerName = $env:ComputerName
$paths = @{
    Local = "$systemDrive\temp"
    Nas = "\\172.29.205.114\mig\Printer"
    NasBak = "\\172.29.205.114\mig\Printer_BACKUP"
}
$fileName = "${computerName}x64.printerExport"
$files = @{
    Local = Join-Path $paths.Local $fileName
    Nas = Join-Path $paths.Nas $fileName
    NasBak = Join-Path $paths.NasBak $fileName
}

# �w�q��x�ɮצW�٩M���|
$logFileName = "${computerName}_PrinterExport.log"
$localLogPath = Join-Path $paths.Local $logFileName
$nasLogPath = Join-Path $paths.Nas $logFileName

# ��ơG�g�J��x
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Write-Host $logMessage
    Add-Content -Path $localLogPath -Value $logMessage
}

# ��ơG�M�z�ª���x�ɮ�
function Clean-OldLogs {
    param([string]$Path)
    if (Test-Path $Path) {
        Remove-Item -Path $Path -Force
        Write-Log "�w�R���ª���x�ɮ�: $Path"
    }
}

# �M�z�ª���x�ɮ�
Clean-OldLogs -Path $localLogPath
Clean-OldLogs -Path $nasLogPath

Write-Log "�}�l����L����ץX�}��"

# �ˬd�@�~�t�Ϊ���
if ([System.Environment]::OSVersion.Version.Major -lt 10) {
    Write-Log "�u�r�I�o�Ӹ}���ݭn Windows 10 �H�W�������~������C"
    exit
}

# �]�m�O�d�Ѽ�
$retentionDays = 30
$cutoffDate = (Get-Date).AddDays(-$retentionDays)

# �ˬd�óB�z�{���� NAS �ɮ�
if (Test-Path $files.Nas) {
    $fileInfo = Get-Item $files.Nas
    if ($fileInfo.CreationTime -gt $cutoffDate) {
        Write-Log "�z�I�o�{�̪񪺶ץX�ɮ׭C�C���ξ�ߡA�ڭ̴N��o���o��"
        exit
    }
    
    try {
        Move-Item -Path $files.Nas -Destination $files.NasBak -Force -ErrorAction Stop
        Remove-Item -Path $files.Nas -Force -ErrorAction Stop
        Write-Log "���\���ʨçR���ª� NAS �ɮסC"
    }
    catch {
        Write-Log "�B�z�ª� NAS �ɮ׮ɵo�Ϳ��~�G$($_.Exception.Message)"
    }
}

# �M�z���ϥΪ��L����s����
Write-Log "�}�l�M�z���ϥΪ��L����s����..."
try {
    $allPorts = Get-PrinterPort -ErrorAction Stop
    $usedPorts = (Get-Printer -ErrorAction Stop).PortName
    $portsToDelete = $allPorts | Where-Object { 
        $_.Name -notin $usedPorts -and 
        $_.CimClass -like "ROOT/StandardCimv2:MSFT_TcpIpPrinterPort"
    }

    foreach ($port in $portsToDelete) {
        if ($null -ne $port.Name) {
            try {
                Remove-PrinterPort -Name $port.Name -ErrorAction Stop
                Write-Log "���\�R�����ϥΪ��s����G$($port.Name)"
            }
            catch {
                Write-Log "�R���s���� $($port.Name) �ɵo�Ϳ��~�G$($_.Exception.Message)"
            }
        }
        else {
            Write-Log "ĵ�i�G�o�{�W�٬� Null ���s����A�w���L�C"
        }
    }
}
catch {
    Write-Log "����L����γs�����T�ɵo�Ϳ��~�G$($_.Exception.Message)"
}

# ����L����ץX
if (Test-Path $files.Local) {
    Remove-Item -Path $files.Local -Force -ErrorAction SilentlyContinue
}

$printbrmPath = "${env:SystemRoot}\system32\Spool\Tools\Printbrm.exe"
Write-Log "�}�l�ץX�L����]�w�o�A�еy���@�U�U��"
try {
    $printbrmOutput = & $printbrmPath -B -F $files.Local 2>&1
    $printbrmExitCode = $LASTEXITCODE

    switch ($printbrmExitCode) {
        0 {
            Write-Log "�L����]�w�ץX���\�����C"
        }
        { $_ -gt 0 } {
            Write-Log "�L����]�w�ץX�����A����ĵ�i�C�h�X�X�G$printbrmExitCode"
            Write-Log "Printbrm ��X�G$printbrmOutput"
        }
        default {
            throw "Printbrm ���楢�ѡC�h�X�X�G$printbrmExitCode"
        }
    }
}
catch {
    Write-Log "�ץX�L����]�w�ɵo�Ϳ��~�G$($_.Exception.Message)"
    if ($printbrmOutput) {
        Write-Log "Printbrm ��X�G$printbrmOutput"
    }
    exit
}

# �ƻs�ɮר� NAS
Write-Log "���b�N�ɮ׽ƻs�� NAS�A���W�N�n�I"
$robocopyArgs = @(
    $paths.Local,
    $paths.Nas,
    $fileName,
    "/PURGE",
    "/XO",
    "/NJH",
    "/NJS",
    "/NDL",
    "/NC",
    "/NS"
)
try {
    $robocopyOutput = & robocopy $robocopyArgs 2>&1
    $robocopyExitCode = $LASTEXITCODE
    if ($robocopyExitCode -ge 8) {
        Write-Log "�ƻs�ɮר� NAS �ɵo�Ϳ��~�CRobocopy �h�X�X�G$robocopyExitCode"
        Write-Log "Robocopy ��X�G$robocopyOutput"
    }
    else {
        Write-Log "�ɮצ��\�ƻs�� NAS�CRobocopy �h�X�X�G$robocopyExitCode"
    }
}
catch {
    Write-Log "���� Robocopy �ɵo�Ϳ��~�G$($_.Exception.Message)"
}

# �ƻs��x�ɮר� NAS
Write-Log "���b�N��x�ɮ׽ƻs�� NAS..."
try {
    Copy-Item -Path $localLogPath -Destination $nasLogPath -Force
    Write-Log "��x�ɮצ��\�ƻs�� NAS�C"
}
catch {
    Write-Log "�ƻs��x�ɮר� NAS �ɵo�Ϳ��~�G$($_.Exception.Message)"
}

Write-Log "�Ӧn�F�I�L����]�w�ץX�y�{�����o�C���W�աI"