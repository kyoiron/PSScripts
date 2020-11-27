function Execute-MySQLNonQuery($conn, [string]$query) { 
    $command = $conn.CreateCommand()                  # Create command object
    $command.CommandText = $query                     # Load query into object
    $RowsInserted = $command.ExecuteNonQuery()        # Execute command 
    $command.Dispose()                                # Dispose of command object
    if ($RowsInserted) { 
        return $RowInserted 
    } else { 
        return $false 
  } 
}
#不列出軟體，須名稱一模一樣
$WhilteList_Software_DisplayName=@("")

New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null
$RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
$RegUninstallPaths += Get-ChildItem HKU: | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object {"HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"}
$Softwares = @()
foreach($Path in $RegUninstallPaths){        
    $Softwares += (Get-ItemProperty $Path | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate,PSPath )
}
#$Softwares |  Export-Csv -Path c:\users\kyoiron\desktop\test.csv -Encoding Unicode -NoTypeInformation
Remove-PSDrive -Name HKU
#$Softwares = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate,PSPath )
$Pc_Info = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object Name,Manufacturer,Model,TotalPhysicalMemory
$Softwares = $Softwares | Where-Object{$_.DisplayName -notlike "*Update for *" -and  $_.DisplayName -notlike "*MUI (Chinese (Traditional))*"}|Where-Object{$WhilteList_Software_DisplayName -notcontains $_.DisplayName}|Select-Object * -Unique

$ComputerName = $Pc_Info.Name.trim()
$OS = (Get-WmiObject Win32_OperatingSystem).Caption.trim()
$SystemManufacturer = $Pc_Info.Manufacturer.trim()
$SystemModel = $Pc_Info.Model.trim()
$Memory = [math]::Round($Pc_Info.TotalPhysicalMemory/1GB)
$MAC = (get-wmiobject -class "Win32_NetworkAdapterConfiguration" | Where-Object{$_.IpEnabled -Match "True"}|select MACAddress)[0].MACAddress
$IP = [System.Net.Dns]::GetHostByName($env:COMPUTERNAME).AddressList[0].IpAddressToString

$SQLServer = "172.29.205.114"
$SQLDBName = "PC_Data"
$SQLUser = "tnduser"
$SQLPW = "me@TND1234"
$ODBC_Connector_Path = "\\172.29.205.114\loginscript\Update\mariadb-connector-odbc"
$ODBC_Connector_File_Prefx = "mariadb-connector-odbc-3.1.10-"

#連接資料庫
$SqlConnection = New-Object System.Data.ODBC.ODBCConnection
$SqlConnection.connectionstring = `
   "DRIVER={MariaDB ODBC 3.1 Driver};" +
   "Server = $SQLServer;" +
   "Database = $SQLDBName;" +
   "UID = $SQLUser;" +
   "PWD= $SQLPW;" +
   "Option = 3"
if([string]::IsNullOrEmpty(($Softwares.displayname -like "MariaDB ODBC Driver*"))){
    if([System.Environment]::Is64BitOperatingSystem){
        $ODBC_Connector_File = $ODBC_Connector_File_Prefx + "win64.msi"
    }else{
        $ODBC_Connector_File = $ODBC_Connector_File_Prefx + "win32.msi"
    }        
    robocopy $ODBC_Connector_Path "$env:systemdrive\temp" $ODBC_Connector_File "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    $arguments = "/i $env:systemdrive\temp\$ODBC_Connector_File /qn /log ""$env:systemdrive\temp\" +  ($ODBC_Connector_File_Prefx.TrimEnd("-")+"_log.txt")+""""
    $msiexec = "msiexec" 
    start-process $msiexec $arguments -Wait -WindowStyle Hidden
}
$SqlConnection.Open()
#顯示連線狀態
$SqlConnection.State
#執行命令

$QuerySelect = 'SELECT `ComputerName`, `OS`, `SystemManufacturer`, `SystemModel`, `Memory`, `IP`, `MAC`, `PropertyNo`, `User`, `Keeper`, `Unit` FROM `PC_Info` WHERE `ComputerName` = ''' + $ComputerName +''''
#$QuerySelect = 'SELECT * FROM `PC_Info`'
$Insert_PC_Info = 'INSERT INTO `PC_Info`(`ComputerName`, `OS`, `SystemManufacturer`, `SystemModel`, `Memory`, `IP`, `MAC`, `PropertyNo`, `User`, `Keeper`, `Unit`) VALUES ("'+$ComputerName+'","'+ $OS+'","'+$SystemManufacturer+'","'+$SystemModel+'","'+$Memory+'","'+$IP+'","'+$MAC+'","","","","") ' + " ON DUPLICATE KEY UPDATE `OS`=VALUES(`OS`) ,`SystemManufacturer`=VALUES(`SystemManufacturer`) ,`SystemModel`=VALUES(`SystemModel`),`Memory`=VALUES(`Memory`),`IP`=VALUES(`IP`),`MAC`=VALUES(`MAC`)"
$DataTable = New-Object -TypeName System.Data.DataTable
$command = $SqlConnection.CreateCommand()
$command.CommandText = $Insert_PC_Info
$results = $command.ExecuteReader()
#$DataTable.Load($results)
$results.Close()
$Insert_Software_Installed_Table_Multiple='INSERT INTO `PC_Software_Installed`(`ComputerName`, `DisplayName`, `Publisher`, `InstallDate`, `DisplayVersion`, `Architecture`) VALUES '
$Insert_Software_Info_Table_Multiple ='INSERT INTO `Software_Info`(`DisplayName`, `Publisher`, `DisplayVersion`) VALUES '
$Softwares| ForEach-Object{
    if(![string]::IsNullOrEmpty($_.Displayname)){
        if($_.PSPath -like "*\Wow6432Node\*"){ $Architecture=0 }else{ $Architecture=1 }
        if(![string]::IsNullOrEmpty($_.InstallDate)){
            $Date = ([datetime]::parseexact($_.InstallDate,"yyyyMMdd",$null)).toString("yyyy-MM-dd")
        }else{
            $Date = "0000-00-00"
        }
        if(![string]::IsNullOrEmpty($_.Publisher)){$Publisher = $_.Publisher.Trim([char]0)}else{$Publisher=""}  # '$_.Publisher:'($_.Publisher.Trim([char]0))
        #$Insert_Software_Installed =  'INSERT INTO `PC_Software_Installed`(`ComputerName`, `DisplayName`, `Publisher`, `InstallDate`, `DisplayVersion`, `Architecture`) VALUES ("'+$ComputerName +'","'+ $_.DisplayName + '","' + $Publisher + '","' + $Date + '","' + $_.DisplayVersion + '","'+ $Architecture +'") ON DUPLICATE KEY UPDATE Publisher ='+ """" + $Publisher  + '", InstallDate ="' + $Date + '",DisplayVersion ="' + $_.DisplayVersion + '",Architecture="'+ $Architecture + """"
        #$Insert_Software_Installed =  'INSERT INTO `PC_Software_Installed`(`ComputerName`, `DisplayName`, `Publisher`, `InstallDate`, `DisplayVersion`, `Architecture`) VALUES ("'+$ComputerName +'","'+ $_.DisplayName + '","' + $Publisher + '","' + $Date + '","' + $_.DisplayVersion + '","'+ $Architecture +'") ON DUPLICATE KEY UPDATE `Publisher`=VALUES(`Publisher`) ,`InstallDate`=VALUES(`InstallDate`),`DisplayVersion`=VALUES(`DisplayVersion`) ,`Architecture`=VALUES(`Architecture`) '
        $Insert_Software_Installed_Table_Multiple += '('''+ $ComputerName +''','''+ $_.DisplayName + ''',''' + $Publisher + ''',''' + $Date + ''',''' + $_.DisplayVersion + ''','''+ $Architecture +'''),'
        $Insert_Software_Info_Table_Multiple +='('''+ $_.DisplayName + ''',''' + $Publisher + ''',''' + $_.DisplayVersion + '''),'
        #$command.CommandText = $Insert_Software_Installed        
    }
}

$command.CommandText = $Insert_Software_Installed_Table_Multiple.TrimEnd(',') + " ON DUPLICATE KEY UPDATE `Publisher`=VALUES(`Publisher`) ,`InstallDate`=VALUES(`InstallDate`),`DisplayVersion`=VALUES(`DisplayVersion`) ,`Architecture`=VALUES(`Architecture`) "
#$command.CommandText.ToString() | Export-Csv -Path c:\users\kyoiron\desktop\test.csv -Encoding Unicode -NoTypeInformation
$results_insert = $command.ExecuteReader()
$results_insert.Close()


$command.CommandText = $Insert_Software_Info_Table_Multiple.TrimEnd(',') + ' ON DUPLICATE KEY UPDATE `Publisher`=VALUES(`Publisher`) ,`DisplayVersion`=VALUES(`DisplayVersion`) '
$results_insert = $command.ExecuteReader()
$results_insert.Close()



#write-host $command.CommandText
$command.Parameters.Clear()
#$results_insert = $command.ExecuteReader()
#$result = $command.ExecuteNonQuery()
$SqlConnection.Close()