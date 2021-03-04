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
#共通參數設定    
    #SQL連線參數
        $MariaDB_Server = "172.29.205.114"
        $MariaDB_DBName = "PC_Data"
        $MariaDB_User = "tnduser"
        $MariaDB_PW = "me@TND1234"
        $ODBC_Connector_Path = "\\172.29.205.114\loginscript\Update\mariadb-connector-odbc"
        $ODBC_Connector_File_Prefx = "mariadb-connector-odbc-3.1.10-"
    #不列出軟體清單（即白名單），備註須與displayname一模一樣之字串。
        #$WhilteList_Software_DisplayName=@("")
    #Email參數
        $From = "${env:computername}<tndi@mail.moj.gov.tw>"
        $To = "kyoiron@mail.moj.gov.tw"
        #寄送副本，指令存參
            #$Cc = "AThirdUser@somewhere.com"
        #夾帶附件，指令存參
            #$Attachment = "C:\users\Username\Documents\SomeTextFile.txt"       
        #SMTP設定
            $SMTPServer = "smtp.moj.gov.tw"
            $SMTPPort = "25"
        #郵件編碼
            $encoding = [System.Text.Encoding]::UTF8
        #郵件主旨
            $Subject = "個人電腦${env:computername}軟體安裝通知   " + (Get-Date -Format "yyy/MM/dd")
    #安裝軟體允許清單位址
        #$SoftwareAllowList_Path="\\172.29.205.114\loginscript\PSScripts\SoftwareAllowList.txt"
        $SoftwareAllowList_Path = "$env:SystemDrive\temp\SoftwareAllowList.txt"

            
#安裝軟體清查
    New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null
    $RegUninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $RegUninstallPaths += Get-ChildItem HKU: | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object {"HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"}
    $Softwares = @()
    foreach($Path in $RegUninstallPaths){        
        $Softwares += (Get-ItemProperty $Path | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate , PSPath )
    }
    #$Softwares |  Export-Csv -Path c:\users\kyoiron\desktop\test.csv -Encoding Unicode -NoTypeInformation
    Remove-PSDrive -Name HKU
    #$Softwares = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate,PSPath )
#電腦基本資訊蒐集
    $Pc_Info = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object Name,Manufacturer,Model,TotalPhysicalMemory
#安裝軟體異動清查
    $WhilteList_Software_DisplayName = Get-Content -Path $SoftwareAllowList_Path
    #$Softwares = $Softwares | Where-Object{$_.DisplayName -notlike "*Update for *" -and  $_.DisplayName -notlike "*MUI (Chinese (Traditional))*"}|Where-Object{$WhilteList_Software_DisplayName -notcontains $_.DisplayName}|Where-Object{![string]::IsNullOrEmpty($_.displayname)}|Select-Object * -Unique | Sort-Object DisplayName 
    $Softwares = $Softwares | Where-Object{$_.DisplayName -notlike "*Update for *" -and  $_.DisplayName -notlike "*MUI (Chinese (Traditional))*"}|Where-Object{![string]::IsNullOrEmpty($_.displayname)}|Select-Object * -Unique | Sort-Object DisplayName 
    #使用SYSTEM身份執行排程的話，$env:TEMP會指$env:systemdrive\Windows\Temp
    $SoftwaresList_Path = "$env:TEMP\${env:COMPUTERNAME}_SoftwaresList.xml"
    if(Test-Path($SoftwaresList_Path)){
        #$oldSoftwares = Import-Csv -Path $SoftwaresList_Path -Delimiter ',' -Encoding UTF8 
        $oldSoftwares = Import-Clixml -Path $SoftwaresList_Path 
    }
    $NewInstallSoftwares = @()
    $NewInstallSoftwares = (Compare-Object -ReferenceObject $oldSoftwares -DifferenceObject $Softwares -Property DisplayName,DisplayVersion,Publisher,InstallDate  | Where-Object{$_.SideIndicator -eq '=>'} | Sort-Object DisplayName | Select-Object DisplayName,DisplayVersion,Publisher,InstallDate )
    #Copy-Item  $SoftwareAllowList_Path -Destination $env:TEMP -Force
    if(Test-Path $SoftwareAllowList_Path){ 
        $SoftwareAllowList = Get-Content -Path $SoftwareAllowList_Path
        $NotAllowSoftware = @()
        $NewInstallSoftwares | ForEach-Object{
            foreach($item in $SoftwareAllowList){
                if(!($_.DisplayName -match $item)){
                    $NotAllowSoftware += $_
                    break
                }
            }
        }
        Remove-Item -Path $SoftwareAllowList_Path -Force
    }
    
#如果電腦有新安裝軟體，則將清單寄送管理者
    if($NotAllowSoftware.Count -gt 0){
        $body = "<table style=""border-collapse: collapse; width: 100%;"" border=""1""><tr><td style=""text-align: center;"">名稱</td><td style=""text-align: center;"">版本</td><td style=""text-align: center;"">安裝日期</td><td style=""text-align: center;"">發行者</td><td style=""text-align: center;width:30%"">需求說明</td></tr>"
        foreach($item in $NotAllowSoftware){            
            $body += "<tr><td>" + $item.DisplayName + "</td><td>" +$item.DisplayVersion + "</td><td>"+ $item.InstallDate + "</td><td>" + $item.Publisher +"</td><td><p>&nbsp;</p></td></tr>"
        }
        $body+='</table><table style="width: 100%; height: 100%;" border="0"><tbody><tr><td style="width: 50%;"><h3>保管人或使用人：</h3></td><td style="width: 50%;height: 5%"><h3>主管：</h3></td></tr><tr><td style="width: 50%;height: 5%"><h3>後會統計室：</h3></td><td style="width: 50%;height: 5%">&nbsp;</td></tr></tbody></table><p>※請擲交統計室備查</p>'
        $body = $body | Out-String
        Send-MailMessage -From $From -to $To -Subject $Subject -Body $body -SmtpServer $SMTPServer  -port $SMTPPort -priority Normal -Encoding $encoding -BodyAsHtml 
    }
    

#$oldSoftwares | Where-Object{ if($Softwares.Contains($_)){$NewInstallSoftwares+=$NewInstallSoftwares}}
#$Softwares  | Where-Object{$oldSoftwares -notcontains $_ }

#$Softwares | Select-Object DisplayName,DisplayVersion,Publisher,InstallDate |  Export-Clixml -Path $SoftwaresList_Path -Force
#if($NewInstallSoftwares.)
#$NewInstallSoftwares=@()
#$NewInstallSoftwares = ($Softwares|Select-Object Displayname,DisplayVersion,Publisher,InstallDate)| Where-Object{ ($oldSoftwares|Select-Object Displayname,DisplayVersion,Publisher,InstallDate).Displayname.trim() -notcontains $_.Displayname.trim()}
    $Softwares |  Export-Clixml -Path $SoftwaresList_Path -Force # Export-Csv -Path $SoftwaresList_Path -Encoding UTF8 -NoTypeInformation -Force
#確認電腦相關基本資料
    $ComputerName = $Pc_Info.Name.trim()
    $OS = (Get-WmiObject Win32_OperatingSystem).Caption.trim()
    $SystemManufacturer = $Pc_Info.Manufacturer.trim()
    $SystemModel = $Pc_Info.Model.trim()
    $Memory = [math]::Round($Pc_Info.TotalPhysicalMemory/1GB)
    $MAC = (get-wmiobject -class "Win32_NetworkAdapterConfiguration" |Where{$_.IpEnabled -Match "True"}|select MACAddress)[0].MACAddress -replace ":","-"
    $IP  = [System.Net.Dns]::GetHostByName($env:COMPUTERNAME).AddressList[0].IpAddressToString
#連接資料庫
    $SqlConnection = New-Object System.Data.ODBC.ODBCConnection
    $SqlConnection.connectionstring = `
    "DRIVER={MariaDB ODBC 3.1 Driver};" +
    "Server = $MariaDB_Server;" +
    "Database = $MariaDB_DBName;" +
    "UID = $MariaDB_User;" +
    "PWD= $MariaDB_PW;" +
    "Option = 3"
    #如果沒有MariaDB ODBC之Driver，則安裝
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
    #顯示連線狀態，Debug用
    $SqlConnection.State
#執行命令
    #指令存參
        #$QuerySelect = 'SELECT `ComputerName`, `OS`, `SystemManufacturer`, `SystemModel`, `Memory`, `IP`, `MAC`, `PropertyNo`, `User`, `Keeper`, `Unit` FROM `PC_Info` WHERE `ComputerName` = ''' + $ComputerName +''''
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
            #確認軟體安裝版本64或32bit
            if($_.PSPath -like "*\Wow6432Node\*"){
                $Architecture=0 
            }else{
                $Architecture=1
            }
            #確認安裝日期為正確的值，如為空字元則補滿0
            if(![string]::IsNullOrEmpty($_.InstallDate)){
                $Date = ([datetime]::parseexact($_.InstallDate,"yyyyMMdd",$null)).toString("yyyy-MM-dd")
            }else{
                    $Date = "0000-00-00"
            }
            #確認安裝軟體發行者的值，並除去異常字元。
            if(![string]::IsNullOrEmpty($_.Publisher)){
                $Publisher = $_.Publisher.Trim([char]0)
            }else{
                $Publisher=""
            }  # '$_.Publisher:'($_.Publisher.Trim([char]0))
            #產生Software_Installed資料表指令
                #指令存參
                    #$Insert_Software_Installed =  'INSERT INTO `PC_Software_Installed`(`ComputerName`, `DisplayName`, `Publisher`, `InstallDate`, `DisplayVersion`, `Architecture`) VALUES ("'+$ComputerName +'","'+ $_.DisplayName + '","' + $Publisher + '","' + $Date + '","' + $_.DisplayVersion + '","'+ $Architecture +'") ON DUPLICATE KEY UPDATE Publisher ='+ """" + $Publisher  + '", InstallDate ="' + $Date + '",DisplayVersion ="' + $_.DisplayVersion + '",Architecture="'+ $Architecture + """"
                    #$Insert_Software_Installed =  'INSERT INTO `PC_Software_Installed`(`ComputerName`, `DisplayName`, `Publisher`, `InstallDate`, `DisplayVersion`, `Architecture`) VALUES ("'+$ComputerName +'","'+ $_.DisplayName + '","' + $Publisher + '","' + $Date + '","' + $_.DisplayVersion + '","'+ $Architecture +'") ON DUPLICATE KEY UPDATE `Publisher`=VALUES(`Publisher`) ,`InstallDate`=VALUES(`InstallDate`),`DisplayVersion`=VALUES(`DisplayVersion`) ,`Architecture`=VALUES(`Architecture`) '
                $Insert_Software_Installed_Table_Multiple += '('''+ $ComputerName +''','''+ $_.DisplayName + ''',''' + $Publisher + ''',''' + $Date + ''',''' + $_.DisplayVersion + ''','''+ $Architecture +'''),'
            #產生Software_Info 資料表指令
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


#指令存參
    #write-host $command.CommandText
    $command.Parameters.Clear()
    #$results_insert = $command.ExecuteReader()
    #$result = $command.ExecuteNonQuery()
#關閉資料庫連線
$SqlConnection.Close()