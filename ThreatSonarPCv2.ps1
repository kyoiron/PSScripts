#共用參數設定
    #機關名稱，須自行修改
        $Name="台南看守所"    
    #LOG檔存放路徑，須自行修改
        $Log_Folder_Path = "\\172.29.205.114\Public\sources\audit\ThreatSonarPC"
    #預設執行時間，可自訂，11:00為預設，請用24H制
    $Specific_Time = "11:00"
    #ThreatSona執行檔檔案存放路徑
        $ThreatSonar_Path = "$env:SystemDrive\ThreatSonar"
    #解壓縮執行程式
        if(test-path "$env:ProgramFiles\7-Zip\7z.exe"){
           $Unzip_EXE = "$env:ProgramFiles\7-Zip\7z.exe"  
        }else{
            if(test-path "$env:ProgramFiles(x86)\7-Zip\7z.exe"){
                $Unzip_EXE = "$env:ProgramFiles(x86)\7-Zip\7z.exe"  
            }else{                
                $Unzip_EXE = $null
            }            
        }
#檢查是否有安裝舊版
    if(test-Path "$ThreatSonar_Path\TS_scan.bat"){
        #1.刪除舊版程式    
            #Stop service 終止 Process
                Taskkill /F /PID ThreatSonar.exe
                Taskkill /F /IM ThreatSonar.exe        
            #Remove Service
                sc delete ThreatSonar        
            #Remove Files and Folders in endpoint
                Get-ChildItem -Path "$env:ProgramData\Task" -Force -Recurse -ErrorAction SilentlyContinue | Remove-Item        
                Get-ChildItem -Path "$env:ProgramData\sonar*" -Include sonar*.*  -Force -Recurse -ErrorAction SilentlyContinue | Remove-Item
                Get-ChildItem -Path "$env:ProgramFiles(x86)\ThreatSonar" -Force -Recurse  -ErrorAction SilentlyContinue| Remove-Item
        #2.刪除排程
            schtasks /Delete /TN "ThreatSonar" /F
            #Unregister-ScheduledTask -TaskName "ThreatSonar" -Confirm:$false -WhatIf
    }
#下載程式
    #機關名稱，請參考網址：http://download.moj/files/工具檔案/T5/所屬PC/    
    #下載檔名
    $ThreatSonar_zip_FileName =  "ThreatSonar_$Name-PC.zip"  
    $url_exe = "http://download.moj/files/工具檔案/T5/所屬PC/$ThreatSonar_zip_FileName"
    #檔案下載
    Start-Job -Name WebReq -ScriptBlock { param($p1, $p2)
        Invoke-WebRequest -Uri $p1 -OutFile "$env:systemdrive\temp\$p2"
    } -ArgumentList $url_exe,$ThreatSonar_zip_FileName

    Wait-Job -Name WebReq -Force
    Remove-Job -Name WebReq -Force

    #如果下載成功，與安裝程式進行判斷，如有變動則使用下載版的程式
    $TEMP_Folder = "$env:systemdrive\temp"
    if(Test-Path "$TEMP_Folder\$ThreatSonar_zip_FileName"){                  
        #取得下載壓縮檔內的檔案清單
            [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
            $FilesInDownloadZip = ([IO.Compression.ZipFile]::OpenRead("$TEMP_Folder\$ThreatSonar_zip_FileName").Entries).fullname            
        #解壓縮下載檔案至暫存資料夾
            if($Unzip_EXE){
                #使用7zip解壓縮
                &$Unzip_EXE x "$TEMP_Folder\$ThreatSonar_zip_FileName" "-o$TEMP_Folder" -y
            }else{
                #使用Powershell之解壓縮命令
                Expand-Archive -Path "$TEMP_Folder\$ThreatSonar_zip_FileName" -DestinationPath $TEMP_Folder -Force
            }
        #比對下載檔案跟本機檔案
            $FileDownload_Hash =  $FilesInDownloadZip | ForEach-Object {(Get-FileHash( (Get-ChildItem -Path ("$TEMP_Folder\$_")).FullName) -Algorithm SHA256).hash}
            $PCFile_Hash = $FilesInDownloadZip | ForEach-Object {(Get-FileHash( (Get-ChildItem -Path ("$ThreatSonar_Path\$_")).FullName) -Algorithm SHA256).hash}
            if((!$PCFile_Hash) -or !(@(Compare-Object $FileDownload_Hash $PCFile_Hash -sync 0).Length -eq 0)){
                #將新檔案替換舊檔
                $FilesInDownloadZip | ForEach-Object { Move-Item -Path "$TEMP_Folder\$_"  -Destination $ThreatSonar_Path -Force}
                #刪除壓縮檔清單以外的檔案
                Get-ChildItem -Path $ThreatSonar_Path -File -Recurse | Where-Object {$FilesInDownloadZip -notcontains $_.Name} | Remove-Item -Force
                #建立（或重建）排程
                schtasks /Delete /TN "ThreatSonar" /F                 
                schtasks /Create /TN ThreatSonar /RU SYSTEM /SC DAILY /RL HIGHEST /TR "$ThreatSonar_Path\ThreatSonar.exe" /ST $Specific_Time /F
            }else{
                $FilesInDownloadZip | ForEach-Object {Remove-Item "$TEMP_Folder\$_" -Force}
            }
        #釋放ZIP檔
        Start-Job -Name ZipClose -ScriptBlock { param($p1)        
            [IO.Compression.ZipFile]::OpenRead("$TEMP_Folder\$ThreatSonar_zip_FileName").Dispose()   
        } -ArgumentList "$TEMP_Folder\$ThreatSonar_zip_FileName"
        Wait-Job -Name ZipClose -Force
        Remove-Job -Name ZipClose -Force
        #刪除暫存檔案
        if("$TEMP_Folder\$ThreatSonar_zip_FileName"){Remove-Item "$TEMP_Folder\$ThreatSonar_zip_FileName" -Force }
        #上次程式錯誤的bug，移除下載檔
        if("$TEMP_Folder\$ThreatSonar_zip_FileName"){Remove-Item "$env:systemdrive\$ThreatSonar_zip_FileName" -Force }
                
    }
    
 #檢查是否安裝成功
    #檢查排程
    if((Get-ScheduledTask -TaskName "ThreatSonar" -ErrorAction Ignore).State -eq 'Ready'){ 
        $ScheduledTask_Check = "ThreatSonar排程已建立"        
        $LastResult = (Get-ScheduledTaskInfo -TaskName "ThreatSonar").LastTaskResult  
        $LastResult_Check = "排程上次執行結果： "      
        Switch ($LastResult){
            0 {$LastResult_Check = $LastResult_Check + "作業成功完成."}
            1 {$LastResult_Check = $LastResult_Check + "調用了不正確的函數或調用了未知的函數。"}
            2 {$LastResult_Check = $LastResult_Check + "檔案未找到。"}
            10 {$LastResult_Check = $LastResult_Check +"環境不正確"} 
            267008 {$LastResult_Check = $LastResult_Check +"排程工作已準備在下個預定時間執行"}
            267009 {$LastResult_Check = $LastResult_Check +"排程工作正在執行中" }
            267010 {$LastResult_Check = $LastResult_Check +"排程工作將不會在預定時間執行，因為該排程已停用" }
            267011 {$LastResult_Check = $LastResult_Check +"排程尚未執行"}
            267012 {$LastResult_Check = $LastResult_Check +"排程工作任務無其他排定執行期程" }
            267013 {$LastResult_Check = $LastResult_Check +"尚未設定此排程工作所需的一個或多個屬性" }
            267014 {$LastResult_Check = $LastResult_Check +"上一回的執行工作已經被使用者終止了。" }
            267015 {$LastResult_Check = $LastResult_Check +"任務沒有觸發器，或者現有觸發器被禁用或未設置" }
            2147750671 {$LastResult_Check = $LastResult_Check +"Credentials became corrupted." }
            2147750687 {$LastResult_Check = $LastResult_Check +"此排程工作的一個實例已在運行" }
            2147943645 {$LastResult_Check = $LastResult_Check +"The service is not available (is ""Run only when an user is logged on"" checked?)."}
            3221225786 {$LastResult_Check = $LastResult_Check +"應用程式因 CTRL+C 而終止."}
            3228369022 {$LastResult_Check = $LastResult_Check +"Unknown software exception."} 
            Default {$LastResult_Check = $LastResult_Check + "執行結果代碼為"+$LastResult}              
        }
    }else{
        $ScheduledTask_Check = "ThreatSonar排程不存在，嘗試從新建立..." 
        schtasks /Create /TN ThreatSonar /RU SYSTEM /SC DAILY /RL HIGHEST /TR "$ThreatSonar_Path\ThreatSonar.exe" /ST $Specific_Time /F
        if((Get-ScheduledTask -TaskName "ThreatSonar" -ErrorAction Ignore).State -eq 'Ready'){
            $ScheduledTask_Check = $ScheduledTask_Check + "排程建立成功！" 
            $LastResult_Check = "排程剛建立"
        }else{
            $ScheduledTask_Check = $ScheduledTask_Check + "排程建立失敗！"
            $LastResult_Check = "排程尚未建立"
        }
               
    }
    #檢查檔案有無缺漏
    $ThreatSonarFiles_Check = $Null
    $FilesInDownloadZip | ForEach-Object{if(Test-Path(Get-ChildItem -Path ("$ThreatSonar_Path\$_")).FullName){$ThreatSonarFiles_Check= $ThreatSonarFiles_Check + "$_ 檔案存在於 $ThreatSonar_Path`r`n"}else{$ThreatSonarFiles_Check = $ThreatSonarFiles_Check + "$_ 檔案不存在於 $ThreatSonar_Path`r`n"}}    
    #將檢查結果會出成log檔並上傳指定位置
    (get-date).ToString() + "`r`n$ScheduledTask_Check`r`n$LastResult_Check`r`n$ThreatSonarFiles_Check`r`n" | Out-File -FilePath "$TEMP_Folder\${env:COMPUTERNAME}_ThreatSonar_Check.txt"
    if(test-path "$env:allusersprofile\Task\sonar.log"){Copy-Item -Path "$env:allusersprofile\Task\sonar.log" -Destination "$TEMP_Folder\${env:COMPUTERNAME}_sonar.log" -Force }
    if(test-path "$env:allusersprofile\Task\sonarError.log"){Copy-Item -Path "$env:allusersprofile\Task\sonarError.log" -Destination "$TEMP_Folder\${env:COMPUTERNAME}_sonarError.log" -Force }    
    robocopy $TEMP_Folder $Log_Folder_Path "${env:COMPUTERNAME}_ThreatSonar_Check.txt" "${env:COMPUTERNAME}_sonar.log" "${env:COMPUTERNAME}_sonarError.log" "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null                
    <#  排程結果回傳碼及相對意義
        0 - The operation completed successfully.
        1 - Incorrect function called or unknown function called. 2 File not found.
        10 - The environment is incorrect. 
        267008 - Task is ready to run at its next scheduled time. 
        267009 - Task is currently running. 
        267010 - The task will not run at the scheduled times because it has been disabled. 
        267011 - Task has not yet run. 
        267012 - There are no more runs scheduled for this task. 
        267013 - One or more of the properties that are needed to run this task on a schedule have not been set. 
        267014 - The last run of the task was terminated by the user. 
        267015 - Either the task has no triggers or the existing triggers are disabled or not set. 
        2147750671 - Credentials became corrupted. 
        2147750687 - An instance of this task is already running. 
        2147943645 - The service is not available (is "Run only when an user is logged on" checked?). 
        3221225786 - The application terminated as a result of a CTRL+C. 
        3228369022 - Unknown software exception.        
    #>
    

