#共用參數設定
    #機關名稱，須自行修改
        $Name="台南看守所"    
    #LOG檔存放路徑，須自行修改
        $Log_Folder_Path = "\\172.29.205.114\Public\sources\audit\ThreatSonarPC"
    #安裝檔存放路徑
        $ThreatSonar_exe_FileName =  "ThreatSonarAnti-Ransomware_$Name-PC.exe" 
        $url_exe = "http://download.moj/files/ThreatSonar/PC-各機關ThreatSonar安裝檔/矯正機關/$Name-PC/$ThreatSonar_exe_FileName"
        $NAS_Path = "\\172.29.205.114\loginscript\Update\ThreatSonar\"
        $NAS_exe = $NAS_Path + $ThreatSonar_exe_FileName 
    #預設執行時間，可自訂，11:00為預設，請用24H制
        $Specific_Time = "11:00"
    #舊版ThreatSona執行檔檔案存放路徑
        $ThreatSonar_Path = "$env:SystemDrive\ThreatSonar"
    #新版ThreatSona執行檔檔案存放路徑
        $ThreatSonar_New_FolderName = "ThreatSonar2023"
        $ThreatSonar_New_Path = "$env:SystemDrive\$ThreatSonar_New_FolderName"
    #新版ThreatSonar.exe的版本
        $ThreatSonar2023_exe_version = "2203p2"
        $ThreatSonar2023_exe_version_order = "298"
    #移除舊版
        $Need_Remove_OldVerion = $true


function Remove_OldVersion(){
    # 停止進程
        Stop-Process -Name "ThreatSonar" -Force

    # 刪除服務
        sc.exe delete ThreatSonar
    
    # 刪除計劃任務
        schtasks.exe /Delete /TN ThreatSonar /F

    # 刪除檔案
        Remove-Item -Path "$env:ProgramData\Task\malicious_result*" -Force -ErrorAction SilentlyContinue -Recurse
        Remove-Item -Path "$env:ProgramData\Task\Server_Config*" -Force -ErrorAction SilentlyContinue -Recurse
        Remove-Item -Path "$env:ProgramData\Task\sonarError.log*" -Force -ErrorAction SilentlyContinue -Recurse
        Remove-Item -Path "$env:ProgramData\Task\sonar.log*" -Force -ErrorAction SilentlyContinue -Recurse
        Remove-Item -Path "$env:ProgramData\Task*" -Force -ErrorAction SilentlyContinue  -Recurse
        Remove-Item -Path "$env:ProgramData\sonar*.log" -Force -ErrorAction SilentlyContinue -Recurse
        Remove-Item -Path "$env:SystemDrive\ThreatSonar*" -Force -ErrorAction SilentlyContinue -Recurse
        Remove-Item -Path "$env:ProgramFiles(x86)\ThreatSonar*" -Force -ErrorAction SilentlyContinue -Recurse
    # 刪除資料夾
        Remove-Item -Path "$env:ProgramData\Task\malicious_result" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$env:ProgramData\Task" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$env:ProgramData\sonar" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$env:SystemDrive\ThreatSonar" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$env:ProgramFiles(x86)\ThreatSonar\" -Recurse -Force -ErrorAction SilentlyContinue
}
        
#檢查是否有安裝舊版    
    if((Test-Path "$ThreatSonar_Path") -eq $true){
        if((Get-ChildItem -Path "$ThreatSonar_Path").count -eq 0){
            Remove-Item -Path $ThreatSonar_Path -Force
        }else{
            $Need_Remove_OldVerion = $true
            Remove_OldVersion
        }     
    }

#安裝新版
    

    if(((Get-Service -Name ThreatSonar -ErrorAction SilentlyContinue) -eq $null) -or  $Need_Remove_OldVerion -eq $true){
        #Copy-Item -Path $NAS_exe -Destination "$env:SystemDrive\temp"
        robocopy $NAS_Path "$env:systemdrive\temp" $ThreatSonar_exe_FileName "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null   
        &"$env:SystemDrive\temp\$ThreatSonar_exe_FileName" -wait
             
        Start-Sleep -Seconds 1
        if(Test-Path -Path "$env:SystemDrive\temp\$ThreatSonar_exe_FileName"){
            Remove-Item -Path "$env:SystemDrive\temp\$ThreatSonar_exe_FileName" -Force 
            #$parameter = "--input 'SomeInputValue'"
            #&del "$env:SystemDrive\temp\$ThreatSonar_exe_FileName" "--input '/Q /F'"            
        }
        if(Test-Path -Path "$env:SystemDrive\temp\ThreatSonar"){
            Remove-Item -Path "$env:SystemDrive\temp\ThreatSonar" -Force 

        }

    }



