<#
    $EICSignTSR_Ini= $env:PUBLIC + "\EICSignTSR\EICSIGNTSR.INI"
    $Check_CONCURRENT = Select-String -Path $EICSignTSR_Ini -Pattern "CONCURRENT_ENABLE=1"
    if($Check_CONCURRENT -eq $null){    
        Add-Content $EICSignTSR_Ini "CONCURRENT_ENABLE=1"
    }
#>
#LOG檔NAS存放路徑
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$Log_Path_Folder = "$Log_Path\EicSignTSR"
#更新檔安裝路徑
$EicSignTSR_EXE_Folder = "\\172.29.205.114\loginscript\Update\EICSignTSR"
#擷取檔案資訊
$EicSignTSR_EXE_NAS = Get-ChildItem -Path "$EicSignTSR_EXE_Folder\EicSignTSR.exe"
$EicSignTSR_EXE_PC = Get-ChildItem -Path "$env:SystemDrive\eic\EICSignTSR\EicSignTSR.exe"
#進行版本比較
if([version]$EicSignTSR_EXE_PC.VersionInfo.FileVersion -lt [version]$EicSignTSR_EXE_NAS.VersionInfo.FileVersion){ 
    #檢查是否已經啟動
    $Restart_EicSignTSR_EXE = Get-Process -Name EicSignTSR -ErrorAction SilentlyContinue
    #如果已經啟動則強制關閉程式
    while((Get-Process -Name EicSignTSR -ErrorAction SilentlyContinue)|Where-Object {!$_.HasExited}){
        Stop-Process -Name EicSignTSR -Force -ErrorAction SilentlyContinue
    }
    #將舊版本檔案重新命名附加文字old和版本號碼
    Rename-Item -Path $EicSignTSR_EXE_PC -NewName ($EicSignTSR_EXE_PC.BaseName + "_old_"+ $EicSignTSR_EXE_PC.VersionInfo.FileVersion +".exe")
    #將新版本檔案同步
    robocopy $EicSignTSR_EXE_Folder "$env:SystemDrive\eic\EICSignTSR" "EicSignTSR.exe" "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    #刪除 C:\Users\Public\EICSignTSR\EicSignTSR.ini
    if(Test-Path -Path "$env:PUBLIC\EICSignTSR\EicSignTSR.ini"){Remove-Item -Path "$env:PUBLIC\EICSignTSR\EicSignTSR.ini"}
    #檢查現在程式之版本
    $EicSignTSR_EXE_PC_Check = Get-ChildItem -Path "$env:SystemDrive\eic\EICSignTSR\EicSignTSR.exe"
    #檢查log存放之資料夾是否存在，不存在則建立新資料夾    
    if(!(Test-Path -Path $Log_Path_Folder)){New-Item -ItemType Directory -Path $Log_Path_Folder -Force}
    #進行更本後之版本比較並匯出成文字檔
    if([version]$EicSignTSR_EXE_PC_Check.VersionInfo.FileVersion -eq [version]$EicSignTSR_EXE_NAS.VersionInfo.FileVersion){
        "版本成功更新為  " +  $EicSignTSR_EXE_PC.VersionInfo.FileVersion | Out-File ("$Log_Path_Folder\${env:COMPUTERNAME}_"+$EicSignTSR_EXE_PC_Check.VersionInfo.ProductName+"_"+$EicSignTSR_EXE_PC_Check.VersionInfo.FileVersion+"_Successes.txt")
    }else{
        "版本成功失敗，程式版本仍為 " + $EicSignTSR_EXE_PC_Check.VersionInfo.FileVersion | Out-File ("$Log_Path_Folder\${env:COMPUTERNAME}_"+$EicSignTSR_EXE_PC_Check.VersionInfo.ProductName+"_"+$EicSignTSR_EXE_PC_Check.VersionInfo.FileVersion + "_Fail.txt")
    }
    #如果更板前已經啟動程式則更版後執行程式
    if($Restart_EicSignTSR_EXE -ne $null){
        & $EicSignTSR_EXE_PC
    }
}


