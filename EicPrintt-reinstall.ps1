#LOG檔NAS存放路徑
$Log_Path = "\\172.29.205.114\Public\sources\audit"
$Log_Path_Folder = "$Log_Path\EicPrintt"
#更新檔安裝路徑
$EicPrint_EXE_Folder = "\\172.29.205.114\loginscript\Update\EicPrintt"
#擷取檔案資訊
$EicPrint_EXE_NAS = Get-ChildItem -Path "$EicPrint_EXE_Folder\EicPrint.exe"
$EicPrint_EXE_PC = Get-ChildItem -Path "$env:SystemDrive\eic\EicPrint\EicPrint.exe"

#進行版本比較
if([version]$EicPrintt_EXE_PC.VersionInfo.FileVersion -lt [version]$EicPrintt_EXE_NAS.VersionInfo.FileVersion){ 
    #檢查是否已經啟動
    $Restart_EicPrint_EXE=$null
    $Restart_EicPrint_EXE = Get-Process -Name EicPrint -ErrorAction SilentlyContinue
    #如果已經啟動則強制關閉程式
    while((Get-Process -Name EicPrint -ErrorAction SilentlyContinue)|Where-Object {!$_.HasExited}){
        Stop-Process -Name EicPrint -Force -ErrorAction SilentlyContinue
    }
    #將舊版本檔案重新命名附加文字old和版本號碼
    Rename-Item -Path $EicPrint_EXE_PC -NewName ($EicPrint_EXE_PC.BaseName + "_old_"+ $EicPrint_EXE_PC.VersionInfo.FileVersion +".exe")
    #將新版本檔案同步
    robocopy $EicPrint_EXE_Folder "$env:SystemDrive\eic\EicPrint" "EicPrint.exe" "/XO /NJH /NJS /NDL /NC /NS".Split(' ') | Out-Null
    #刪除 C:\Users\Public\EICSignTSR\EicSignTSR.ini
    if(Test-Path -Path "$env:PUBLIC\EicPrint\EicPrint.ini"){Remove-Item -Path "$env:PUBLIC\EicPrint\EicPrint.ini"}
    #檢查現在程式之版本
    $EicPrint_EXE_PC_Check = Get-ChildItem -Path "$env:SystemDrive\eic\EicPrint\EicPrint.exe"
    #檢查log存放之資料夾是否存在，不存在則建立新資料夾    
    if(!(Test-Path -Path $Log_Path_Folder)){New-Item -ItemType Directory -Path $Log_Path_Folder -Force}
    #進行更本後之版本比較並匯出成文字檔
    if([version]$EicPrint_EXE_PC_Check.VersionInfo.FileVersion -eq [version]$EicPrint_EXE_NAS.VersionInfo.FileVersion){
        "版本成功更新為  " +  $EicPrint_EXE_PC.VersionInfo.FileVersion | Out-File ("$Log_Path_Folder\${env:COMPUTERNAME}_"+$EicPrint_EXE_PC_Check.VersionInfo.ProductName+"_"+$EicPrint_EXE_PC_Check.VersionInfo.FileVersion+"_Successes.txt")
    }else{
        "版本成功失敗，程式版本仍為 " + $EicPrint_EXE_PC_Check.VersionInfo.FileVersion | Out-File ("$Log_Path_Folder\${env:COMPUTERNAME}_"+$EicPrint_EXE_PC_Check.VersionInfo.ProductName+"_"+$EicPrint_EXE_PC_Check.VersionInfo.FileVersion + "_Fail.txt")
    }
    #如果更板前已經啟動程式則更版後執行程式
    if($Restart_EicPrint_EXE -ne $null){
        & $EicPrint_EXE_PC
    }
}