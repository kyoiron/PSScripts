#參數設定
$EmailPOP3UserName = 'lyc1116'
$EmailAccountName = "$EmailPOP3UserName@mail.moj.gov.tw"
$EmailDisplayName = "劉耀謙"
$EmailAddress = $EmailAccountName
$defaultOutlookFolder = "D:\Outlook"

#將Outlook預設之pst存檔位置改到D:\Outlook
#Outlook版本號
#Outlook 2007:12.0
#Outlook 2010:14.0
#Outlook 2013:15.0
#Outlook 2016:16.0
$OfficeVersionArray = '12.0','14.0','15.0','16.0'
foreach ($item in $OfficeVersionArray){

    write-host Test-Path "HKCU:Software\Microsoft\Office\$item\Outlook"
    if (Test-Path "HKCU:Software\Microsoft\Office\$item\Outlook")
    {
          New-ItemProperty -Path "HKCU:Software\Microsoft\Office\$item\Outlook" -Name ForcePSTPath -PropertyType ExpandString -Value $defaultOutlookFolder |  Out-Null
    }
}
#確認是否存在預設outlook pst檔預設資料夾，如無則新增一個
if (-Not(Test-Path d:\Outlook)){
    New-Item D:\Outlook -type Directory
}

#取的PRF的範本檔並將修改參數
#手動改為：Templates.PRF用記事本打開後，修改第53、56、57、59行
$PRFfile = ((((Get-Content -Path \\172.29.205.114\Public\sources\Outlook設定\Templates.PRF -Encoding Unknown) -replace '%AccountName%' ,"$EmailAccountName") -replace '%POP3UserName%' ,"$EmailPOP3UserName") -replace '%EmailAddress%' , "$EmailAddress") -replace '%DisplayName%' , "$EmailDisplayName" -replace "%PSTFileFolder%",$defaultOutlookFolder
$TXTPRFfile = "$env:TEMP\${EmailPOP3UserName}_PRFfile.prf"
Set-Content -Path $TXTPRFfile -Value $PRFfile
#檢查現在電腦有沒有執行Outlook，如果有，必須先關掉才可以執行設定動作
$OutlookProcess = Get-Process outlook -ErrorAction SilentlyContinue
if ($OutlookProcess){
    #取得outlook.exe所在的位置
    $OutlookPath = $OutlookProcess.Path
    $OutlookProcess.CloseMainWindow()
    Sleep 5
    if (!$OutlookProcess.HasExited) {
        $OutlookProcess | Stop-Process -Force
    }
}else{
    
    $Outlook = New-Object -comobject Outlook.Application
    #取得outlook.exe所在的位置
    $OutlookPath = (Get-Process "outlook").Path
    $Outlook.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)
    Remove-Variable Outlook
}


#將TRF檔匯入
#$AddOutlookAccountcommand = """"+$OutlookPath + """ /importprf " + $TXTPRFfile
#顯示要執行的指令，debug用
#write-host $AddOutlookAccountcommand
Start-Process -FilePath $OutlookPath -arg " /importprf $TXTPRFfile " -Wait -WindowStyle Hidden
