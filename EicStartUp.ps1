$Computer_StartUp = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
$Computer_StartUp_EICSignTSR_lnk = $Computer_StartUp +"\筆硯列印工具.lnk"
$Computer_StartUp_Print_lnk =  $Computer_StartUp +"\筆硯簽章工具.lnk"

$User_StartUp ="$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
$User_StartUp_EICSignTSR_lnk = $User_StartUp +"\筆硯列印工具.lnk"
$User_StartUp_Print_lnk =  $User_StartUp +"\筆硯簽章工具.lnk"

$EICSignTSR_lnk = "$env:PUBLIC\Desktop\筆硯列印工具.lnk"
$EicPrint_lnk = "$env:PUBLIC\Desktop\筆硯簽章工具.lnk"

if((Test-Path -Path $EICSignTSR_lnk) -and  (!(Test-Path -Path $User_StartUp_EICSignTSR_lnk))){
    Copy-Item  $EICSignTSR_lnk -Destination $User_StartUp
}

if( (Test-Path -Path $EicPrint_lnk) -and (!(Test-Path -Path $User_StartUp_Print_lnk)) ){
    Copy-Item  $EicPrint_lnk -Destination $User_StartUp
}