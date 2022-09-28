#檢查是否安裝較新版程式
function Is-InstalledOlder([string] $program ,[string] $version ) {
     
    #$x86InstalledProgramVersion = ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") | Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" }).GetValue("DisplayVersion")
    #$x64InstalledProgramVersion = ((Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") | Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" }).GetValue("DisplayVersion")
    $x86InstalledProgram =  ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") | Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" })
    IF (($x86InstalledProgram) -ne $null ){
        $x86InstalledProgramVersion = $x86InstalledProgram.GetValue("DisplayVersion")
        if($x86InstalledProgramVersion -eq $null){
            $x86InstalledProgramVersion = $x86InstalledProgram.GetValue("displayname").split(" ").split(".")[1] + $x86InstalledProgram.GetValue("displayname").split(" ").split(".")[2]
            $x86 = [version]$x86InstalledProgramVersion -lt [version]$version
             
        }else{
        
            if (![string]::IsNullOrEmpty($x86InstalledProgramVersion)) { $x86 = [version]$x86InstalledProgramVersion -lt [version]$version}
        }
    }
    
    if (Test-Path "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"){
        $x64InstalledProgram =  ((Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") | Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" })
        IF (($x64InstalledProgram) -ne $null){
            $x64InstalledProgramVersion = $x64InstalledProgram.GetValue("DisplayVersion")
            
            if (![string]::IsNullOrEmpty($x64InstalledProgramVersion)) { $x64 = [version]$x64InstalledProgramVersion -lt [version]$version}
        }
    }
    return $x86 -or $x64;
}

#檢查現在安裝之版本並匯出log檔
Function OutputCheck_Version($fprogram,$fLogfile)
{
    #檢查現在安裝之版本
    $x86InstalledProgram = ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") | Where-Object { $_.GetValue( "DisplayName" ) -like "*$fprogram*" })
    if (Test-Path "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"){
        $x64InstalledProgram =  (Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") | Where-Object { $_.GetValue( "DisplayName" ) -like "*$fprogram*" }
    }

    If (!(Test-Path "$fLogfile" )){
        IF ($x86InstalledProgram -ne $null){ 
            $x86InstalledVersion=$x86InstalledProgram.GetValue("DisplayVersion")
            "已安裝程式名稱："+$programName+" 版本："+$x86InstalledVersion | Out-File -Append -FilePath $fLogfile
        }
        IF ($x64InstalledProgram -ne $null){ 
            $x64InstalledVersion=$x64InstalledProgram.GetValue("DisplayVersion")
            "已安裝程式名稱："+$programName+" 版本："+$x64InstalledVersion | Out-File -Append -FilePath $fLogfile
        }
   }
}
    

#參數設定
$programName ="7-zip"
$version = "18.05"
$uninstalFileName = "uninstall.exe"
$msiFileNameX86 ="7z1805.msi"
$msiFileNameX64 ="7z1805-x64.msi"
$ResultLogD = $env:systemdrive+'\temp'
$Logfile = $ResultLogD+'\'+$env:computername+"_${programName}VersionCheck.txt"


#檢查是否有小於1805版
if ((Is-InstalledOlder $programName $version) -eq $True) {  
   
    #msi版或有於windows中有登錄移除資訊
    $uninstall32 = gci "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { $_ -match "$programName" } | select UninstallString
    $uninstall64 = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { $_ -match "$programName" } | select UninstallString
    
    if ($uninstall64) {
        $uninstall64 = $uninstall64.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
        $uninstall64 = $uninstall64.Trim()
        #Write "Uninstalling..."
        start-process "msiexec.exe" -arg "/X $uninstall64 /quiet" -Wait
    }
    if ($uninstall32) {
        $uninstall32 = $uninstall32.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
        $uninstall32 = $uninstall32.Trim()
        #Write "Uninstalling..."
        start-process "msiexec.exe" -arg "/X $uninstall32 //quiet" -Wait
    }
    
    #exe版移除
    #檢查是否有7-zip有就移除
    if((Test-Path $env:ProgramFiles\$ProgramName ) -and (Test-Path "$env:ProgramFiles\$ProgramName\$uninstalFileName")){
            Start-Process -FilePath "$env:ProgramFiles\$ProgramName\$uninstalFileName" -ArgumentList '/S', '/V', '/qn' , '/norestart'
    }
    #如果os是64位元，額外檢查Program Files (x86)資料夾
    if ([Environment]::Is64BitOperatingSystem){
        
        if ((Test-Path ${env:ProgramFiless(x86)}\$ProgramName ) -and (Test-Path "${env:ProgramFiless(x86)}\$ProgramName\$uninstalFileName" )){
            Start-Process -FilePath "${env:ProgramFiles(x86)}\$ProgramName\$uninstalFileName" -ArgumentList '/S', '/V', '/qn' , '/norestart'
        }
    }
    
    #安裝新版本
    if ([Environment]::Is64BitOperatingSystem){

        if (Test-Path "$env:SystemDrive\temp\$msiFileNameX64" ) {
            $msiexec = "msiexec"
            $arguments = "/quiet  /i "+$envSystemDrive+"\temp\"+$msiFileNameX64+"  /log  "+$env:SystemDrive+"\temp\"+${env:computername}+"_"+${programName}+"Log.txt"
            start-process $msiexec $arguments -Wait
            
    
        }

    }else{
    
        if (Test-Path "$env:SystemDrive\temp\$msiFileNameX86" ) {
            $msiexec = "msiexec"
            $arguments = "/quiet  /i "+$envSystemDrive+"\temp\"+$msiFileNameX86+"  /log  "+$env:SystemDrive+"\temp\"+${env:computername}+"_"+${programName}+"Log.txt"
            start-process $msiexec $arguments   -Wait
            
        }

    }
    
    OutputCheck_Version $programName $Logfile


}else{
    
    OutputCheck_Version $programName $Logfile

}