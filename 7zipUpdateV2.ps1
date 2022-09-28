#參數設定
$programName ="7-zip"
$version = "18.05"
$uninstalFileName = "uninstall.exe"
$msiFileNameX86 ="7z1805.msi"
$msiFileNameX64 ="7z1805-x64.msi"
$ResultLogD = $env:systemdrive+'\temp'
$Logfile = $ResultLogD+'\'+$env:computername+"_${programName}VersionCheck.txt"
$x64Reg = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$x86Reg = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

#如果沒有18.05的有安裝log，一律安裝（即使重複安裝）
IF(!(Test-Path $Logfile)){
    #從註冊檔取得移除資訊並移除（msi板安裝）
    $uninstall32 = gci $x64Reg | foreach { gp $_.PSPath } | ? { $_ -match "$programName" } | select UninstallString
    $uninstall64 = gci $x86Reg | foreach { gp $_.PSPath } | ? { $_ -match "$programName" } | select UninstallString    
    if ($uninstall64 -contains "msiexec.exe") {
        $uninstall64 = $uninstall64.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
        $uninstall64 = $uninstall64.Trim()
        #Write "Uninstalling..."
        start-process "msiexec.exe" -arg "/X $uninstall64 /quiet /norestart" -Wait -WindowStyle Hidden
    }
    if ($uninstall32 -contains "msiexec.exe") {
        $uninstall32 = $uninstall32.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
        $uninstall32 = $uninstall32.Trim()
        #Write "Uninstalling..."
        start-process "msiexec.exe" -arg "/X $uninstall32 /quiet /norestart" -Wait -WindowStyle Hidden
    }
    #exe版移除（exe版安裝且註冊檔沒有移除資訊）
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
            start-process $msiexec $arguments -Wait -WindowStyle Hidden
        }
    }else{
    
        if (Test-Path "$env:SystemDrive\temp\$msiFileNameX86" ) {
            $msiexec = "msiexec"
            $arguments = "/quiet  /i "+$envSystemDrive+"\temp\"+$msiFileNameX86+"  /log  "+$env:SystemDrive+"\temp\"+${env:computername}+"_"+${programName}+"Log.txt"
            start-process $msiexec $arguments  -Wait -WindowStyle Hidden
        }

    }
    Start-Sleep -s 3
    #檢查現在安裝之版本
    if (Test-Path $x86Reg){
        $x86InstalledProgram = ((Get-ChildItem $x86Reg) | Where-Object { $_.GetValue( "DisplayName" ) -like "*$programName*" })
    }
    IF ($x86InstalledProgram -ne $null){ 
         $x86InstalledVersion=$x86InstalledProgram.GetValue("DisplayVersion")
         if ($x86InstalledVersion -match "18.05.00.0"){
         
                     "已安裝程式名稱："+$programName+" 版本："+$x86InstalledVersion | Out-File -Append -FilePath $Logfile
         }
    }
    if (Test-Path $x64Reg){
        $x64InstalledProgram =  (Get-ChildItem $x64Reg) | Where-Object { $_.GetValue( "DisplayName" ) -like "*$programName*" }
    }
    IF ($x64InstalledProgram -ne $null){ 
        $x64InstalledVersion=$x64InstalledProgram.GetValue("DisplayVersion")
        if ($x64InstalledVersion -match "18.05.00.0"){

            "已安裝程式名稱："+$programName+" 版本："+$x64InstalledVersion | Out-File -Append -FilePath $Logfile
        }
    }
}