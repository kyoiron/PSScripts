if(Test-Path -Path ){
}

function Create-Shortcut{
    param (
        [string]$Name,
        [string]$TargetPath,
        [string]$ShortcutPath,
        [string]$IconPath = $null
    )

    # 建立 WScript.Shell 物件
    $shell = New-Object -ComObject WScript.Shell

    # 建立捷徑檔
    $shortcut = $shell.CreateShortcut($ShortcutPath)

    # 設定捷徑檔的屬性
    $shortcut.TargetPath = $TargetPath

    # 如果提供了圖示路徑，則設定圖示
    if ($IconPath) {
        $shortcut.IconLocation = $IconPath
    }


    # 儲存捷徑檔
    $shortcut.Save()
    


    
}


# 目標檔案路徑
        $targetPath = "$env:SystemDrive\eic\EicPrint\EicPrint.exe"
    
    # 捷徑檔的儲存路徑和檔名
        $shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\"+"$Name"+".lnk"
       
    # 圖示檔案路徑
        $iconPath = "$env:SystemDrive\eic\EicPrint\EicPrint.ico"
    
    # 建立 WScript.Shell 物件
        $shell = New-Object -ComObject WScript.Shell
    
    # 建立捷徑檔
        $shortcut = $shell.CreateShortcut($shortcutPath)
    
    # 設定捷徑檔的屬性
        $shortcut.TargetPath = $targetPath
    
    # 設定捷徑檔的圖示
        $shortcut.IconLocation = $iconPath
    
    # 儲存捷徑檔
        $shortcut.Save()