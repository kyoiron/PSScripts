#Requires -version 2.0
    Add-Type -AssemblyName microsoft.office.interop.outlook 
#規則名稱
    $RuleTitleName = "寄件者郵件IP過濾"
#可疑郵件IP
    $SuspectIP = @('117.56.7.27','117.56.7.28')
#未設定過outlook，則程式結束不執行
$IsNotExistOutlookDefaultProfile_2010Before = ((Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles' -ErrorAction SilentlyContinue).DefaultProfile -eq $null) 
$isNotExistOutlookDefaultProfile_2013 = ((Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\15.0\Outlook' -ErrorAction SilentlyContinue).DefaultProfile -eq $null)
$isNotExistOutlookDefaultProfile_2016 = ((Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\Outlook' -ErrorAction SilentlyContinue).DefaultProfile -eq $null)
if($IsNotExistOutlookDefaultProfile_2010Before -and $isNotExistOutlookDefaultProfile_2013 -and $isNotExistOutlookDefaultProfile_2016){exit}

#取得Outlook及其過濾規則
    $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
    $olRuleType = "Microsoft.Office.Interop.Outlook.OlRuleType" -as [type]
    $outlook = New-Object -ComObject outlook.application
    $namespace  = $Outlook.GetNameSpace("mapi")
    $inbox = $namespace.getDefaultFolder($olFolders::olFolderInbox)
    $rules = $outlook.session.DefaultStore.GetRules()

    $isExistSameRuleName = $fals
    $isSuspectMailTitlesSame = $false
    #檢驗是否有相同名稱的規則否存
    $isExistSameRuleName = ($rules | Where-Object{ $_.name -eq $RuleTitleName}).name -contains $RuleTitleName
    #檢視相同名稱規則是否過濾條件一樣
    if ($isExistSameRuleName -eq $true){ 
        $isSuspectMailTitlesSame = ((Compare-Object -ReferenceObject $SuspectIP -DifferenceObject (($rules |Where-Object{$_.name -eq $RuleTitleName}).conditions.MessageHeader.text) -PassThru).count -eq 0)
        if($isSuspectMailTitlesSame -eq $true){
            #過濾規則已存在，過濾規則相同，程式結束
            exit
        }else{        
            #刪除同名規則
            #$rule = ($rules | Where-Object{$_.name -eq $RuleTitleName})
            $rules.Remove(($rules | Where-Object{$_.name -eq $RuleTitleName}).name)
            $rules.Save()       
        }
    }        

#建立新規則
    $rule = $rules.Create($RuleTitleName,$olRuleType::OlRuleReceive)
#設定規則之「條件」
    $SubjectCondition = $rule.Conditions.MessageHeader
    $SubjectCondition.text = $SuspectIP
    $SubjectCondition.Enabled = $true
#設定規則之「動作」
    #丟到「刪除的郵件」
        #$MoveRuleAction = $rule.actions.Delete
    #丟到「直接刪除」
        $MoveRuleAction = $rule.actions.DeletePermanently
    $MoveRuleAction.Enabled = $true
#將設定之規則存檔
    $rules.Save()
#立刻執行此次新設定規則，背景執行不會跳出outlook
    $rule.execute()