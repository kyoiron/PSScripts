$OU="tnd"
$OU_HOMEPAGE_1 = "http://www."+$OU+".moj"
$OU_HOMEPAGE_2 = $OU_HOMEPAGE_1 + "/mp.asp?mp=084"
$IE_HOMEPAHE_Path = 'HKCU:\Software\Microsoft\Internet Explorer\Main\'
$Start_Page = 'Start Page'
$Secondary_Start_Pages = 'Secondary Start Pages'
$First_Exist = (Get-Itemproperty -Path $IE_HOMEPAHE_Path -Name $Start_Page).$Start_Page -match $OU_HOMEPAGE_1 
#$Secondary_Start_Pages_key_Exist = (Get-ItemPropertyValue  $IE_HOMEPAHE_Path -name $Secondary_Start_Pages -ErrorAction SilentlyContinue)
$Secondary_Start_Pages_key_Exist = (Get-ItemProperty $IE_HOMEPAHE_Path)."Secondary Start Pages"
if($Secondary_Start_Pages_key_Exist){
    $Second_Exist = ![string]::IsNullOrEmpty((Get-Itemproperty -Path $IE_HOMEPAHE_Path -Name $Secondary_Start_Pages).$Secondary_Start_Pages -match $OU_HOMEPAGE_1)
}else{
    $Second_Exist = "False"
}
if((!($First_Exist -eq "true") -and (!($Second_Exist -eq "true")))){      
    $temp_Start_Page = (Get-Itemproperty -Path $IE_HOMEPAHE_Path -Name $Start_Page).$Start_Page 
    set-ItemProperty -Path $IE_HOMEPAHE_Path -Name $Start_Page -Value $OU_HOMEPAGE_1 
    [System.Collections.ArrayList]$temp_Second_pages = @()
    $temp_Second_pages.Add($temp_Start_Page)
    if(!$Secondary_Start_Pages_key_Exist){           
        New-ItemProperty -Path $IE_HOMEPAHE_Path -Name $Secondary_Start_Pages -PropertyType MultiString -force
    }
    (Get-Itemproperty -Path $IE_HOMEPAHE_Path -Name $Secondary_Start_Pages).$Secondary_Start_Pages  | where-Object{$temp_Second_pages.Add($_)}
    if($temp_Second_pages.Count -gt 0){Set-ItemProperty -Path $IE_HOMEPAHE_Path -Name $Secondary_Start_Pages $temp_Second_pages -Type MultiString}
}