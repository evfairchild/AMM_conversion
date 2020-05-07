Write-Host -NoNewline "Converting AMM."

$root_dir = (Get-Item -Path ".\").FullName + "\"
#$root_dir = "\\bgmfiler01\G_Drive\TechOps\Engineering\Reliability\AMM_script\"
$filename = Get-ChildItem $root_dir\*.SGM | select -expand BaseName
$original_file = $root_dir + $filename + '.SGM'
$conversion1_file = $root_dir + 'amm_conversion1.txt'
$conversion2_file = $root_dir + 'amm_conversion2.txt'
$conversion3_file = $root_dir + 'amm_conversion3.txt'
$final_file = $root_dir + 'amm.txt'
$remove_str = Get-Content ($root_dir + 'remove_str.txt')  #### AMMEND THIS FILE WITH STRINGS TO BE DELETED

# Text Manipulation --------------------------------------------------------------------------------------------------------------------------------
Write-Host -NoNewline "."
(get-content $original_file | Out-String ) -replace '<SUBTASK', "`r`n<SUBTASK" -replace '<TASK', "`r`n<TASK" | Out-File $conversion1_file
Write-Host -NoNewline "." 
get-content $conversion1_file -ReadCount 1000 | ForEach-Object { $_ -match "<TASK|<SUBTASK"} > $conversion2_file
Write-Host -NoNewline "."
get-content $conversion2_file -ReadCount 1000 | ForEach-Object { $_.TrimEnd() } | Set-Content $conversion1_file
Write-Host ".done."

#Remove Unused Columns of DATA.  Reduce file size (a lot!). ----------------------------------------------------------------------------------------
Write-Host "Deleting Unused Data..." 
(get-content $conversion1_file) | ForEach-Object { $_.substring(1,$_.IndexOf("SUBJNBR")+22) } | Out-File $conversion2_file 

#Remove Erroneous data -----------------------------------------------------------------------------------------------------------------------------
Write-Host "Deleting the following strings:"  
for ($i = 0;$i -le $remove_str.Length;$i++){
    $str = $remove_str[$i] -replace "`n", "" -replace "`r", ""
    (Get-content $conversion2_file | Out-String) -replace $str, '' | Out-File $conversion2_file
    Write-Host $str
}

#Add Spaces to make a Space-Delimited text file ----------------------------------------------------------------------------------------------------
(Get-Content $conversion2_file).Replace(">", " ") | Set-Content $conversion1_file

#Remove TASKREQ rows.  Only TASKS and SUBTASKS should exist. ---------------------------------------------------------------------------------------
Get-Content $conversion1_file | Where { $_ -notmatch "TASKREQ" } | Set-Content $conversion2_file

#Add 'CONFNBR=' to rows missing this field ---------------------------------------------------------------------------------------------------------

#Remove whitespace ---------------------------------------------------------------------------------------------------------------------------------
get-content $conversion2_file -ReadCount 1000 | ForEach-Object { $_.TrimEnd() } | Set-Content $final_file

Write-Host "`nRemoving temp files...done."
Remove-Item $conversion1_file
Remove-Item $conversion2_file

Write-Host -NoNewline "Starting Excel Application..."
$excel = New-Object -ComObject excel.application
$excel.Visible = $true
Write-Host ".done."

Write-Host -NoNewline "Beginning text import macro..."
$workbook = $excel.Workbooks.Open($root_dir + "AMMmacro.xlsm")
Write-Host ".done."

Write-Host -NoNewline "Excel now importing text..."
$excel.Run("textImportInput", $final_file)
$workbook = $excel.Workbooks.Open($root_dir + "AMMmacro.xlsm")
$excel.Run("FormatAMM", $root_dir + 'AMM.xlsx')
Write-Host ".done."

$workbook.Close()
#$excel.Quit()

Read-Host "Press any key to exit."
exit