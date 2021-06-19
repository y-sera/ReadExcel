$SetExcel = {
    echo $FileList[$i]
    $script:workbook = $excel.Workbooks.Open($filepath+"\"+$FileList[$i])
}
$OpenSheet = {
    try{
        $script:worksheet = $workbook.Sheets($sheetname)
    } catch{
	echo "!!! No exist the default sheet. Read first sheet."
    echo "-------------------"
    $script:worksheet = $excel.worksheets.Item(1)
    }
}
$ReadExcel = {
	##<< test code.　Please change this part.
    $name=$worksheet.Range("C2").Text 
	echo "Name: $name"
	$profile=$worksheet.Range("C3").Text 
	echo "profile: $profile" 
        ## << test code.
}

$CloseBook = {
     Set-Variable -Name workbook -Scope script
     Set-Variable -Name worksheet -Scope script
     [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) > $null
     $workbook.Close()
     [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) > $null
}

$SelectFile = {
    Clear-Host
    echo "--- Excel file list. ---"
    $k=0
    $FileList | foreach{
      echo "$k : $_"
      $k++
    }
    echo ""
    $Filenum = Read-Host "Please select number of file to read."
    & $CloseBook
    try { 
    	$script:workbook = $excel.Workbooks.Open($filepath+"\"+$FileList[$Filenum])
    } catch { 
        echo "!!! No exist this file. Reread previous file."
    	echo "-------------------"
        $script:workbook = $excel.Workbooks.Open($filepath+"\"+$FileList[$i])
    }
    Write-Host "Please any key．．．" -NoNewLine
    [Console]::ReadKey($true) > $null
    Clear-host
}

$GetSheet ={
    Clear-Host
    $sheetTotal=$workbook.worksheets.count
    echo "Total Sheet: $sheetTotal"
    $sheetList= @($workbook.worksheets | ForEach-Object {
    $_.Name })
    echo ""
    echo "--- Sheet list ---"
    $j=0
    $Sheetlist | foreach{
      echo "$j : $_"
      $j++
      }
    echo ""
    $sheetnum = Read-Host "Please select sheet number to read"
    $script:sheetname= $sheetlist[$sheetnum]
    Write-Host "Please any key．．．" -NoNewLine
    [Console]::ReadKey($true) > $null
    Clear-host
}

$SetSheetName = {
    echo ""
    echo "--- Set sheetname ---"
    echo "Change default sheetname to read."
    echo "now default name: $script:SheetName"
    $script:SheetName= Read-Host "Please input sheetname."
    echo "new default sheetname : $SheetName"
    echo ""
}

#-------------------------------------------------------------------
$filepath=(pwd).Path

$FileList=@(dir -NAME | Select-String "^.*\.xlsx$" |sort)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $False
$excel.DisplayAlerts = $False

$script:sheetname= Read-Host "Please input default sheetname"

$i=0
& $SetExcel
& $OpenSheet
& $ReadExcel
while ($i -ge 0 -and $i -le ($FileList.Length-1 )) {
    	# Keyboad input
	Write-Host "Next file: n,  Before file: b,  Default SheetName: d,  Sheet Select: s,  File Select: f :" -NoNewLine
	$keyInfo = [Console]::ReadKey($true)
	switch ($keyInfo.Key){
	   "n" { #nextfile
           	Write-Host "next"
           	& $Closebook
            Clear-Host
            $i++
            if ($i -ge ($FileList.Length) ) {break}
            & $SetExcel
           	& $OpenSheet
		    & $ReadExcel
           } 
	   "b" { #before file
        	write-host "before"
            & $Closebook
            Clear-Host
            $i--
            if ($i -le -1 ) {break}
            & $SetExcel
            & $OpenSheet
		    & $ReadExcel
           } 
	   "d" { #default sheet name setting
            write-host "Default Sheet Name" 
            & $SetSheetName
           } 
	   "s" { #sheet select
            write-host "Sheet Select"
            & $GetSheet
            & $OpenSheet
		    & $ReadExcel
           } 
	   "f" { # file select
            write-host "File List"
            & $SelectFile
            & $OpenSheet
		    & $ReadExcel
           } 
          }
}
$excel.Quit()

echo "Finish read."

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) > $null

Write-Host "Please any key to finish．．．" -NoNewLine
[Console]::ReadKey($true) > $null