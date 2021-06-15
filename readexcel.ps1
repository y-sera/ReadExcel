$SetExcel = {
    echo $ReportList[$i]
    #ブックの指定
	$script:workbook = $excel.Workbooks.Open($filepath+"\"+$ReportList[$i])
	#シート名
	$script:sheetname=$month +"月" +$day +"日"
 
}
$ReadExcel = {
    #シートを指定
    #今日の日付の分が無ければ一枚目のシートを指定   
	try{
        $script:worksheet = $workbook.Sheets($sheetname)
    } catch{
	echo "※対象シートが存在しないため, 最前シートを読みこみます"
        $script:worksheet = $excel.worksheets.Item(1)
    }
    
    
	##テスト用コード
	echo "名前:" $worksheet.Range("C2").Text 
	echo ""
    echo "profile:" 
	echo $worksheet.Range("B3").Text 
  #  & $CloseBook
}

$CloseBook = {
Set-Variable -Name workbook -Scope script
Set-Variable -Name worksheet -Scope script
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
            $workbook.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
}

$selectfile = {
    Clear-Host
    echo "--- excel file list. ---"
    $k=0
    $Reportlist | foreach{
      echo "$k : $_"
      $k++
      }
    echo ""
    $Reportnum = Read-Host "読み込みたいファイルの番号を指定してください"
    & $CloseBook
    try { 
    	$script:workbook = $excel.Workbooks.Open($filepath+"\"+$ReportList[$Reportnum])
    } catch { 
    echo "該当ファイルが存在しないため, 現在のファイルを表示させます"
    	$script:workbook = $excel.Workbooks.Open($filepath+"\"+$ReportList[$i])
    }
    Write-Host "何かキーを押すとシートを表示します．．．" -NoNewLine
	[Console]::ReadKey($true) > $null
    Clear-host
}
$GetSheet ={
    Clear-Host
    $sheettotal=$workbook.worksheets.count
    echo "全シート数: $sheettotal"
    $sheetlist= @($workbook.worksheets | ForEach-Object {
    $_.Name })
    echo ""
    echo "--- sheet list ---"
    $j=0
    $Sheetlist | foreach{
      echo "$j : $_"
      $j++
      }
    echo ""
    $sheetnum = Read-Host "読みこみたいシートの番号を指定してください"
    $script:sheetname= $sheetlist[$sheetnum]
    Write-Host "何かキーを押すとシートを表示します．．．" -NoNewLine
	[Console]::ReadKey($true) > $null
    Clear-host

}

$setdate = {
    echo ""
    echo "--- set date ---"
    echo "読み込むシートの日付を設定します."
    echo "現在値: $script:sheetname"
    $month = Read-Host "月:" 
    $day = Read-Host "日:"
	$script:sheetname=$month +"月" +$day +"日"
    echo "シート名は $sheetname に設定されました."
    echo ""
}

#カレントディレクトリの取得
#相対パスで指定する場合
#[System.IO.Directory]::SetCurrentDirectory($pwd)
#絶対パスの取得
$filepath=(pwd).Path

#excelファイルの一覧を変数へ格納
$ReportList=@(dir -NAME | Select-String "^[0-9]{6}.*\.xlsx$" |sort)


# 日付の取得
$month=Get-Date -Format MM | % { $_ -replace "^0", ""}
$day=Get-Date -Format dd| % { $_ -replace "^0", ""}
$weekday=Get-Date -Format ddd


#新規COMオブジェクト作成
$excel = New-Object -ComObject Excel.Application
#ウィンドウ起動しない
$excel.Visible = $False
#読み取り専用警告オフ
$excel.DisplayAlerts = $False


$i=0
& $SetExcel
& $ReadExcel
while ($i -ge 0 -and $i -le ($ReportList.Length-1 )) {
    #キーボード入力
	 Write-Host "次ファイル:n  前ファイル:b 日付設定:d シート変更:s ファイル選択:f ::" -NoNewLine
	 $keyInfo = [Console]::ReadKey($true)
     switch ($keyInfo.Key){
	   "n" {
            Write-Host "next"
            & $Closebook
            Clear-Host
            $i++
            if ($i -ge ($ReportList.Length) ) {break}
             & $SetExcel
             & $ReadExcel
            } #nextfile
	   "b" {
            write-host "before"
            & $Closebook
            Clear-Host
            $i--
            if ($i -le -1 ) {break}
            & $SetExcel
            & $ReadExcel
            } #beforefile
	   "d" {
             write-host "day" 
             & $setdate
             } #day setting
	   "s" {
             write-host "sheet"
             & $GetSheet
             & $ReadExcel
             } #sheet change
	   "f" { 
            write-host "file list"
            & $selectfile
            & $ReadExcel
             } #menu 
          }
}
$excel.Quit()

echo "読み込みを終了しました"
echo "オブジェクトの解放を行います"


#変数の破棄
#
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

echo "enterキーを押してください"



	Write-Host "処理を継続する場合は何かキーを押してください．．．" -NoNewLine
	[Console]::ReadKey($true) > $null
