$ReadExcel = {
    
    echo $ReportList[$i]
    #ブックの指定
	$script:workbook = $excel.Workbooks.Open($filepath+"\"+$ReportList[$i])
	#シート名
	$script:sheetname=$month +"月" +$day +"日"
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

$Menu = {
    Clear-Host
    echo "--- excel file list. ---"
    echo $Reportlist
    echo "(現在読みこんでるブックの前後を表示させたい)"
    echo ""
    echo "sheet list of now book"
    echo "現在読んでいるシートのリストを表示"
    echo "出来れば, キー入力で動的にリストを更新できると尚良い"
}
$GetSheet ={
    Clear-Host
    $sheettotal=$workbook.worksheets.count
    echo "全シート数: $sheettotal"
    $sheetlist= @($workbook.worksheets | ForEach-Object {
    $_.Name })
    echo ""
    echo "--- sheet list ---"
    echo $sheetlist
    echo "シートタイトルを列挙して表示(現在のシートの周り10件)"
    echo "読みこみたいシート名を指定してください"
    
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
& $ReadExcel
while ($i -ge 0 -and $i -le ($ReportList.Length-1 )) {
    #キーボード入力
	 Write-Host "次ファイル:n  前ファイル:b 日付設定:d シート変更:s メニュー:m ::" -NoNewLine
	 $keyInfo = [Console]::ReadKey($true)
     switch ($keyInfo.Key){
	   "n" {
            Write-Host "next"
            & $Closebook
            Clear-Host
            $i++
            if ($i -ge ($ReportList.Length) ) {break}
             & $ReadExcel
            } #nextfile
	   "b" {
            write-host "before"
            & $Closebook
            Clear-Host
            $i--
            if ($i -le -1 ) {break}
            & $ReadExcel
            } #beforefile
	   "d" {
             write-host "day" 
             } #day setting
	   "s" {
             write-host "sheet"
             & $GetSheet
             } #sheet change
	   "m" { 
            write-host "menu"
            & $Menu
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

#日付による分岐(金曜日ならプラス2日)
#シートの作成

# $newsheet=$tomrrowmonth + "月" + $tomorrowday + "日"
# シートHOGE1をコピーする
#$workbook.Worksheets.item($sheetname).copy($workbook.Worksheets.item($newsheet))
#月から木なら日付の数+1したシートを, 金曜日なら+3した日付のシートをコピーして作成


	Write-Host "処理を継続する場合は何かキーを押してください．．．" -NoNewLine
	[Console]::ReadKey($true) > $null
