#excelファイルの一覧を変数へ格納
$ReportList=dir -NAME | Select-String "^[0-9]{6}.*\.xlsx$" |sort

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

#カレントディレクトリの取得
#相対パスで指定する場合
#[System.IO.Directory]::SetCurrentDirectory($pwd)
#絶対パスの取得
$filepath=(pwd).Path

# 一行ずつ処理する
foreach( $LINE in $ReportList ){
	

	Clear-Host
	#ブックの指定
	$workbook = $excel.Workbooks.Open($filepath+"\"+$LINE)
	#シート名
	$sheetname=$month +"月" +$day +"日"
	
    #シートを指定
    #今日の日付の分が無ければ一枚目のシートを指定
	try{
        $worksheet = $workbook.Sheets($sheetname)
    } catch{
	echo "※このファイルは, シート名が今日の日付ではありません"
        $worksheet = $excel.worksheets.Item(1)
    }

	##テスト用コード
	echo "名前:" $worksheet.Range("C2").Text 
	echo ""
    echo "profile:" 	
	echo $worksheet.Range("B3").Text 
		

	# キー入力待ち　
	Write-Host "処理を継続する場合は何かキーを押してください．．．" -NoNewLine
	[Console]::ReadKey($true) > $null

	#キーボード入力
	# Write-Host "次ファイル:n  前ファイル:b 日付設定:d シート変更:s メニュー:m ::" -NoNewLine
	# $keyInfo = [Console]::ReadKey($true)
       #  swich ($keyInfo.Key){
	 #  "n" {} #nextfile
	 #  "b" {} #beforefile
	 #  "d" {} #day setting
	 #  "s" {} #sheet change
	 #  "m" {} #menu 
	 # }
	
	$workbook.Close()
    Clear-Host
}
$excel.Quit()

echo "読み込みを終了しました"
echo "オブジェクトの解放を行います"

#変数の破棄
#
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

echo "enterキーを押してください"

#日付による分岐(金曜日ならプラス2日)
#シートの作成

# $newsheet=$tomrrowmonth + "月" + $tomorrowday + "日"
# シートHOGE1をコピーする
#$workbook.Worksheets.item($sheetname).copy($workbook.Worksheets.item($newsheet))
#月から木なら日付の数+1したシートを, 金曜日なら+3した日付のシートをコピーして作成
