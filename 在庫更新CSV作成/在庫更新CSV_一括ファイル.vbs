'ファイル名を格納する配列
Dim QtyBooks(1)
QtyBooks(0) = "在庫ファイル1.xlsm"
QtyBooks(1) = "在庫ファイル2.xlsm"

'ファイル操作、Dir取得FSOオブジェクト
Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")

dim CurrentPath
CurrentPath = Fso.GetFolder(".")

dim Mon,Today
Mon = Mid(date,6,2)

Today = Right(date,2)

'CSVとヘッダーを用意
Dim QtyCsv
Set QtyCsv = Fso.CreateTextFile(CurrentPath & "\" & "ヤフー在庫更新" & Mon & Today & ".csv")
QtyCsv.WriteLine("code,quantity,allow-overdraft")

QtyCsv.Close

'各ExcelのCSV追記マクロを呼び出す
Dim exApp
Set exApp = Wscript.CreateObject("Excel.Application")
exApp.Visible = True

For i =0 to 2

	Dim wb
	set wb = exApp.Workbooks.Open (Fso & "\" & QtyBooks(i))
	exApp.Run ("AppendQtyCsv")

	wb.Close True

Next

exApp.Quit
Set exApp = Nothing
