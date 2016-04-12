Dim QtyBooks(2)
QtyBooks(0) = "在庫ファイル01.xlsm"
QtyBooks(1) = "在庫ファイル02.xlsm"
QtyBooks(2) = "在庫ファイル03.xlsm"

Dim objPath
Set objPath = CreateObject("Scripting.FileSystemObject").GetFolder(".")

Dim exApp
Set exApp = Wscript.CreateObject("Excel.Application")
exApp.Visible = True

For i =0 to 2

	Dim wb
	set wb = exApp.Workbooks.Open (objPath & "\" & QtyBooks(i))
	exApp.Run ("CSV生成")

	wb.Close True

Next

exApp.Quit
Set exApp = Nothing
