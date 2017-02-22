Attribute VB_Name = "Module1"
Option Explicit
Sub ICOKURI連結()

    Call ConcatenateICOKURI
    
    '実行PCのデスクトップをフルパスで
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
        
    Dim PutFolder As String
    PutFolder = CStr(wsh.SpecialFolders("desktop")) & "\ICOKURI連結済み\"
        
    '実行したPCのデスクトップに「ICOKURI連結済みフォルダ」が無い場合
    If Dir(PutFolder) = "" Then
    
        PutFolder = Replace(PutFolder, "ICOKURI連結済み\", "")
    
    End If
    
    'マクロ起動ボタンを削除
    Sheet1.Shapes(1).Delete
    
    '連結したCSVをxlsx形式で保存
    Application.DisplayAlerts = False
    
        ThisWorkbook.SaveAs Filename:=PutFolder & "ICOKURI" & Format(Date, "MMdd") & ".xlsx"
     
    Application.DisplayAlerts = True
    
End Sub

Private Sub ConcatenateICOKURI()
Const ICOKURI_PC As String = "\\mos10\"

Dim FolderName(2) As String
FolderName(0) = "アマゾン宅配便"
FolderName(1) = "楽天発払"
FolderName(2) = "ヤフー\ヤフー発払い"

Dim i As Integer

For i = 0 To 2

    Dim CsvPath As String
    CsvPath = FindTodaysCSV(ICOKURI_PC & FolderName(i))
    
    If CsvPath <> "" Then
        Call ImportICOKURI(CsvPath)
    End If

Next

End Sub

Sub ImportICOKURI(ByVal CsvPath)
Attribute ImportICOKURI.VB_ProcData.VB_Invoke_Func = " \n14"

'最終行のRange、書き出し先セルの特定
Dim LastRow As Long
LastRow = Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row

Dim PutStartCell As Range

If LastRow = 1 Then
    Set PutStartCell = Range("A1")
Else
    Set PutStartCell = Cells(LastRow + 1, 1)
End If

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & CsvPath, Destination:=PutStartCell _
    )
    .Name = "ICOKURI"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
    2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

ActiveWorkbook.Connections(1).Delete

End Sub

Function FindTodaysCSV(ByVal CsvFolderPath As String) As String

'CSVフォルダのパス指定、最後\マークかチェック
If Not Right(CsvFolderPath, 1) = "\" Then
    CsvFolderPath = CsvFolderPath & "\"
End If

'実行時バインディングでファイルオブジェクト
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim f As Object, TodayCSV As Object
      
'指定フォルダー内のFileNameを含むファイル名を調べて、本日 日付ファイルを一つ取得する

For Each f In FSO.GetFolder(CsvFolderPath).Files

    If DateDiff("D", f.DateLastModified, DateValue(Date)) = 0 And f.Name Like "ICOKURI*" Then
    
        Set TodayCSV = f
        Exit For
        
    End If

Next

If f Is Nothing Then

      Exit Function
      
End If

FindTodaysCSV = CsvFolderPath & f.Name

End Function
