Attribute VB_Name = "LoadPurchaseReq"
Option Explicit

Const PICKING_FOLDER As String = "\\server02\商品部\ネット販売関連\ピッキング\"

Sub LoadAllPicking()
'手配依頼チェック済のピッキングファイルを一括して読込
'手配依頼として背景色が変えてある行をコピーします。

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'セラー分ピッキングファイル読み込み
Dim PickingFiles As Variant, File As Variant

PickingFiles = Array( _
    "ピッキングシート", _
    "楽天Pシート", _
    "ヤフーPシート" _
    )

For Each File In PickingFiles
    Call LoadSellerPicking(CStr(File) & Format(Date, "MMdd") & "-a.xlsx")
Next

'卸分 ファイル読み込み
PickingFiles = Array( _
    "アマゾン棚なし" & Format(Date, "MMdd") & ".xlsx", _
    "アマゾン棚なし" & Format(Date, "MMdd") & "-outdoor.xlsx" _
    )
    
For Each File In PickingFiles
    Call LoadPoFile(CStr(File))
Next
End Sub

Sub LoadSellerPicking(ByVal FileName As String)
'セラー分のピッキングファイル読み込み

Dim Mall As String, PickingFileName As String

'ピッキングシート名からモール記号をセット
Select Case True
    Case FileName Like "ピッキング*"
        Mall = "A"
    Case FileName Like "楽天*"
        Mall = "R"
    Case FileName Like "ヤフー*"
        Mall = "Y"
    Case Else
        Mall = "SP"
End Select

'ピッキングシートブックを開く、アクティブなまま使う
On Error Resume Next
    
    Workbooks.Open FileName:=PICKING_FOLDER & FileName
    If Err Then Exit Sub

On Error GoTo 0


'開いているピッキングシートから、手配依頼読込シートへデータコピー
With ThisWorkbook.Worksheets("セラー分")
    Dim WriteRow As Long, i As Long
    WriteRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
    
    For i = 3 To ActiveSheet.UsedRange.Rows.Count
        
        If Cells(i, 2).Interior.Color <> RGB(255, 255, 255) Then
            
            '背景白でない行を一旦コピー
            Range(Cells(i, 2), Cells(i, 5)).Copy
            '値で貼り付け
            .Cells(WriteRow, 2).PasteSpecial Paste:=xlPasteValues
            
            .Cells(WriteRow, 1).Value = Mall
            
            WriteRow = WriteRow + 1
        End If
    Next
End With

ActiveWorkbook.Close Savechanges:=False


End Sub
Sub LoadPoFile(ByVal FileName As String)
'Amazon卸のピッキングファイル読み込み

'ピッキングシートブックを開く、アクティブなまま使う
On Error Resume Next
    Workbooks.Open FileName:=PICKING_FOLDER & FileName
    If Err Then Exit Sub

On Error GoTo 0


'開いているピッキングシートから、手配依頼読込シートへデータコピー
With ThisWorkbook.Worksheets("卸分")
    Dim WriteRow As Long, i As Long
    WriteRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
    
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        
        If Cells(i, 2).Interior.Color <> RGB(255, 255, 255) Then
            
            'POとJANをコピー・貼り付け
            Range(Cells(i, 1), Cells(i, 2)).Copy
            .Cells(WriteRow, 2).PasteSpecial Paste:=xlPasteValues
            
            '商品名
            Cells(i, 5).Copy
            .Cells(WriteRow, 4).PasteSpecial Paste:=xlPasteValues
            
            '数量
            Cells(i, 9).Copy
            .Cells(WriteRow, 5).PasteSpecial Paste:=xlPasteValues
            
            .Cells(WriteRow, 1).Value = "V"
            
            WriteRow = WriteRow + 1
        End If
    Next
End With

ActiveWorkbook.Close Savechanges:=False

End Sub
