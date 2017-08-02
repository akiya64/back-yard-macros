Attribute VB_Name = "LoadPurchaseReq"
Option Explicit

Const PICKING_FOLDER As String = "\\server02\商品部\ネット販売関連\ピッキング\"

Sub LoadAllPicking()
'手配依頼チェック済のピッキングファイルを一括して読込
'手配依頼として背景色が変えてある行をコピーします。

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

Call ApendSpToPurchseReq
Call VerifySyokonRegist

End Sub

Private Sub LoadSellerPicking(ByVal FileName As String)
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
    WriteRow = IIf(.Range("A2").Value = "", 2, .Range("A1").End(xlDown).Row + 1)
    
    For i = 3 To ActiveSheet.Range("A1").SpecialCells(xlLastCell).Row
        
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

ActiveWorkbook.Close SaveChanges:=False

End Sub
Private Sub LoadPoFile(ByVal FileName As String)
'Amazon卸のピッキングファイル読み込み

'ピッキングシートブックを開く、アクティブなまま使う
On Error Resume Next
    Workbooks.Open FileName:=PICKING_FOLDER & FileName
    If Err Then Exit Sub

On Error GoTo 0


'開いているピッキングシートから、手配依頼読込シートへデータコピー
With ThisWorkbook.Worksheets("卸分")
    Dim WriteRow As Long, i As Long
    WriteRow = IIf(.Range("A2").Value = "", 2, .Range("A1").End(xlDown).Row + 1)
    
    For i = 2 To ActiveSheet.Range("A1").SpecialCells(xlLastCell).Row
        
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

ActiveWorkbook.Close SaveChanges:=False

End Sub

Private Sub ApendSpToPurchseReq()
'手入力分を卸分シート、セラー分シートへ振り分けてコピー
'180番、187番で商品別に数量を合算するた

With Worksheets("手入力分")

    If IsEmpty(.Range("B2").Value) Then
        Exit Sub
    Else
        Dim CodeRange As Range
        Set CodeRange = .Range(.Cells(2, 2), .Cells(1, 2).End(xlDown))
    End If
    
End With

Dim r As Range, MallTicker As String

For Each r In CodeRange
    MallTicker = r.Offset(0, -1).Value
    
    If MallTicker Like "*[V|v]*" Then
    
        With Worksheets("卸分")
            .Range("A1").End(xlDown).Offset(1, 0).Value = "V"
            .Range("C1").End(xlDown).Offset(1, 0).NumberFormatLocal = "@"
            .Range("C1").End(xlDown).Offset(1, 0).Resize(1, 3).Value = r.Resize(1, 3).Value
        End With
        
    Else
    
        With Worksheets("セラー分")
            .Range("A1").End(xlDown).Offset(1, 0).Value = "SP"
            .Range("C1").End(xlDown).Offset(1, 0).NumberFormatLocal = "@"
            .Range("C1").End(xlDown).Offset(1, 0).Resize(1, 3).Value = r.Resize(1, 3).Value
        End With
    
    End If

Next

End Sub

Private Sub VerifySyokonRegist()
'DBに接続して最終行から上へ調べていく

'接続のためのオブジェクトを定義、DB接続設定をセット
Dim DbCnn As New ADODB.Connection
Dim DbCmd  As New ADODB.Command
Dim DbRs As New ADODB.Recordset

DbCnn.ConnectionTimeout = 0
DbCnn.Open "PROVIDER=SQLOLEDB;Server=Server02;Database=ITOSQL_REP;UID=sa;PWD=;"
DbCmd.CommandTimeout = 180
Set DbCmd.ActiveConnection = DbCnn

With PurchaseReqSeller
    Dim EndRow As Long
    EndRow = PurchaseReqSeller.Range("A1").End(xlDown).Row
    
    Dim i As Long
    For i = EndRow To 2 Step -1
        
        If .Cells(i, 1).Value <> "SP" Then Exit Sub
        
        Dim Code As String
        Code = .Cells(i, 3).Value
        
        If Not .Cells(i, 3).Value Like String(13, "#") Then GoTo Continue
    
        'クエリで6ケタ取得
        Dim Sql As String
            
        Sql = "SELECT 商品コード FROM 商品マスタ WHERE JANコード = '" & Code & "'"
        
        Set DbRs = DbCnn.Execute(Sql)
    
        If Not DbRs.EOF Then
            .Cells(i, 3).Value = CStr(DbRs("商品コード"))
        End If
    
Continue:
    Next

End With

End Sub


