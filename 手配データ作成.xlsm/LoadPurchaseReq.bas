Attribute VB_Name = "LoadPurchaseReq"
Option Explicit

Sub LoadAllPicking(Optional ByRef TargetFolder As String)
'手配依頼チェック済のピッキングファイルを一括して読込
'手配依頼として背景色が変えてある行をコピーします。

Dim PickingFilePaths() As String
PickingFilePaths = SearchPickingFiles(TargetFolder)

'セラー分ピッキングファイル読み込み
Dim Path As Variant
For Each Path In Array(PickingFilePaths(0), PickingFilePaths(1), PickingFilePaths(2))
    Call LoadSellerPicking(Path)
Next

'卸分ピッキングファイル読み込み
For Each Path In Array(PickingFilePaths(3), PickingFilePaths(4))
    Call LoadPoFile(Path)
Next

Call ApendSpToPurchseReq
Call VerifySyokonRegist

End Sub

Private Sub LoadSellerPicking(ByVal PickingFilePath As String)
'セラー分のピッキングファイル読み込み

Dim Mall As String

'ピッキングシート名からモール記号をセット
Select Case True
    Case PickingFilePath Like "*ピッキングシート*"
        Mall = "A"
    Case PickingFilePath Like "*楽天*"
        Mall = "R"
    Case PickingFilePath Like "*ヤフー*"
        Mall = "Y"
    Case Else
        Mall = "SP"
End Select

'ピッキングシートブックを開く
On Error Resume Next
    
    Workbooks.Open FileName:=PickingFilePath
    If Err Then Exit Sub

On Error GoTo 0

Dim NoLocationSheet As Worksheet
Set NoLocationSheet = ActiveSheet


'開いているピッキングシートから、背景色を判定しつつ1行ずつデータコピー
Dim WriteRow As Long, i As Long
WriteRow = IIf(PurchaseReqSeller.Range("A2").Value = "", 2, PurchaseReqSeller.Range("A1").End(xlDown).Row + 1)

For i = 3 To NoLocationSheet.Range("A1").SpecialCells(xlLastCell).Row
    
    If NoLocationSheet.Cells(i, 2).Interior.Color <> RGB(255, 255, 255) Then
        
        'ピッキング-aの背景白でない行を一旦コピー
        NoLocationSheet.Range(Cells(i, 2), Cells(i, 5)).Copy
        '値で貼り付け
        PurchaseReqSeller.Cells(WriteRow, 2).PasteSpecial Paste:=xlPasteValues
        
        PurchaseReqSeller.Cells(WriteRow, 1).Value = Mall
        
        WriteRow = WriteRow + 1
        
    End If

Next

ActiveWorkbook.Close SaveChanges:=False

End Sub
Private Sub LoadPoFile(ByVal PickingFilePath As String)
'Amazon卸のピッキングファイル読み込み

'ピッキングシートブックを開く
On Error Resume Next

    Workbooks.Open FileName:=PickingFilePath
    If Err Then Exit Sub

On Error GoTo 0

Dim NoLocationSheet As Worksheet
Set NoLocationSheet = ActiveSheet

'開いているピッキングシートから、手配依頼読込シートへデータコピー

Dim WriteRow As Long, i As Long
WriteRow = IIf(PurchaseReqWholesall.Range("A2").Value = "", 2, PurchaseReqWholesall.Range("A1").End(xlDown).Row + 1)

For i = 2 To ActiveSheet.Range("A1").SpecialCells(xlLastCell).Row
    
    If Cells(i, 2).Interior.Color <> RGB(255, 255, 255) Then
        
        'POとJANをコピー・貼り付け
        NoLocationSheet.Range(Cells(i, 1), Cells(i, 2)).Copy
        PurchaseReqWholesall.Cells(WriteRow, 2).PasteSpecial Paste:=xlPasteValues
        
        '商品名
        NoLocationSheet.Cells(i, 5).Copy
        PurchaseReqWholesall.Cells(WriteRow, 4).PasteSpecial Paste:=xlPasteValues
        
        '数量
        NoLocationSheet.Cells(i, 9).Copy
        PurchaseReqWholesall.Cells(WriteRow, 5).PasteSpecial Paste:=xlPasteValues
        
        PurchaseReqWholesall.Cells(WriteRow, 1).Value = "V"
        
        WriteRow = WriteRow + 1
    End If
Next

ActiveWorkbook.Close SaveChanges:=False

End Sub

Private Sub ApendSpToPurchseReq()
'手入力分を卸分シート、セラー分シートへ振り分けてコピー
'180番、187番で商品別に数量を合算するため

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

'接続のためのオブジェクトを定義、DB接続設定をセット
Dim DbCnn As New ADODB.Connection
Dim DbCmd  As New ADODB.Command
Dim DbRs As New ADODB.Recordset

DbCnn.ConnectionTimeout = 0
DbCnn.Open "PROVIDER=SQLOLEDB;Server=Server02;Database=ITOSQL_REP;UID=sa;PWD=;"
DbCmd.CommandTimeout = 180
Set DbCmd.ActiveConnection = DbCnn

'セラー分は手入力だけJANだけど登録有りがないかを調べればよいので、最終行から上へ調べていく
With PurchaseReqSeller
    Dim EndRow As Long
    EndRow = PurchaseReqSeller.Range("A1").End(xlDown).Row
    
    Dim i As Long
    For i = EndRow To 2 Step -1
        
        If .Cells(i, 1).Value <> "SP" Then Exit For
        
        Dim Code As String
        Code = .Cells(i, 3).Value
        
        If Not .Cells(i, 3).Value Like String(13, "#") Then GoTo Continue
    
        'クエリで6ケタ取得
        Dim Sql As String
            
        Sql = "SELECT 商品コード FROM 商品マスタ WHERE JANコード = '" & Code & "'"
        
        Set DbRs = DbCnn.Execute(Sql)
    
        If Not DbRs.EOF Then
            .Cells(i, 3).NumberFormatLocal = "@"
            .Cells(i, 3).Value = IIf(Len(DbRs("商品コード")) = 5, "0" & CStr(DbRs("商品コード")), CStr(DbRs("商品コード")))
        End If
    
Continue:
    Next

End With

'卸分は、全部調べる
With PurchaseReqWholesall
    EndRow = Range("A1").End(xlDown).Row
    
    For i = 2 To EndRow
        
        If .Cells(i, 3).Value = "" Then Exit For
        
        Code = .Cells(i, 3).Value
        
        If Not .Cells(i, 3).Value Like String(13, "#") Then GoTo Continue2

        Sql = "SELECT 商品コード FROM 商品マスタ WHERE JANコード = '" & Code & "'"
        
        Set DbRs = DbCnn.Execute(Sql)
    
        If Not DbRs.EOF Then
            .Cells(i, 3).NumberFormatLocal = "@"
            .Cells(i, 3).Value = IIf(Len(DbRs("商品コード")) = 5, "0" & CStr(DbRs("商品コード")), CStr(DbRs("商品コード")))
        End If
    
Continue2:
    Next

End With

End Sub

Private Function SearchPickingFiles(Optional ByRef FolderPath As String) As String()
'フォルダ指定を元に、ピッキングファイルのパスを配列で返します。

'PickingFiles(0) : Amazonセラー
'PickingFiles(1) : 楽天
'PickingFiles(2) : ヤフー
'PickingFiles(3) : Amazon卸
'PickingFiles(4) : Amazon卸アウトドアカテゴリ

Dim Fso As New FileSystemObject, PickingFolder As Folder, File As File

Set PickingFolder = Fso.GetFolder(FolderPath)
Dim PickingFiles(4) As String

For Each File In PickingFolder.Files

    Select Case True
        Case File.Name Like "ピッキングシート*-a*"
            PickingFiles(0) = FolderPath & "\" & File.Name
            
        Case File.Name Like "楽天*-a*"
            PickingFiles(1) = FolderPath & "\" & File.Name
            
        Case File.Name Like "ヤフー*-a*"
            PickingFiles(2) = FolderPath & "\" & File.Name
        
        Case File.Name Like "アマゾン棚なし####.xlsx"
            PickingFiles(3) = FolderPath & "\" & File.Name
            
        Case File.Name Like "アマゾン棚なし*-outdoor*"
            PickingFiles(4) = FolderPath & "\" & File.Name
    
    End Select
        
Next

SearchPickingFiles = PickingFiles

End Function
Private Sub TestGetPickingFiles()

Dim Files As Variant, File As Variant

Files = GetPickingFiles(PICKING_FOLDER)

For Each File In Files

    Debug.Print File

    If File Like "*ピッキングシート*" Then
        Debug.Print "Amazon OK"
    ElseIf File Like "*楽天*" Then
        Debug.Print "楽天 OK"
    ElseIf File Like "*ヤフー*" Then
        Debug.Print "ヤフー OK"
    End If
Next

End Sub
