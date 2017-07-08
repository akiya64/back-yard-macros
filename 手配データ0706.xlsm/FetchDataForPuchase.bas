Attribute VB_Name = "FetchDataForPuchase"
Option Explicit

Sub CreateQuantitySheet()

'セラー分、卸分、手配数量入力シートを用意
Dim Sh As Variant
For Each Sh In Array(Worksheets("セラー分"), Worksheets("卸分"), Worksheets("手配数量入力シート"))
    Call PrepareSheet(Sh)
Next

'アマゾン・楽天・ヤフーの各棚なしピッキングシート、アマゾン卸の手配依頼読込
Call LoadPurchaseReq.LoadAllPicking

ThisWorkbook.SaveAs FileName:=ThisWorkbook.path & "\" & "手配データ" & Format(Date, "MMdd") & ".xlsm"

Worksheets("手配数量入力シート").Activate

'商品別に手配依頼数量を集計
Call SumPuchaseRequest

'発注に必要な情報をデータベース・Excelファイルから取得
Call FetchSyokonData
Call FetchExcellForPurchase

Call CalcPurchaseQuantity

Call FetchPickupFlag

'Excelで管理されている棚なし在庫のロケーションを取得
Call FetchExcellJanInventory

With Worksheets("手配数量入力シート")

    Dim EndRow As Long
    EndRow = Worksheets("手配数量入力シート").UsedRange.Rows.Count

    If .Buttons.Count > 0 Then
        .Buttons(1).Delete
    End If

    With .Buttons.Add( _
        Range("B" & EndRow).Left - 20, _
        Range("B" & EndRow).Top + 20, _
        200, _
        30 _
        )
        
        .OnAction = "BuildPurcahseData"
        .Characters.Text = "発注システム用データ出力"
        .Name = "BuidDataButton"
        
    End With

    .Range("A2").Activate

End With

ActiveWindow.ScrollColumn = 1
ActiveWindow.ScrollRow = 1

MsgBox Prompt:="手配数量入力シート、データ入力完了" & vbLf & "保留チェック、手配数量の修正を行ってください。", Buttons:=vbInformation

End Sub

Private Sub FetchSyokonData()
'商魂商品マスターのデータ取得
'実際に読みに行くテーブルは、受注詳細確認用に毎朝レプリケーションを作るServer3のEC課用マスタ

'接続のためのオブジェクトを定義、DB接続設定をセット
Dim DbCnn As New ADODB.Connection
Dim DbCmd  As New ADODB.Command
Dim DbRs As New ADODB.Recordset

DbCnn.ConnectionTimeout = 0
DbCnn.Open "PROVIDER=SQLOLEDB;Server=Server02;Database=ITOSQL_REP;UID=sa;PWD=;"
DbCmd.CommandTimeout = 180
Set DbCmd.ActiveConnection = DbCnn

'商品コードのレンジをセット、1セルずつSQL実行
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    Dim sql As String, Code As String
    Code = r.Value
        
    sql = "SELECT 商品コード, 取扱区分, ロット数, 仕入原価, 仕入先, 仕入先マスタ.仕入先略称, 仕入先マスタ.発注区分 " & _
          "FROM 商品マスタ JOIN 仕入先マスタ ON 商品マスタ.仕入先 = 仕入先マスタ.仕入先コード " & _
          "WHERE 商品コード = " & Code & "OR JANコード = '" & Code & "'"
    
    Set DbRs = DbCnn.Execute(sql)

    If Not DbRs.EOF Then
        Cells(r.Row, 3).Value = DbRs("ロット数")
        Cells(r.Row, 4).Value = DbRs("仕入先")
        Cells(r.Row, 5).Value = DbRs("仕入先略称")
        Cells(r.Row, 10).Value = DbRs("仕入原価")
        Cells(r.Row, 2).Value = GetKubunLabel(DbRs("取扱区分"))
        Cells(r.Row, 11).Value = DbRs("発注区分")
        
        'JAN受注分の商品コード置換、主にAmazon卸用
        If Len(Code) > 6 Then
            r.NumberFormatLocal = "@"
            r.Value = IIf(Len(DbRs("商品コード")) = 5, "0" & DbRs("商品コード"), DbRs("商品コード"))
        End If
    
    End If

Next

End Sub

Private Sub FetchExcellForPurchase()
'発注用商品情報ブックより仕入先・ロット・手配時備考の取得

Dim DataBook As Workbook, DataSheet As Worksheet, PurDataCodeRange As Range, PurDataJanRange As Range

Set DataSheet = FetchWorkBook("\\Server02\商品部\ネット販売関連\発注関連\発注用商品情報.xlsm").Worksheets("商品情報")

DataSheet.Activate

Set PurDataJanRange = DataSheet.Range(Cells(1, 1), Cells(DataSheet.UsedRange.Rows.Count, 1))
Set PurDataCodeRange = DataSheet.Range(Cells(1, 2), Cells(DataSheet.UsedRange.Rows.Count, 2))

ThisWorkbook.Activate
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange

    Dim Code As String, HitRow As Double
        
    Code = r.Value

    On Error Resume Next
        
        'エラー時に前回格納した値のままになるので明示的に初期化
        HitRow = 0
        
        HitRow = WorksheetFunction.Match(Code, PurDataCodeRange, 0)

        If Err Then
            Err.Clear
            HitRow = WorksheetFunction.Match(Code, PurDataJanRange, 0)
            
            If Err And IsEmpty(Cells(r.Row, 4).Value) Then '仕入先コードが商魂にもエクセルにもなければ、発注できないので注意書きを入れる
                Cells(r.Row, 2).Value = "発注用商品情報 該当JANなし"
            End If
        
        End If
            
    '手配時注意、メーカーロット、仕入先名
    Cells(r.Row, 2).Value = Cells(r.Row, 2).Value & DataSheet.Cells(HitRow, 35).Value '手配時注意
    Cells(r.Row, 12).Value = DataSheet.Cells(HitRow, 5).Value '発注用商品情報のロット数
    Cells(r.Row, 13).Value = DataSheet.Cells(HitRow, 4).Value '仕入先名
    
    '仕入先コード、原価、仕入先名は6ケタにない時のみ入れる
    If IsEmpty(Cells(r.Row, 4).Value) Then
    
        Cells(r.Row, 4).Value = DataSheet.Cells(HitRow, 32).Value '仕入先コード
        Cells(r.Row, 5).Value = DataSheet.Cells(HitRow, 4).Value '仕入先名
        Cells(r.Row, 10).Value = DataSheet.Cells(HitRow, 13).Value '原価

    End If

    On Error GoTo 0

Next

End Sub

Private Sub FetchExcellJanInventory()
'棚なし在庫表データの確認

Const NOLOCATION_INVENTRY_EXCELL As String = "\\server02\商品部\ネット販売関連\棚無在庫確認表.xlsm"
Const INVENTRY_SHEET As String = "棚無データ"

Workbooks.Open FileName:=NOLOCATION_INVENTRY_EXCELL, ReadOnly:=True

With ThisWorkbook.Worksheets("手配数量入力シート")
    Dim CodeRange As Range, r As Range
    Set CodeRange = .Range(.Cells(2, 7), .Cells(2, 7).End(xlDown))
End With

With Workbooks(Dir(NOLOCATION_INVENTRY_EXCELL)).Worksheets(INVENTRY_SHEET)
    Dim InventryRange As Range
    Set InventryRange = .Range(.Cells(1, 2), .Cells(1, 2).End(xlDown))
End With

For Each r In CodeRange

    Dim Code As String, HitRow As Double, Location As String, StockQuantity As Long
    
    Code = r.Value

    On Error Resume Next
    
    'エラー時に前回格納した値のままになるので明示的に初期化
    HitRow = 0
    
    HitRow = WorksheetFunction.Match(Code, InventryRange, 0)
    
    If Err = 0 Then
    
        StockQuantity = InventryRange.Cells(HitRow, 1).Offset(0, 1).Value
            
        If StockQuantity > 0 Then
            
            Location = InventryRange.Cells(HitRow, 1).Offset(0, 3).Value
            r.Offset(0, -5).Value = "棚無:" & StockQuantity & "場所:" & Location
        
        End If
            
    End If
    
    On Error GoTo 0

Next

Workbooks(Dir(NOLOCATION_INVENTRY_EXCELL)).Close SaveChanges:=False

End Sub

Private Sub CalcPurchaseQuantity()
'手配依頼数量から、ロット単位・発注単位で丸めた発注数数量を算出し、A列へ入れる。

Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    Dim i As Long
    i = r.Row
    
    Dim Rot As Double, Qty As Long, RequestQty As Double
    Rot = Cells(i, 12).Value
    
    If IsEmpty(Rot) Or Rot = 0 Then
        Rot = 1
    End If
    
    RequestQty = Cells(i, 9).Value
    
    'セイリング関数にてロット数の倍数で手配依頼数を丸める
    Qty = WorksheetFunction.Ceiling(RequestQty, Rot)

    Cells(i, 1).Value = Qty

    'ロットが1でない場合は、手配数量が修正されるため強調表示
    If Rot <> 1 Then
    
        With Union(Cells(i, 1), Cells(i, 9)).Interior
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    
    End If
    
Next

End Sub

Private Function GetKubunLabel(ByVal KubunCode As Variant) As String
'商品マスタでは区分は1～9の数字なので、表示名で置き換える。
'数字-区分名の組は名称マスタに格納されている。

Dim tmp As String

Select Case KubunCode
    Case 3
        tmp = "商魂:販売中止"
    Case 7
        tmp = "商魂:在庫廃番"
    Case 8
        tmp = "商魂:在庫処分"
    Case 9
        tmp = "商魂:メーカー廃番"
    Case Else
        tmp = ""
End Select

GetKubunLabel = tmp

End Function

Private Sub FetchPickupFlag()
'引取の仕入先について、仕入先リストのシートからVlookup関数にて引取の区分番号を取得する。
'引取以外は発注区分は 2

'仕入先コードのレンジをセット、1セルずつVlookupを行う
Dim CodeRange As Range, r As Range, VendorsRange As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown)).Offset(0, -3)

'仕入先コードでVlookupして探すレンジ
Set VendorsRange = ThisWorkbook.Worksheets("仕入先リスト").Range("A1").CurrentRegion

For Each r In CodeRange

    Dim VendorCode As String, DeliveryDiv As Integer
    
    VendorCode = r.Value
    
    On Error Resume Next
    
        DeliveryDiv = WorksheetFunction.VLookup(VendorCode, VendorsRange, 3, False)
        
        If Err Or DeliveryDiv = 0 Then
            DeliveryDiv = 2
        End If
        
    On Error GoTo 0
    
    Cells(r.Row, 11).Value = DeliveryDiv
    
Next

End Sub

