Attribute VB_Name = "ConnectDB"
Sub Make_List(Optional ByVal arg As Boolean)
'DBへ接続して、商品マスタ・在庫マスタからロケーション・現在庫を取得
'在庫マスタがない場合は、商品マスタより社内コードとJANのみ取得
'ベース作成：商品部


'SQL系変数
Dim DB_Cnn As New ADODB.Connection
Dim DB_Cmd  As New ADODB.Command
Dim DB_Rs As New ADODB.Recordset
Dim SQL_W1 As String

'状態確認変数
Dim Target_RowEnd As Integer
Dim Loop_Count As Integer
Dim A As Integer
Dim Target_Code As String
Dim Loc_Text As String

'定数セット
Const Target_RowBase = 2
Const Target_ColBase = 9
Const Output_RowBase = 2
Const OutPut_ColBase = 12

'シートタイトル用変数
Dim S_HEAD(3)

'シートタイトルセット
S_HEAD(0) = "JANコード"
S_HEAD(1) = "商品コード"
S_HEAD(2) = "現在庫数"
S_HEAD(3) = "ロケーション"


'Workbookが開いているか確認
A = 0
For Each wn In Workbooks
A = A + 1
Next
If A = 0 Then
MsgBox ("シートを開いてください。")
End
End If

'SQL Server接続
DB_Cnn.ConnectionTimeout = 0
DB_Cnn.Open "PROVIDER=SQLOLEDB;Server=Server02;Database=ITOSQL_REP;UID=sa;PWD=;"
DB_Cmd.CommandTimeout = 180
Set DB_Cmd.ActiveConnection = DB_Cnn


'---処理開始---
'ヘッダーセット
For Loop_Count = 0 To 3
    Cells(1, 12 + Loop_Count).Select
    Cells(1, 12 + Loop_Count).Value = S_HEAD(Loop_Count)
Next Loop_Count

'最終Row取得
Cells(2, 1).Select
Range(Selection, Selection.End(xlDown)).Select
Target_RowEnd = ActiveSheet.Cells.SpecialCells(xlLastCell).Row

'画面更新、再計算抑止
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'メインループ
Cells(Target_RowBase, Target_ColBase).Select
For Loop_Count = Target_RowBase To Target_RowEnd
    Target_Code = Cells(Loop_Count, Target_ColBase).Value
    
    'コード判別（インストア・JAN）
    If Len(Target_Code) <= 6 Then
        SQL_W1 = "SELECT 商品マスタ.商品コード, 商品マスタ.JANコード, Sum(在庫マスタ.現在庫数) AS 現在庫数計 "
        SQL_W1 = SQL_W1 & "FROM 商品マスタ INNER JOIN 在庫マスタ ON 商品マスタ.商品コード = 在庫マスタ.商品コード "
        SQL_W1 = SQL_W1 & "GROUP BY 商品マスタ.商品コード, 商品マスタ.JANコード "
        SQL_W1 = SQL_W1 & "HAVING (((商品マスタ.商品コード)=" & Target_Code & "));"
    Else
        SQL_W1 = "SELECT 商品マスタ.商品コード, 商品マスタ.JANコード, Sum(在庫マスタ.現在庫数) AS 現在庫数計 "
        SQL_W1 = SQL_W1 & "FROM 商品マスタ INNER JOIN 在庫マスタ ON 商品マスタ.商品コード = 在庫マスタ.商品コード "
        SQL_W1 = SQL_W1 & "GROUP BY 商品マスタ.商品コード, 商品マスタ.JANコード "
        SQL_W1 = SQL_W1 & "HAVING (((商品マスタ.JANコード)='" & Target_Code & "'));"
    End If
    
    Set DB_Rs = DB_Cnn.Execute(SQL_W1)

    If Not DB_Rs.EOF Then
        Cells(Loop_Count, OutPut_ColBase).Value = DB_Rs("JANコード")
        Cells(Loop_Count, OutPut_ColBase + 1).NumberFormatLocal = "@"
        Cells(Loop_Count, OutPut_ColBase + 1).Value = Format(DB_Rs("商品コード"), "000000")
        Cells(Loop_Count, OutPut_ColBase + 2).Value = DB_Rs("現在庫数計")
        
        'ロケーション情報の取得
        SQL_W1 = "SELECT 在庫マスタ.商品コード,"
        SQL_W1 = SQL_W1 & "[在庫マスタ].[階]+'-'+[在庫マスタ].[通路]+'-'+[在庫マスタ].[棚番]+'-'+[在庫マスタ].[段]+'-'+[在庫マスタ].[順] AS ロケーション "
        SQL_W1 = SQL_W1 & "FROM 在庫マスタ WHERE (在庫マスタ.商品コード=" & DB_Rs("商品コード") & ");"
        
        Set DB_Rs = DB_Cnn.Execute(SQL_W1)
        
        Loc_Text = ""
        Do While Not DB_Rs.EOF
            Loc_Text = Loc_Text & "[" & DB_Rs("ロケーション") & "]"
            DB_Rs.MoveNext
        Loop
        
        Cells(Loop_Count, OutPut_ColBase + 3).Value = Loc_Text
        
    Else
        '在庫マスターに登録がない場合、商品マスタから商品コードとJANのみ取得する
        
        'コード判別（インストア・JAN）-> WHERE句セット DBでコードは数値型、JANはテキスト型
        Dim Clause_WHERE As String
        Clause_WHERE = IIf(Len(Target_Code) <= 6, "商品マスタ.商品コード = " & Target_Code, "商品マスタ.JANコード = '" & Target_Code & "'")
    
        SQL_W1 = "SELECT 商品マスタ.商品コード, 商品マスタ.JANコード "
        SQL_W1 = SQL_W1 & "FROM 商品マスタ "
        SQL_W1 = SQL_W1 & "WHERE " & Clause_WHERE
        
        'SQL実行して出力
        Set DB_Rs = DB_Cnn.Execute(SQL_W1)
        
        If Not DB_Rs.EOF Then
            Cells(Loop_Count, OutPut_ColBase).Value = DB_Rs("JANコード")
            
            Cells(Loop_Count, OutPut_ColBase + 1).NumberFormatLocal = "@"
            Cells(Loop_Count, OutPut_ColBase + 1).Value = Format(DB_Rs("商品コード"), "000000")
        End If
        
    End If
Next Loop_Count
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    'Cells.Select
    'Cells.EntireColumn.AutoFit
    Range("A1").Select

DB_Cnn.Close

Set DB_Rs = Nothing
Set DB_Cnn = Nothing
Set DB_Cmd = Nothing


End Sub


