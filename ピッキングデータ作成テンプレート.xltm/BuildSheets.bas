Attribute VB_Name = "BuildSheets"
Option Explicit
Sub CreateSorterSheet(MallName As String)

'単体商品の振分け用シートを用意
Worksheets("振分用テンプレート").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = MallName & "_振分用"
    .PageSetup.RightHeader = Format(Date, "M/dd") & " " & MallName
End With
Dim ForSorterSheet As Worksheet
Set ForSorterSheet = ActiveSheet

'セット商品の振分け用シートを用意
Worksheets("振分用テンプレート").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = MallName & "_振分用-セット"
    .PageSetup.RightHeader = Format(Date, "M/dd") & " " & MallName & "-セット商品"
End With
Dim ForSorterSetItemSheet As Worksheet
Set ForSorterSetItemSheet = ActiveSheet

'アクティブなシートはコピーしたシートから受注シートに変えておく
OrderSheet.Activate

Dim i As Long, k As Long, j As Long, Order(6) As Variant
'受注データシート行カウンタ
i = 2

'振分け用シート行カウンタ
j = 2

'振分け用セットシート行カウンタ
k = 2

Do

    '引数で渡されたモール以外は飛ばす
    If Not Range("F" & i).Value Like (MallName & "*") Then GoTo Continue

    '配列に行を格納
    Order(0) = CStr(Range("B" & i).Value) '受注時コード
    Order(1) = Range("C" & i).Value '商品名
    Order(2) = Range("D" & i).Value '受注数量
    Order(3) = CStr(Range("L" & i).Value) 'JAN
    Order(4) = Range("G" & i).Value 'お届け先名
    Order(5) = Range("N" & i).Value '現在庫
    
    
    '現在庫が取得できてないときは、印刷レイアウトの関係のため空白1文字入れておく
    If Order(5) = "" Then Order(5) = " "
    
    '転記先判定
    '7777始まりセットとセット構成商品、受注時コード7777***
    If Range("B" & i) Like "7777*" Then
       
       Order(0) = Range("I" & i).Value
       
        With ForSorterSetItemSheet
            
            .Range("A" & j & ":F" & j).NumberFormatLocal = "@"
            .Range("A" & j & ":F" & j) = Order
            
            '数量、現在庫は右寄せ
            .Range("C" & j).HorizontalAlignment = xlRight
            .Range("F" & j).HorizontalAlignment = xlRight
        
        End With
        
        j = j + 1
          
    Else
    
        With ForSorterSheet
        
            .Range("A" & k & ":F" & k).NumberFormatLocal = "@"
            .Range("A" & k & ":F" & k) = Order
       
           '数量、現在庫は右寄せ、JANは中央
            .Range("C" & k).HorizontalAlignment = xlRight
            .Range("D" & k).HorizontalAlignment = xlCenter
            .Range("F" & k).HorizontalAlignment = xlRight
       
            '棚番なしは、行に色を付ける。
            If OrderSheet.Range("K" & i).Value = "" Then
                     
                With .Range("A" & k & ":F" & k).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
             
            End If
        
        End With
        
        k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))


Call Sort.振分用シート_ソート(ForSorterSheet)

ForSorterSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
ForSorterSetItemSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous


'終了処理 Sheet内容確定

'念のため幅を再指定
Call AdjustWidth(ForSorterSheet)
Call AdjustWidth(ForSorterSetItemSheet)

ForSorterSheet.Protect
ForSorterSetItemSheet.Protect

End Sub

Sub OutputPickingData(MallName As String)

'モール名、クロスモール内では「アマゾン」でなく「Amazon」
Dim MallIdentify As String
If MallName = "アマゾン" Then
    MallIdentify = "Amazon"
Else
    MallIdentify = MallName
End If

'モール毎の受注件数がゼロ件ならファイル生成しない。
If WorksheetFunction.CountIf(OrderSheet.Range("F:F"), CStr(MallIdentify) & "*") = 0 Then GoTo Continue


'提出用ファイルを用意
'100番/200番棚有り -2-3、電算室提出
Dim ForSlimsBook As Workbook, ForSlimsSheet As Worksheet
Set ForSlimsBook = PreparePickingBook(MallName & "Pシート" & Format(Date, "MMdd") & "-2-3")
Set ForSlimsSheet = ForSlimsBook.Worksheets(1)

'登録無し、棚無し -a
Dim NoEntryItemBook As Workbook, NoEntryItemSheet As Worksheet
Set NoEntryItemBook = PreparePickingBook(MallName & "Pシート" & Format(Date, "MMdd") & "-a")
Set NoEntryItemSheet = NoEntryItemBook.Worksheets(1)

OrderSheet.Activate

Dim i As Long, k As Long, j As Long, Order(6) As Variant
'受注データシート行カウンタ
i = 2

'棚無しシート行カウンタ
j = 2

'100番シート行カウンタ
k = 2

'用意したブックへ1行ずつコピー
Do

    '引数で渡されたモール以外は飛ばす
    If Not Range("F" & i).Value Like (MallIdentify & "*") Then GoTo Continue
    
    '受注時コードの7777は電算提出データに含めない。
    If Range("I" & i).Value Like "7777*" Then GoTo Continue

    '提出するコードの振替
    'SLIMSに投入するのはロケーション有りの6ケタのみ
    Dim OrderedCode As String, AddinResultCode As String, Code As String
    
    OrderedCode = CStr(Range("B" & i).Value)
    AddinResultCode = CStr(Range("M" & i).Value)
    
    If AddinResultCode = "" Then
        Code = OrderedCode
    Else
        Code = AddinResultCode
    End If
    
    '配列に提出ファイル1行分のデータを格納
    'アマゾンのみ、電算室処理でアマゾン注文番号を判定している、連番不可
    If MallIdentify = "Amazon" Then
        Order(0) = CStr(Range("H" & i).Value) 'モール側採番の注文番号
    Else
        Order(0) = CStr(Range("A" & i).Value) 'クロスモール採番の連番
    End If
    
    Order(1) = CStr(Code) '商品コード
    Order(2) = Range("C" & i).Value '商品名
    Order(3) = Range("J" & i).Value '数量
    Order(4) = Range("E" & i).Value '販売価格
    Order(5) = Range("N" & i).Value '現在庫
    Order(6) = Range("K" & i).Value '有効ロケーション
    
    '転記先判定  コードが入る列は書式：文字列として、先頭ゼロがカットされないように
    
    'ロケーションなし
    If Order(6) = "" Then
        
        NoEntryItemSheet.Range("B" & j & ":C" & j).NumberFormatLocal = "@"
        NoEntryItemSheet.Range("B" & j & ":H" & j) = Order
    
        j = j + 1
    
    Else

        ForSlimsSheet.Range("B" & k & ":C" & k).NumberFormatLocal = "@"
        ForSlimsSheet.Range("B" & k & ":H" & k) = Order
       
        k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

'Pシートのブック保存処理
'Amazonのみ送料列が必要、送料列 0円 で埋める
With ForSlimsSheet
    .Activate

    If MallName = "アマゾン" Then
        .Columns("G").Insert
        .Range("G1").Value = "送料"
        .Range(Cells(2, 7), Cells(.UsedRange.Rows.Count, 7)).Value = 0
    End If

    '罫線引いて保存
    .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
    
End With
ForSlimsBook.Close SaveChanges:=True
    
With NoEntryItemSheet
    .Activate
    
    If MallName = "アマゾン" Then
        .Columns("G").Insert
        .Range("G1").Value = "送料"
        .Range(Cells(2, 7), Cells(.UsedRange.Rows.Count, 7)).Value = 0
    End If
    
    '罫線引いて保存
    .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
    
End With
NoEntryItemBook.Close SaveChanges:=True

End Sub

Private Function PreparePickingBook(BookName As String) As Workbook
'ブック名を変えるために、所定の場所へ先にデータなしで保存する

'引数の名前で新規ブックを作成する
ThisWorkbook.Worksheets("ピッキングシート提出用テンプレート").Copy
ActiveSheet.Name = BookName

'一旦新規作成ブックを保存することでブック名を変更する
'新規作成ファイルの保存時はファイルフォーマットを明示する必要な模様
    Dim SavePath As String
    Const PICKING_FOLDER As String = "\\Server02\商品部\ネット販売関連\ピッキング\"
    
    If Dir(PICKING_FOLDER, vbDirectory) <> "" Then
        '既に本日ファイルがあれば、時刻付けて保存
        If Dir(PICKING_FOLDER & BookName & ".xlsx") = "" Then
            SavePath = PICKING_FOLDER & BookName
        Else
            SavePath = PICKING_FOLDER & Format(Time, "hhmm") & BookName
        End If
        
            ActiveWorkbook.SaveAs Filename:=SavePath, FileFormat:=xlWorkbookDefault
    
    Else
        
        Dim DeskTopPath As String
        If Dir(DeskTopPath & BookName & ".xlsx") = "" Then
            DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & BookName
        Else
            DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & Format(Time, "hhmm") & BookName
        End If
        
        MsgBox "ネット販売関連に繋がらないため、" & BookName & "をデスクトップに保存します。"
            
        ActiveWorkbook.SaveAs Filename:=DeskTopPath, FileFormat:=xlWorkbookDefault
    
    End If

Set PreparePickingBook = ActiveWorkbook

ThisWorkbook.ActiveSheet.Activate

End Function

Private Sub AdjustWidth(TargetSheet As Worksheet)
'列幅 調整時にアラートが出るのを抑止
Application.DisplayAlerts = False

Dim WidthArr As Variant
WidthArr = Array(14.75, 84.13, 4.25, 15.88, 14.88, 6.63)

TargetSheet.Activate

Dim k As Long
For k = 0 To 5
    TargetSheet.Columns(k + 1).ColumnWidth = WidthArr(k)
Next

Application.DisplayAlerts = True

End Sub

Sub Test_buildsheet()
    Call BuildSheets.OutputPickingData("アマゾン")
End Sub

