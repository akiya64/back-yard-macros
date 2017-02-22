Attribute VB_Name = "BuildSheets"
Option Explicit

Sub CreateSorterSheet(Mall As String)

'単体商品の振分け用シートを用意
Worksheets("振分用テンプレート").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = Mall & "_振分用"
    .PageSetup.LeftHeader = Format(Date, "M/dd") & " " & Mall
End With
Dim ForSorterSheet As Worksheet
Set ForSorterSheet = ActiveSheet

'セット商品の振分け用シートを用意
Worksheets("振分用テンプレート").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = Mall & "_振分用-セット"
    .PageSetup.LeftHeader = Format(Date, "M/dd") & " " & Mall & "-セット商品"
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
    '配列に行を格納
    Order(0) = CStr(Range("D" & i).Value) '6ケタ
    Order(1) = Range("E" & i).Value '商品名
    Order(2) = Range("F" & i).Value '数量
    Order(3) = CStr(Range("I" & i).Value) 'JAN
    Order(4) = Range("B" & i).Value 'お届け先名
    Order(5) = Range("Q" & i).Value '現在庫
    
    '転記先判定
    '7777始まりセットとセット構成商品、受注時コード7777***
    If Range("C" & i) Like "7777*" Then
       
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
       
           '数量、現在庫は右寄せ
            .Range("C" & k).HorizontalAlignment = xlRight
            .Range("F" & k).HorizontalAlignment = xlRight
       
            '棚番なしは、行に色を付ける。
            If OrderSheet.Range("H" & i).Value = "" Then
                     
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
    '受注時コードの7777は電算提出データに含めない。
    If Range("D" & i).Value Like "7777*" Then GoTo Continue

    '提出するコードの振替
    'SLIMSに投入するのはロケーション有りの6ケタのみ
    Dim OrderedCode As String, AddinResultCode As String, Code As String
    
    OrderedCode = CStr(Range("C" & i).Value)
    AddinResultCode = CStr(Range("J" & i).Value)
    
    If AddinResultCode = "" Then
        Code = OrderedCode
    Else
        Code = AddinResultCode
    End If
    
    '配列に提出ファイル1行分のデータを格納
    
    Order(0) = Range("A" & i).Value '注文番号
    Order(1) = Code '商品コード
    Order(2) = Range("E" & i).Value '商品名
    Order(3) = Range("F" & i).Value '数量
    Order(4) = Range("G" & i).Value '販売価格
    Order(5) = Range("G" & i).Value '現在庫
    Order(6) = Range("H" & i).Value '有効ロケーション
    
    
    '転記先判定  コードが入る列は書式：文字列として、先頭ゼロがカットされないように
    
    'ロケーションなし
    If Order(6) = "" Then
        
        NoEntryItemSheet.Range("C" & j).NumberFormatLocal = "@"
        NoEntryItemSheet.Range("B" & j & ":G" & j) = Order
    
        j = j + 1
    
    Else

        ForSlimsSheet.Range("C" & k).NumberFormatLocal = "@"
        ForSlimsSheet.Range("B" & k & ":H" & k) = Order
       
        k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

'罫線を引く
ForSlimsSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
NoEntryItemSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

'ブックを保存
ForSlimsBook.Close SaveChanges:=True
NoEntryItemBook.Close SaveChanges:=True

End Sub

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

Private Function PreparePickingBook(BookName As String) As Workbook
'ブック名を変えるために、所定の場所へ先にデータなしで保存する

ThisWorkbook.Worksheets("ピッキングシート提出用テンプレート").Copy

ActiveSheet.Name = BookName

'ファイル保存処理
'擬似的なTry-Catchでファイルを保存する
On Error Resume Next
    
    'Try 保存
  
    ActiveWorkbook.SaveAs FileName:="\\Server02\商品部\ネット販売関連\ピッキング\" & BookName & ".xlsx"

    'catch
    If Err Then
        Err.Clear
        MsgBox "ネット販売関連に繋がらないため、" & BookName & "は、デスクトップに保存します。"
        ActiveWorkbook.SaveAs FileName:="C:" & Environ("HOMEPATH") & "\Desktop\" & BookName & ".xlsx"
    End If
    
    'catch2
    If Err Then
        Err.Clear
        MsgBox "ファイルを保存できませんでした。手動で保存してください。"
    End If

Set PreparePickingBook = ActiveWorkbook

ThisWorkbook.ActiveSheet.Activate

End Function
