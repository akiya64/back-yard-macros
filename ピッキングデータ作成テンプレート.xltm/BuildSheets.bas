Attribute VB_Name = "BuildSheets"
Option Explicit

Sub 電算提出_振分けシート作成()

Const OUTPUT_FOLDER As String = "\\Server02\商品部\ネット販売関連\ピッキング\"

OrderSheet.Activate

If InStr(Range("A1").Value, "アドイン指定") > 0 Then
    MsgBox "アドインを実行して下さい。"
End If

SyokonData.TransferOrderSheet

'振分け用シートの列幅固定のための保護を解除
ForSorterSheet.Unprotect
ForSorterSetItemSheet.Unprotect

BuildSheets.TransferPickingData
BuildSheets.TransferSorterSheet

Dim i As Long

'罫線引く
For i = 2 To 5

    With Worksheets(i).Range("A1").CurrentRegion.Borders
        .LineStyle = xlContinuous
    End With

Next

'振分け用シートを表示して、ソート
Worksheets("振分け用一覧シート").Activate
Sort.振分用シート_ソート

ForSorterSheet.Protect AllowSorting:=True, AllowFiltering:=True
ForSorterSetItemSheet.Protect AllowSorting:=True, AllowFiltering:=True

'電算提出用保存 100番 棚有り
Sheets("100番").Copy
ActiveWorkbook.SaveAs filename:=OUTPUT_FOLDER & "ヤフーPシート" & Format(Date, "MMdd") & "-2-3.xlsx"
ActiveWorkbook.Close

'電算提出用保存 棚無し
Sheets("棚無し").Copy
ActiveWorkbook.SaveAs filename:=OUTPUT_FOLDER & "ヤフーPシート" & Format(Date, "MMdd") & "-a.xlsx"
ActiveWorkbook.Close

'このファイルを保存
Application.DisplayAlerts = False
ThisWorkbook.SaveAs filename:="\\MOS10\Users\mos10\Desktop\ヤフー\ピッキング生成用過去ファイル\" & "ヤフー提出・振分け用" & Format(Date, "MMdd") & ".xlsx"

End Sub

Private Sub TransferSorterSheet()

Worksheets("振分け用一覧シート").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!ショッピング"
Worksheets("振分け用一覧シート-セット").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!ショッピング セット"

Dim i As Long, k As Long, j As Long, Order(8) As Variant
'受注データシート行カウンタ
i = 2

'振分け用シート行カウンタ
j = 2

'振分け用セットシート行カウンタ
k = 2

Do
    '配列に行を格納
    Order(0) = CStr(Range("A" & i).Value) '注文番号
    Order(1) = Range("B" & i).Value 'お届け先名
    Order(2) = CStr(Range("D" & i).Value) '6ケタ
    Order(3) = Range("E" & i).Value '商品名
    Order(4) = Range("F" & i).Value '数量
    Order(5) = CStr(Range("L" & i).Value) 'JAN
    Order(6) = Range("I" & i).Value '現在庫
    Order(7) = Range("K" & i).Value '備考
    Order(8) = Range("J" & i).Value 'ロケーション
    
    '転記先判定
    '7777始まりセットとセット内容品
    If Order(2) Like "7777*" Or Range("C" & i).Value = "Set" Then
       
        With Worksheets("振分け用一覧シート-セット")
            
            .Range("A" & j & ":I" & j).NumberFormatLocal = "@"
            .Range("A" & j & ":I" & j) = Order
            
            '数量、現在庫は右寄せ
            .Range("E" & j).HorizontalAlignment = xlRight
            .Range("G" & j).HorizontalAlignment = xlRight
        
        End With
        
        j = j + 1
          
    'それ以外
    Else
        With Worksheets("振分け用一覧シート")
        
            .Range("A" & k & ":I" & k).NumberFormatLocal = "@"
            .Range("A" & k & ":I" & k) = Order
       
           '数量、現在庫は右寄せ
            .Range("E" & k).HorizontalAlignment = xlRight
            .Range("G" & k).HorizontalAlignment = xlRight
        
       
        '棚番なしは、行に色を付ける。
            If Order(8) = "" Then
                     
                With .Range("A" & k & ":I" & k).Interior
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

Worksheets("振分け用一覧シート").Range("A1").CurrentRegion.Font.Size = 9
Worksheets("振分け用一覧シート-セット").Range("A1").CurrentRegion.Font.Size = 9

End Sub

Private Sub TransferPickingData()

Dim i As Long, k As Long, j As Long, Order(6) As Variant
'受注データシート行カウンタ
i = 2

'棚無しシート行カウンタ
j = 2

'100番シート行カウンタ
k = 2

Do
    '配列に行を格納
    Order(0) = Range("A" & i).Value '注文番号
    Order(1) = CStr(Range("D" & i).Value) '6ケタ
    Order(2) = Range("E" & i).Value '商品名
    Order(3) = Range("F" & i).Value '数量
    Order(4) = Range("G" & i).Value 'ヤフー販売価格
    Order(5) = Range("I" & i).Value '現在庫
    Order(6) = Range("J" & i).Value '棚番
    
    '転記先判定
    'ロケーションなし
    If Order(6) = "" Then
        
        If Not Order(0) Like "7777*" Then
           
           'コードが入る列は文字列として、先頭ゼロがカットされないように
           Worksheets("棚無し").Range("C" & j).NumberFormatLocal = "@"
           Worksheets("棚無し").Range("B" & j & ":G" & j) = Order
        
           j = j + 1
        
        End If
        
    'ロケーション有り
    Else
    
       'コードが入る列は文字列として、先頭ゼロがカットされないように
       Worksheets("100番").Range("C" & k).NumberFormatLocal = "@"
       Worksheets("100番").Range("B" & k & ":H" & k) = Order
       
       k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

End Sub
