Attribute VB_Name = "BuildSheets"
Option Explicit

Sub シート作成()

OrderSheet.Activate

'InhouseData.TransferOrderSheet

BuildSheets.TransferPickingData
BuildSheets.TransferSorterSheet

End Sub

Private Sub TransferSorterSheet()

Worksheets("振分け用一覧シート").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!ショッピング"
Worksheets("振分け用一覧シート-セット").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!ショッピング セット"

Dim i As Long, k As Long, j As Long, Order(8) As Variant
'受注データシート行カウンタ
i = 2

'ネット用在庫シート行カウンタ
j = 2

'倉庫1シート行カウンタ
k = 2

Do
    '配列に行を格納
    Order(0) = CStr(Range("A" & i).Value) '注文番号
    Order(1) = Range("B" & i).Value 'お届け先名
    Order(2) = Range("D" & i).Value '社内コード
    Order(3) = Range("E" & i).Value '商品名
    Order(4) = Range("F" & i).Value '数量
    Order(5) = Range("L" & i).Value 'JAN
    Order(6) = Range("I" & i).Value '現在庫
    Order(7) = Range("J" & i).Value 'ロケーション
    Order(8) = Range("K" & i).Value '備考



    '転記先判定
       
    If Order(2) Like "7777*" Or Range("C" & i).Value = "Set" Then
       
       Worksheets("振分け用一覧シート-セット").Range("A" & j & ":I" & j).NumberFormatLocal = "@"
       'Worksheets("振分け用一覧シート-セット").Range("E" & j & ",G" & j).NumberFormatLocal = "G/標準"
       Worksheets("振分け用一覧シート-セット").Range("A" & j & ":I" & j) = Order
    
       j = j + 1
          
    'それ以外
    Else
    
       Worksheets("振分け用一覧シート").Range("A" & k & ":I" & k).NumberFormatLocal = "@"
       Worksheets("振分け用一覧シート").Range("A" & k & ":I" & k) = Order
       
       k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

End Sub

Private Sub TransferPickingData()

Dim i As Long, k As Long, j As Long, Order(4) As Variant
'受注データシート行カウンタ
i = 2

'ネット用在庫シート行カウンタ
j = 2

'倉庫1シート行カウンタ
k = 2

Do
    '配列に行を格納
    Order(0) = Range("D" & i).Value '社内コード
    Order(1) = Range("E" & i).Value '商品名
    Order(2) = Range("F" & i).Value '数量
    Order(3) = Range("G" & i).Value 'ヤフー販売価格
    Order(4) = Range("H" & i).Value '原価

    '転記先判定[
    'ロケーションなし
    If Range("J" & i).Value = "" Then
        
        If Not Order(0) Like "7777*" Then
           
           Worksheets("ネット用在庫").Range("B" & j).NumberFormatLocal = "@"
           Worksheets("ネット用在庫").Range("B" & j & ":F" & j) = Order
        
           j = j + 1
        
        End If
        
    'ロケーション有り
    Else
    
       Worksheets("倉庫1").Range("B" & k).NumberFormatLocal = "@"
       Worksheets("倉庫1").Range("B" & k & ":F" & k) = Order
       
       k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

End Sub
