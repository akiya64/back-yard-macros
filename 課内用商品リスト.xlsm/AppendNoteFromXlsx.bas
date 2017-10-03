Attribute VB_Name = "AppendNoteFromXlsx"
Sub InsertCaution()

'元のエクセルブック名、元シートの範囲は毎回指定のこと
'イミディエイトで、Workbooks(1).nameでワークブック名が確認できる。
Set Rng = Workbooks("藤井電工発注ストップ分.xlsx").Sheets(1).Range("D2:D76")

'追記したい文字列を指定
Dim AdditionalNote As String
AdditionalNote = "卸分在庫有 16年4月"

For Each r In Rng

    Dim Code As String
    Code = r.Value
    
    Dim c As Range
    
    'B列を検索して、該当六ケタがあれば、仕入先に追記する
    With Workbooks("発注用商品情報.xlsm").Worksheets("商品情報").Columns(2)
    
        Set c = .Find(what:=Code, LookIn:=xlValues, LookAt:=xlWhole)

        If Not c Is Nothing Then
           '最初のセルのアドレスを覚える
           FirstAddress = c.Address
           
           '繰返し検索し、条件を満たすすべてのセルを検索する
           Do
              
               c.Offset(0, 2) = c.Offset(0, 2) & " " & AdditionalNote
               
               Set c = .FindNext(c)
               If c Is Nothing Then Exit Do
           
           Loop Until c.Address = FirstAddress
         
         End If

    End With

Next

End Sub
