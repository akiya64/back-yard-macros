Attribute VB_Name = "AppendNoteFromXlsx"
Sub InsertCaution()

'元のエクセルブック名、元シートの範囲は毎回指定のこと
'イミディエイトで、Workbooks(1).nameでワークブック名が確認できる。
Set Rng = Workbooks("発注ストップ分.xlsx").Sheets(1).Range("D2:D76")

'追記したい文字列を指定
Dim AdditionalNote As String
AdditionalNote = "発注ストップ"

For Each r In Rng

    Dim Code As String
    Code = r.Value
    
    Dim c As Range
    
    'B列を検索して、該当コードがあれば、仕入先列に追記する
    With Workbooks("商品リスト.xlsm").Worksheets("商品情報").Columns(2)
    
        Set c = .Find(what:=Code, LookIn:=xlValues, LookAt:=xlWhole)

        If Not c Is Nothing Then
           '最初のセルのアドレスを記録
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
