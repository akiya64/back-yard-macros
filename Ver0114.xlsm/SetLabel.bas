Attribute VB_Name = "SetLabel"
Private Sub Labeling()

Dim rng As Range

r = 317

Do While Cells(r, 2).Value <> ""

    Id = Cells(r, 2).Value

    Set rng = Range(Cells(r, 1), Cells(r, 7))
    
    i = 0
    
    Do 'リサイズすべき行数を決定する、オフセットして注文番号をチェック
        
        i = i + 1
    
    Loop While rng.Cells(1, 2).Offset(i, 0).Value = Id
    
    If i > 1 Then
        
        Set rng = rng.Resize(i, 7)
    
    End If
         
    Dim Label As String
   
    Label = "NO" & Id
    
    Call SetLabel(Label, rng)
    
    r = r + i

Loop

End Sub

Private Sub SetLabel(Name As String, rng As Range)

End Sub
