Attribute VB_Name = "Filter"
Option Explicit


Sub SetStatusFilter()

With yahoo6digit
    
    .Activate
       
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    '"ＳＰ扱い" 以下の文字列は、商魂の「区分」に合わせてください
    '空白と登録なしをフィルターで非表示にします、マクロでフィルターを記録して書き換えると楽です。
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
    "ＳＰ扱い", "メ廃番品", "限定品", "在庫処分", "在庫廃番", "直送扱い", "登録のみ", "販売中止", "販路限定", "標準"), Operator _
    :=xlFilterValues

End With

End Sub
