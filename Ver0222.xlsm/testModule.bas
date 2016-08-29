Attribute VB_Name = "testModule"
Private Sub test()

Dim o As Dictionary
Set o = OrderSheet.getUndispatchOrders

For Each v In o
    Debug.Print o(v).Id & ":" & o(v).CanNotShipping
Next

End Sub
