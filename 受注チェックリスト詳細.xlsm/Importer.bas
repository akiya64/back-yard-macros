Attribute VB_Name = "Importer"
Option Explicit
Sub 受注チェックCSV読込()

GetOrderCheckListPath


End Sub
Private Function GetOrderCheckListPath(FileName As String) As String
'ピッキングシートの-a＝棚無のセット分解前ファイルを探してフルパスをセット

Const SANTYOKU_DUMP_FOLDER As String = "\\Server02\商品部\ネット販売関連\ピッキング\" '末尾\マーク必須

'楽天の場合、楽天Pシート0627-a.xls

'実行時バインディング ScriptingRuntimeはDictionary配列使うのに必要で参照ONだから、事前バインディングでいいかも。
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim f As Object, Newest As Object
      
'事前バインディング
'Dim FSO As FileSystemObject
'Set FSO = New FileSystemObject

'Dim f As File, Newest As File


'指定フォルダー内のFileNameを含むファイル名を調べて、最新のファイルを1つ取得する。
'LINQか何か、1構文で済むの欲しい

For Each f In FSO.GetFolder(SANTYOKU_DUMP_FOLDER).Files

    If f.Name Like FileName & ".csv" Then
    
        Set Newest = f
    
        Exit For
    End If

Next


RetrievePickingFilePath = PICKING_FILE_FOLDER & Newest.Name

End Function
