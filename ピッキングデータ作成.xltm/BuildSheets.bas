Attribute VB_Name = "BuildSheets"
Option Explicit

Sub CreateSorterSheet(MallName As String)
'振分用シートへ商品情報を転記する。
'テンプレートシートを2回コピー、単体商品とセット商品用を用意する。
'受注データのシートを2～最終行まで、受注モール・受注時コードを判定しつつシートへコピー

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
    Order(1) = ValidateName(Range("C" & i).Value) '商品名、≪≫など削除した上で転記
    Order(2) = Range("D" & i).Value '受注数量
    Order(3) = CStr(Range("L" & i).Value) 'JAN
    Order(4) = Range("K" & i).Value '有効ロケーション
    Order(5) = Range("N" & i).Value '現在庫
    
    
    '現在庫が取得できてないときは、印刷レイアウトの関係のため空白1文字入れておく
    If Order(5) = "" Then Order(5) = " "
    
    '商魂JANが空欄かつ、受注時コードがJANならJAN項目に入れる
    If Order(3) = "" Then
        Dim RawCode As String
        RawCode = Range("B" & i).Value
        If RawCode Like String(13, "#") _
            And Not RawCode Like "77777*" Then
                Order(3) = RawCode
        End If
    End If
    
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


Call SortHasLocation(ForSorterSheet)

ForSorterSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
ForSorterSetItemSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous


'終了処理 Sheet内容確定

'念のため幅を再指定
Call AdjustWidth(ForSorterSheet)
Call AdjustWidth(ForSorterSetItemSheet)

ForSorterSheet.Protect
ForSorterSetItemSheet.Protect

End Sub

Sub OutputPickingData(ByVal MallName As String)
'電算室提出の棚有りピッキングシート、棚なしピッキングシートを作成する。
'ピッキング用のブックを二つ用意する 棚有り-2-3と -a
'モール名・有効ロケーションを判定しつつ、受注データシートを1行ずつコピー。
'全行終わったら、-2-3と -aの二つのブックは閉じる。


'引数の名前で新規ブックを作成する
'ファイル名はAmazon-ピッキングシート、Yahoo=ヤフーPシート、電算室側の処理の関係で固定
Dim BookName As String
If MallName = "Amazon" Then
    BookName = "ピッキングシート"
ElseIf MallName = "Yahoo" Then
    BookName = "ヤフーPシート"
Else
    BookName = MallName & "Pシート"
End If

'提出用ファイルを用意
'100番/200番棚有り -2-3、電算室提出
Dim ForSlimsBook As Workbook, ForSlimsSheet As Worksheet
Set ForSlimsBook = PreparePickingBook(BookName & Format(Date, "MMdd") & "-2-3")
Set ForSlimsSheet = ForSlimsBook.Worksheets(1)

'登録無し、棚無し -a
Dim NoEntryItemBook As Workbook, NoEntryItemSheet As Worksheet
Set NoEntryItemBook = PreparePickingBook(BookName & Format(Date, "MMdd") & "-a")
Set NoEntryItemSheet = NoEntryItemBook.Worksheets(1)

OrderSheet.Activate

Dim i As Long, k As Long, j As Long, Order(6) As Variant
'受注データシート行カウンタ
i = 2

'棚無しシート行カウンタ
j = 3

'100番シート行カウンタ
k = 3

'用意したブックへ1行ずつコピー
Do

    '引数で渡されたモール以外は飛ばす
    If Not Range("F" & i).Value Like (MallName & "*") Then GoTo Continue
    
    '受注時コードの7777は電算提出データに含めない。
    If Range("I" & i).Value Like "7777*" Then GoTo Continue

    '提出するコードの振替
    'SLIMSに投入するのはロケーション有りの6ケタのみ
    Dim OrderedCode As String, ForAddinCode As String, AddinResultCode As String, Code As String
    
    OrderedCode = CStr(Range("B" & i).Value)
    ForAddinCode = CStr(Range("I" & i).Value)
    AddinResultCode = CStr(Range("M" & i).Value)
    
    If AddinResultCode <> "" Then
        Code = AddinResultCode
    ElseIf ForAddinCode <> "" Then
        Code = ForAddinCode
    Else
        Code = OrderedCode
    End If
    
    '配列に提出ファイル1行分のデータを格納
    'アマゾンのみ、電算室処理でアマゾン注文番号を判定している、連番不可
    If MallName = "Amazon" Then
        Order(0) = CStr(Range("H" & i).Value) 'モール側採番の注文番号
    Else
        Order(0) = CStr(Range("A" & i).Value) 'クロスモール採番の連番
    End If
    
    Order(1) = CStr(Code) '商品コード
    Order(2) = ValidateName(Range("C" & i).Value)  '商品名≪≫など削除して転記
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
ForSlimsBook.Activate
With ForSlimsSheet

    If MallName = "Amazon" Then
        .Columns("G").Insert
        .Range("G2").Value = "送料"
        .Range(Cells(3, 7), Cells(ForSlimsSheet.UsedRange.Rows.Count, 7)).Value = 0
    End If

    '罫線引いて保存
    .Range("A2:I2").Resize(Range("B2").CurrentRegion.Rows.Count - 1, 9).Borders.LineStyle = xlContinuous
    
End With
ForSlimsBook.Close SaveChanges:=True

NoEntryItemSheet.Activate
With NoEntryItemSheet
    .Activate
    
    If MallName = "Amazon" Then
        .Columns("G").Insert
        .Range("G1").Value = "送料"
        .Range(Cells(2, 7), Cells(.UsedRange.Rows.Count, 7)).Value = 0
    End If
    
    '罫線引いて保存
    .Range("A2:I2").Resize(Range("B2").CurrentRegion.Rows.Count - 1, 9).Borders.LineStyle = xlContinuous
    
End With
NoEntryItemBook.Close SaveChanges:=True

End Sub

Private Function PreparePickingBook(ByVal BookName As String) As Workbook
'ピッキングシート用のブックを用意する。
'VBAでブック名でブックオブジェクトを呼び出せるよう、最初に引数のブック名で保存して、そのワークブックオブジェクトを返す。
'ピッキングシートファイルの重複判定もここで行う。
'既に同じファイル名があれば、ファイル名AR入りの楽天・プライム分のピッキングシートかダイアログでユーザーに決めてもらう。

Const PICKING_FOLDER As String = "\\server02\商品部\ネット販売関連\ピッキング\" '最後、必ず\マーク

ThisWorkbook.Worksheets("ピッキングシート提出用テンプレート").Copy
ActiveSheet.Name = BookName

'一旦新規作成ブックを保存することでブック名を変更する
'新規作成ファイルの保存時はファイルフォーマットを明示する必要な模様
Dim SavePath As String, SaveFolder As String

'保存先と保存ファイル名の決定

'ネット販売のフォルダに繋がるか判定
If Dir(PICKING_FOLDER, vbDirectory) <> "" Then
    SaveFolder = PICKING_FOLDER
Else
    SaveFolder = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\"
    MsgBox "ネット販売関連に繋がらないため、" & BookName & "をデスクトップに保存します。"
End If

'あす楽・プライム分のピッキングか？
If Main.IsSecondPicking = True Then
    BookName = Replace(BookName, Format(Date, "MMdd"), (Format(Date, "MMdd") & "AR"))
End If

'時刻保存のフラグがあるか
If Main.IsTimeStampMode = True Then
    BookName = Replace(BookName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

'ファイルがあって､あす楽プライム分のピッキングでない時のみ､選択ダイアログを表示
If Dir(PICKING_FOLDER & BookName & ".xlsx") <> "" And Main.IsSecondPicking = False Then
    
    Dim IsAR As Integer
    IsAR = MsgBox(prompt:="本日分のファイルが既に存在します。" & vbLf & "あす楽・プライム分として保存しますか？", _
            Buttons:=vbExclamation + vbYesNo)
    
    'あす楽プライムモードのフラグを立てる
    If IsAR = vbYes Then
        BookName = Replace(BookName, Format(Date, "MMdd"), (Format(Date, "MMdd") & "AR"))
        Main.IsSecondPicking = True
    
    'あす楽プライム分でないとき、時刻を含めたファイル名保存フラグを立てる
    Else
        BookName = Replace(BookName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
        Main.IsTimeStampMode = True
    
    End If

End If
    
'上記条件に全てヒットしない場合は、当日1回目の生成となり、BookNameは変更されていない。

If Dir(PICKING_FOLDER & BookName & ".xlsx") <> "" Then
    '保存しようとするファイル名で、既にファイルがある場合ファイル名を日付-時刻とする
    BookName = Replace(BookName, Format(Date, "MMdd"), (Format(Date, "MMdd") & "-" & Format(Time, "hhmm")))
End If

SavePath = SaveFolder & BookName

ActiveWorkbook.Sheets(1).Name = BookName
ActiveWorkbook.SaveAs Filename:=SavePath, FileFormat:=xlWorkbookDefault

Set PreparePickingBook = ActiveWorkbook

ThisWorkbook.ActiveSheet.Activate

End Function

Private Sub AdjustWidth(TargetSheet As Worksheet)
'A4横の一枚に収まるように振分シートの列幅を再調整する。

'列幅 調整時にアラートが出るのを抑止
Application.DisplayAlerts = False

Dim WidthArr As Variant
WidthArr = Array(19, 82, 4, 16, 14, 6)

TargetSheet.Activate

Dim k As Long
For k = 0 To 5
    TargetSheet.Columns(k + 1).ColumnWidth = WidthArr(k)
Next

Application.DisplayAlerts = True

End Sub

Private Sub SortHasLocation(Sheet As Worksheet)
'振分シートで、棚無しを下に集めて商品コード昇順に並び替える。

Dim SortRange As Range
Set SortRange = Sheet.Range("A1").CurrentRegion

Dim CodeRange As Range
Set CodeRange = Sheet.Range("A2:A" & SortRange.Rows.Count)

'ソート条件をセット
With Sheet.Sort
    
    '一旦ソートをクリア
    .SortFields.Clear
    
    'ソートキーをセット 第一キー 商品コード：色、第二キー 商品コード：昇順
    .SortFields.Add Key:=CodeRange, SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=CodeRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    'ソート対象のデータが入ってる範囲を指定して
    .SetRange SortRange
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    
    'セットした条件を適用
    .Apply

End With

'カレントリージョンがセレクトされているので、選択セルをセルA1にセットし直す
Sheet.Activate
Range("A1").Select

End Sub
