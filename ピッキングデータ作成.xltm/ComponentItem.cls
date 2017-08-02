VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'セットを構成する商品1つを表すクラス。
'ユーザー定義型では配列に入らないので、フィールドのみのクラスを作成。

Public Jan As String
Public Code As String
Public Name As String
Public Quantity As Long
