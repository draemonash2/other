VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'$memoを複数開くとフィルタダイアログ表示ショートカットキーが正しく機能しないため、無効化する
'Private Sub Workbook_Open()
'    Application.OnKey "%^{DOWN}", "フィルタダイアログ表示"
'    Application.OnKey "%^{UP}", "フィルタダイアログ表示"
'End Sub

