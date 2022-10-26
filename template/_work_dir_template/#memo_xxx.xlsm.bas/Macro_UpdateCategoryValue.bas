Attribute VB_Name = "Macro_UpdateCategoryValue"
Option Explicit

' ====================================================================
' = 概要：カテゴリ列の値を更新する
' ====================================================================
Public Sub カテゴリ列_値更新()
Attribute カテゴリ列_値更新.VB_ProcData.VB_Invoke_Func = "m\n14"
    Const lCTGRY_CLM As Long = 2
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim lClmIdx As Long
    lClmIdx = lCTGRY_CLM '高速化のため列番号は固定
    
    'カテゴリ列 値判定
    Dim sCellValueBfr As String
    Dim sCellValueAft As String
    sCellValueBfr = ActiveSheet.Cells(ActiveCell.Row, lClmIdx).Value
    Select Case sCellValueBfr
        Case "TODO未"
            sCellValueAft = "TODO完"
        Case "TODO完"
            sCellValueAft = "TODO保"
        Case "TODO保"
            sCellValueAft = "MEMO"
        Case "MEMO"
            sCellValueAft = "TODO未"
        Case Else
            Debug.Assert 0
    End Select
    
    'カテゴリ列 値更新
    Dim vCell As Variant
    For Each vCell In Selection
        ActiveSheet.Cells(vCell.Row, lClmIdx).Value = sCellValueAft
    Next
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

