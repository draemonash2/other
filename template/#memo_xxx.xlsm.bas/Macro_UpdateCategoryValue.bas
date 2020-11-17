Attribute VB_Name = "Macro_UpdateCategoryValue"
Option Explicit

' ====================================================================
' = 概要：カテゴリ列の値を更新する
' ====================================================================
Public Sub カテゴリ列_値更新()
Attribute カテゴリ列_値更新.VB_ProcData.VB_Invoke_Func = "m\n14"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim rFindResult As Range
    Dim sFindKeyword As String
    Dim shTrgtSht As Worksheet
    Set shTrgtSht = ActiveSheet
    sFindKeyword = "カテゴリ"
    Set rFindResult = shTrgtSht.Cells.Find(sFindKeyword, LookAt:=xlWhole)
    If rFindResult Is Nothing Then
        MsgBox _
            "セルが見つからなかったため、処理を中断します。" & vbNewLine & _
            "　検索対象シート：" & shTrgtSht.Name & vbNewLine & _
            "　検索対象キーワード：" & sFindKeyword, _
            vbCritical
        End
    End If
    
    Dim lClmIdx As Long
    lClmIdx = rFindResult.Column
    
    Dim cCategoryCell As Range
    Set cCategoryCell = ActiveSheet.Cells(ActiveCell.Row, lClmIdx)
    
    '入力規則設定あり？
    If cCategoryCell.Validation.Type = xlValidateList Then
        'カテゴリ列 入力規則取得
        Dim vInputRuleLists As Variant
        Dim sInputRule As String
        sInputRule = cCategoryCell.Validation.Formula1
        vInputRuleLists = Split(sInputRule, ",")
        
        '入力規則検索
        Dim lIdx As Long
        For lIdx = LBound(vInputRuleLists) To UBound(vInputRuleLists)
            If vInputRuleLists(lIdx) = cCategoryCell.Value Then
                Exit For
            End If
        Next lIdx
        
        'カテゴリ列 値更新
        If (lIdx + 1) > UBound(vInputRuleLists) Then
            cCategoryCell.Value = vInputRuleLists(LBound(vInputRuleLists))
        Else
            cCategoryCell.Value = vInputRuleLists(lIdx + 1)
        End If
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub



