Attribute VB_Name = "Macro_UpdateCategoryValue"
Option Explicit

' ====================================================================
' = �T�v�F�J�e�S����̒l���X�V����
' ====================================================================
Public Sub �J�e�S����_�l�X�V()
Attribute �J�e�S����_�l�X�V.VB_ProcData.VB_Invoke_Func = "m\n14"
    Const lCTGRY_CLM As Long = 2
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim lClmIdx As Long
    lClmIdx = lCTGRY_CLM '�������̂��ߗ�ԍ��͌Œ�
    
    '�J�e�S���� �l����
    Dim sCellValueBfr As String
    Dim sCellValueAft As String
    sCellValueBfr = ActiveSheet.Cells(ActiveCell.Row, lClmIdx).Value
    Select Case sCellValueBfr
        Case "TODO��"
            sCellValueAft = "TODO��"
        Case "TODO��"
            sCellValueAft = "TODO��"
        Case "TODO��"
            sCellValueAft = "MEMO"
        Case "MEMO"
            sCellValueAft = "TODO��"
        Case Else
            Debug.Assert 0
    End Select
    
    '�J�e�S���� �l�X�V
    Dim vCell As Variant
    For Each vCell In Selection
        ActiveSheet.Cells(vCell.Row, lClmIdx).Value = sCellValueAft
    Next
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

