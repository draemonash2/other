Attribute VB_Name = "Macro_UpdateCategoryValue"
Option Explicit

' ====================================================================
' = �T�v�F�J�e�S����̒l���X�V����
' ====================================================================
Public Sub �J�e�S����_�l�X�V()
Attribute �J�e�S����_�l�X�V.VB_ProcData.VB_Invoke_Func = "m\n14"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim rFindResult As Range
    Dim sFindKeyword As String
    Dim shTrgtSht As Worksheet
    Set shTrgtSht = ActiveSheet
    sFindKeyword = "�J�e�S��"
    Set rFindResult = shTrgtSht.Cells.Find(sFindKeyword, LookAt:=xlWhole)
    If rFindResult Is Nothing Then
        MsgBox _
            "�Z����������Ȃ��������߁A�����𒆒f���܂��B" & vbNewLine & _
            "�@�����ΏۃV�[�g�F" & shTrgtSht.Name & vbNewLine & _
            "�@�����ΏۃL�[���[�h�F" & sFindKeyword, _
            vbCritical
        End
    End If
    
    Dim lClmIdx As Long
    lClmIdx = rFindResult.Column
    
    Dim cCategoryCell As Range
    Set cCategoryCell = ActiveSheet.Cells(ActiveCell.Row, lClmIdx)
    
    '���͋K���ݒ肠��H
    If cCategoryCell.Validation.Type = xlValidateList Then
        '�J�e�S���� ���͋K���擾
        Dim vInputRuleLists As Variant
        Dim sInputRule As String
        sInputRule = cCategoryCell.Validation.Formula1
        vInputRuleLists = Split(sInputRule, ",")
        
        '���͋K������
        Dim lIdx As Long
        For lIdx = LBound(vInputRuleLists) To UBound(vInputRuleLists)
            If vInputRuleLists(lIdx) = cCategoryCell.Value Then
                Exit For
            End If
        Next lIdx
        
        '�J�e�S���� �l�X�V
        If (lIdx + 1) > UBound(vInputRuleLists) Then
            cCategoryCell.Value = vInputRuleLists(LBound(vInputRuleLists))
        Else
            cCategoryCell.Value = vInputRuleLists(lIdx + 1)
        End If
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub



