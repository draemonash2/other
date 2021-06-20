Attribute VB_Name = "Macro_FilterAtKeyword"
Option Explicit

'filter at keyword v1.2.0

' ====================================================================
' = �T�v�F�ݒ肳��Ă���I�[�g�t�B���^�͈̔͂ɑ΂��āA
' =       �I�[�g�t�B���^������s���B����̓R�}���h�{�^����
' =       �������ǂݎ���āA�ȉ��̑�����s���B
' =         �E�u�S�����v�̏ꍇ
' =             �S��̍i���݂���������
' =         �E�u�����v�̏ꍇ
' =             �R�}���h�{�^�����z�u���ꂽ��ɑ΂���i���݂�
' =             ��������
' =         �E��L�ȊO�̏ꍇ
' =             �R�}���h�{�^�����z�u���ꂽ��ɑ΂��āA
' =             �R�}���h�{�^���̕�����ōi���݂��s��
' =                 ��j�R�}���h�{�^���̕����񂪁u�������v�̏ꍇ
' =                       �R�}���h�{�^�����z�u���ꂽ��̒��ŁA
' =                       �u�������v�Ɉ�v����s�𒊏o����B
' =
' = ���ӁF�K���R�}���h�{�^���������̓I�[�g�V�F�C�v����{�}�N����
' =       �Ăяo������
' =
' = �o���F�E������v���������ꍇ�A���C���h�J�[�h���g�p���邱��
' =           ��j�{�^���̕����񂪁u*TODO*�v���w�肷��ƁuTODO�v��
' =               �܂ލs�𒊏o����
' =       �E�u�󔒃Z���v�𒊏o�������ꍇ�͋�s�����邱��
' =       �E�n�q�����������ꍇ�͉��s��}�����邱��
' =           ��j�{�^���̕����񂪁u������(���s)�����v�̏ꍇ�A
' =               �u�������v�������́u�����v�Ɉ�v����s�𒊏o����
' =
' ====================================================================
Public Sub �I�[�g�t�B���^����with�{�^��()
    Dim objBtn
    Set objBtn = ActiveSheet.Shapes(Application.Caller)
    Dim sBtnText As String
    If objBtn.AutoShapeType = msoShapeMixed Then
        sBtnText = objBtn.AlternativeText
    Else
        sBtnText = objBtn.TextEffect.Text
    End If
    MsgBox sBtnText
    Call FilterAtKeyword(sBtnText, objBtn.TopLeftCell.Column)
End Sub

' ====================================================================
' = �T�v�F�ݒ肳��Ă���I�[�g�t�B���^�͈̔͂ɑ΂��āA
' =       �J�e�S���s�̃L�[���[�h�Ńt�B���^����B
' ====================================================================
Public Sub �I�[�g�t�B���^����at���ݍs�J�e�S��()
Attribute �I�[�g�t�B���^����at���ݍs�J�e�S��.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim lClmIdx As Long
    
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
    lClmIdx = rFindResult.Column

    If ActiveCell.Row = ActiveSheet.ListObjects(1).HeaderRowRange.Row Then
        '�A�N�e�B�u�Z�����^�C�g���s�̏ꍇ�A�t�B���^����������
        Call FilterAtKeyword("����", lClmIdx)
    Else
        Call FilterAtKeyword(ActiveSheet.Cells(ActiveCell.Row, lClmIdx).Value, lClmIdx)
    End If
End Sub

' ====================================================================
' = �T�v�F�ݒ肳��Ă���I�[�g�t�B���^�͈̔͂ɑ΂��āA
' =       ������or�ۗ��Ńt�B���^����B
' ====================================================================
Public Sub �I�[�g�t�B���^����at������ۗ�()
    Const sFILTER_KEYWORD As String = "*��*" & vbLf & "*��*"
    Dim lClmIdx As Long
    
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
    lClmIdx = rFindResult.Column

    If ActiveCell.Row = ActiveSheet.ListObjects(1).HeaderRowRange.Row Then
        '�A�N�e�B�u�Z�����^�C�g���s�̏ꍇ�A�t�B���^����������
        Call FilterAtKeyword("����", lClmIdx)
        Debug.Print "����"
    Else
        Call FilterAtKeyword(sFILTER_KEYWORD, lClmIdx)
        Debug.Print sFILTER_KEYWORD
    End If
End Sub

' ==================================================================
' = �T�v    �I�[�g�t�B���^�𑀍삷��
' = ����    sFilterKeyword  [in]    String  �t�B���^�L�[���[�h
' = ����    lTrgtClm        [in]    Long    �Ώۗ�
' = �ߒl    �Ȃ�
' = �o��    �E�t�B���^�L�[���[�h��"����"���w�肵���ꍇ�A�Ώۗ�̃t�B���^����������
' =         �E�t�B���^�L�[���[�h��"�S����"���w�肵���ꍇ�A�t�B���^�͈͓��̃t�B���^��S�ĉ�������
' =         �EsDELIMITER��؂�Ńt�B���^�L�[���[�h���w�肵���ꍇ�AOR�����Ńt�B���^����
' = �ˑ�    �Ȃ�
' = ����    Macro_FilterAtKeyword.bas
' ==================================================================
Private Sub FilterAtKeyword( _
    ByVal sFilterKeyword, _
    ByVal lTrgtClm _
)
    Const sKEYWORD_RELEASE As String = "����"
    Const sKEYWORD_ALL_RELEASE As String = "�S����"
    Const sDELIMITER As String = vbLf
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    Dim rFltTrgtRng As Range
    Set rFltTrgtRng = ActiveSheet.AutoFilter.Range
    '�e�[�u���͈͊O�Ńt�B���^�͈͎擾���́A�^�C�g����Ɉړ����Ă���ēx�擾
    If Err.Number <> 0 Then
        Dim lTitleRow As Long
        lTitleRow = ActiveSheet.ListObjects(1).HeaderRowRange.Row
        ActiveSheet.Cells(lTitleRow, ActiveCell.Column()).Activate
        Set rFltTrgtRng = ActiveSheet.AutoFilter.Range
    End If
    On Error GoTo 0
    
    Select Case sFilterKeyword
        '*** �S���� ***
        Case sKEYWORD_ALL_RELEASE:
            rFltTrgtRng.AutoFilter '�S����
            rFltTrgtRng.AutoFilter '�Đݒ�
        
        '*** ���� ***
        Case sKEYWORD_RELEASE:
            rFltTrgtRng.AutoFilter _
                Field:=lTrgtClm
        
        '*** ����ȊO�i�t�B���^�ݒ�j ***
        Case Else:
            If InStr(sFilterKeyword, sDELIMITER) Then
                Dim vBtnTexts As Variant
                vBtnTexts = Split(sFilterKeyword, sDELIMITER)
                
                '�󔒃Z�����o�̂��߂ɐ��^����
                Dim lIdx As Long
                For lIdx = LBound(vBtnTexts) To UBound(vBtnTexts)
                    If vBtnTexts(lIdx) = "" Then
                        vBtnTexts(lIdx) = "="
                    Else
                        'Do Nothing
                    End If
                Next lIdx
                
                '�t�B���^�����O���s
                rFltTrgtRng.AutoFilter _
                    Field:=lTrgtClm, _
                    Criteria1:=vBtnTexts, _
                    Operator:=xlFilterValues
            Else
                '�t�B���^�����O���s
                rFltTrgtRng.AutoFilter _
                    Field:=lTrgtClm, _
                    Criteria1:=sFilterKeyword, _
                    Operator:=xlFilterValues
            End If
    End Select
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

