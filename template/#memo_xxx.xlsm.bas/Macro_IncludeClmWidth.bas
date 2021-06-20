Attribute VB_Name = "Macro_IncludeClmWidth"
Option Explicit

'include column width v1.0.0

' ====================================================================
' = �T�v�F�I���V�[�g�̗񕝂�ʃu�b�N�����荞�ށB
' = �o���F�E�񏇂͍l�����Ȃ��B
' ====================================================================
Public Sub �񕝎�荞��()
    Const sMACRO_NAME = "�񕝎�荞��"
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
        
    '��荞�ݑΏۃu�b�N���擾
    Dim sInBookPath As String
    sInBookPath = ShowFileSelectDialog( _
        ActiveWorkbook.Path & "\" & ActiveWorkbook.Name, _
        sMACRO_NAME, _
        "" _
    ) '��TODO:�t�B���^�w��ǉ�
    
    '��荞�ݑΏۃu�b�N�I�[�v��
    Dim wInBook As Workbook
    Set wInBook = ExcelFileOpen(sInBookPath)
    
    '�A�N�e�B�u�V�[�g�؂�ւ�
    '  ��荞�ݑΏۃu�b�N���J�������ƂŃA�N�e�B�u�V�[�g���ς�邽�߁A
    '  ActiveWindow.SelectedSheets�����s����O��ThisWorkbook��activate����B
    ThisWorkbook.Sheets(1).Activate
    
    '�I���V�[�g�ꗗ�擾
    Dim cSelectSheetNames As Variant
    Set cSelectSheetNames = CreateObject("System.Collections.ArrayList")
    Dim shMySheet As Worksheet
    For Each shMySheet In ActiveWindow.SelectedSheets
        cSelectSheetNames.Add shMySheet.Name
    Next
    
    '�񕝃R�s�[���y�[�X�g
    Dim vSelectSheetName As Variant
    For Each vSelectSheetName In cSelectSheetNames
        Set shMySheet = ThisWorkbook.Sheets(vSelectSheetName)
        
        '�V�[�g���݊m�F
        Dim bExistSheet As Boolean
        bExistSheet = False
        Dim shInSheet As Worksheet
        For Each shInSheet In wInBook.Sheets
            If shInSheet.Name = shMySheet.Name Then
                bExistSheet = True
                Exit For
            End If
        Next
        If bExistSheet = True Then
            '�񕝃R�s�[
            shInSheet.Range( _
                shInSheet.Cells(1, 1), _
                shInSheet.Cells(1, shInSheet.UsedRange.Columns.Count + 1) _
            ).Copy
            '�񕝓\��t��
            shMySheet.Range( _
                shMySheet.Cells(1, 1), _
                shMySheet.Cells(1, shInSheet.UsedRange.Columns.Count + 1) _
            ).PasteSpecial (xlPasteColumnWidths)
        Else
            '�V�[�g����荞�݌��u�b�N�ɑ��݂��Ȃ��ꍇ�A�f�o�b�O���b�Z�[�W���o���Ė�������
            Debug.Print "[Error]" & sMACRO_NAME & " : " & shMySheet & "�V�[�g����荞�݌��u�b�N�ɂ���܂���"
        End If
    Next
    
    '�I������
    For Each shMySheet In ThisWorkbook.Sheets
        shMySheet.Activate
        shMySheet.Cells(1, 1).Select
    Next
    ThisWorkbook.Sheets(1).Activate
    
    '��荞�ݑΏۃu�b�N�N���[�Y
    Call ExcelFileClose(wInBook, False)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "��荞�݊����I", vbOKOnly, sMACRO_NAME
End Sub


' ==================================================================
' = �T�v    �t�@�C���i�P��j�I���_�C�A���O��\������
' = ����    sInitPath       String  [in]  �f�t�H���g�t�@�C���p�X�i�ȗ��j
' = ����    sTitle          String  [in]  �^�C�g�����i�ȗ��j
' = ����    sFilters        String  [in]  �I�����̃t�B���^�i�ȗ��j(��)
' = �ߒl                    String        �I���t�@�C��
' = �o��    (��)�_�C�A���O�̃t�B���^�w����@�͈ȉ��B
' =              ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
' =                    �E�g���q����������ꍇ�́A";"�ŋ�؂�
' =                    �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
' =                    �E�t�B���^����������ꍇ�A","�ŋ�؂�
' =         sFilters ���ȗ��������͋󕶎��̏ꍇ�A�t�B���^���N���A����B
' = �ˑ�    Mng_FileSys.bas/SetDialogFilters()
' = ����    Mng_FileSys.bas
' ==================================================================
Private Function ShowFileSelectDialog( _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sTitle As String = "", _
    Optional ByVal sFilters As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    If sTitle = "" Then
        fdDialog.Title = "�t�@�C����I�����Ă�������"
    Else
        fdDialog.Title = sTitle
    End If
    fdDialog.AllowMultiSelect = False
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) '�t�B���^�ǉ�
    
    '�_�C�A���O�\��
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then '�L�����Z������
        ShowFileSelectDialog = ""
    Else
        Dim sSelectedPath As String
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FileExists(sSelectedPath) Then
            ShowFileSelectDialog = sSelectedPath
        Else
            ShowFileSelectDialog = ""
        End If
    End If
    
    Set fdDialog = Nothing
End Function

' ==================================================================
' = �T�v    ShowFileSelectDialog() �� ShowFilesSelectDialog() �p�̊֐�
' =         �_�C�A���O�̃t�B���^��ǉ�����B�w����@�͈ȉ��B
' =           ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
' =               �E�g���q����������ꍇ�́A";"�ŋ�؂�
' =               �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
' =               �E�t�B���^����������ꍇ�A","�ŋ�؂�
' = ����    sFilters    String      [in]    �t�B���^
' = ����    fdDialog    FileDialog  [out]   �t�@�C���_�C�A���O
' = �ߒl    �Ȃ�
' = �o��    sFilters ���󕶎��̏ꍇ�A�t�B���^���N���A����B
' = �ˑ�    �Ȃ�
' = ����    Mng_FileSys.bas
' ==================================================================
Private Function SetDialogFilters( _
    ByVal sFilters As String, _
    ByRef fdDialog As FileDialog _
)
    fdDialog.Filters.Clear
    If sFilters = "" Then
        'Do Nothing
    Else
        Dim vFilter As Variant
        If InStr(sFilters, ",") > 0 Then
            Dim vFilters As Variant
            vFilters = Split(sFilters, ",")
            Dim lFilterIdx As Long
            For lFilterIdx = 0 To UBound(vFilters)
                If InStr(vFilters(lFilterIdx), "/") > 0 Then
                    vFilter = Split(vFilters(lFilterIdx), "/")
                    If UBound(vFilter) = 1 Then
                        fdDialog.Filters.Add vFilter(0), vFilter(1), lFilterIdx + 1
                    Else
                        MsgBox _
                            "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                            """/"" �͈�����w�肵�Ă�������" & vbNewLine & _
                            "  " & vFilters(lFilterIdx)
                        MsgBox "�����𒆒f���܂��B"
                        End
                    End If
                Else
                    MsgBox _
                        "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                        "��ʂƊg���q�� ""/"" �ŋ�؂��Ă��������B" & vbNewLine & _
                        "  " & vFilters(lFilterIdx)
                    MsgBox "�����𒆒f���܂��B"
                    End
                End If
            Next lFilterIdx
        Else
            If InStr(sFilters, "/") > 0 Then
                vFilter = Split(sFilters, "/")
                If UBound(vFilter) = 1 Then
                    fdDialog.Filters.Add vFilter(0), vFilter(1), 1
                Else
                    MsgBox _
                        "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                        """/"" �͈�����w�肵�Ă�������" & vbNewLine & _
                        "  " & sFilters
                    MsgBox "�����𒆒f���܂��B"
                    End
                End If
            Else
                MsgBox _
                    "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                    "��ʂƊg���q�� ""/"" �ŋ�؂��Ă��������B" & vbNewLine & _
                    "  " & sFilters
                MsgBox "�����𒆒f���܂��B"
                End
            End If
        End If
    End If
End Function



