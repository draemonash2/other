Attribute VB_Name = "Mng_ExcelFile"
Option Explicit
 
' excel file manage library v1.01

'************************************************************
'* �\���̒�`
'************************************************************
Type T_EXCEL_FILE_INFO
    wTrgtBook As Workbook
    sFilePath As String
    bIsAlreadyOpen As Boolean
End Type

'************************************************************
'* ���W���[���� �ϐ���`
'************************************************************
Dim gatExcelFileInfo() As T_EXCEL_FILE_INFO

'************************************************************
'* �֐���`
'************************************************************
' ==================================================================
' = �T�v    ����������
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelFile.bas
' ==================================================================
Public Function ExcelFileInfoInit()
    Dim atExcelFileInfoInit() As T_EXCEL_FILE_INFO
    gatExcelFileInfo = atExcelFileInfoInit '������
End Function

' ==================================================================
' = �T�v    �G�N�Z���t�@�C���I�[�v��
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelFile.bas
' ==================================================================
Public Function ExcelFileOpen( _
    ByVal sTrgtBookPath As String _
) As Workbook
    Dim wTrgtBook As Workbook
    Dim bIsAlreadyOpen As Boolean
    Dim lExcelFileIdx As Long
 
    '�z��Ē�`
    If Sgn(gatExcelFileInfo) = 0 Then '���������z��
        lExcelFileIdx = 0
    Else '�v�f���P�ȏ�̔z��
        lExcelFileIdx = UBound(gatExcelFileInfo) + 1
    End If
    ReDim Preserve gatExcelFileInfo(lExcelFileIdx)
 
    '�u�b�N���݃`�F�b�N
    If Dir(sTrgtBookPath) = "" Then
        MsgBox "�u�b�N�u" & sTrgtBookPath & "�v" & vbNewLine & "�����݂��܂���B", vbExclamation
        Exit Function
    Else
        'Do Nothing
    End If
 
    '�u�b�N�����ɊJ����Ă��邩�`�F�b�N
    bIsAlreadyOpen = False
    For Each wTrgtBook In Workbooks
        If wTrgtBook.Name = Dir(sTrgtBookPath) Then
            bIsAlreadyOpen = True
        Else
            'Do Nothing
        End If
    Next wTrgtBook
 
    '�t�@�C���I�[�v��
    If bIsAlreadyOpen = True Then
        Set wTrgtBook = Workbooks(Dir(sTrgtBookPath))
    Else
        Set wTrgtBook = Workbooks.Open(sTrgtBookPath)
    End If
 
    '�u�b�N���i�[
    gatExcelFileInfo(lExcelFileIdx).bIsAlreadyOpen = bIsAlreadyOpen
    gatExcelFileInfo(lExcelFileIdx).sFilePath = sTrgtBookPath
    Set gatExcelFileInfo(lExcelFileIdx).wTrgtBook = wTrgtBook
 
    Set ExcelFileOpen = wTrgtBook
End Function

' ==================================================================
' = �T�v    �G�N�Z���t�@�C���N���[�Y
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelFile.bas
' ==================================================================
Public Function ExcelFileClose( _
    ByVal wTrgtBook As Workbook, _
    ByVal bIsSave As Boolean _
)
    Dim lExcelFileIdx As Long
    Dim lArrayIdx As Long
    Dim atExcelFileInfoInit() As T_EXCEL_FILE_INFO
 
    '�J���Ă���u�b�N������
    For lExcelFileIdx = 0 To UBound(gatExcelFileInfo)
        If gatExcelFileInfo(lExcelFileIdx).sFilePath = wTrgtBook.FullName Then
            Exit For
        Else
            'Do Nothing
        End If
    Next lExcelFileIdx
 
    Debug.Assert UBound(gatExcelFileInfo) >= lExcelFileIdx
 
    '����
    If gatExcelFileInfo(lExcelFileIdx).bIsAlreadyOpen = False Then
        If bIsSave = True Then
            gatExcelFileInfo(lExcelFileIdx).wTrgtBook.Close SaveChanges:=True
        Else
            gatExcelFileInfo(lExcelFileIdx).wTrgtBook.Close SaveChanges:=False
        End If
    Else
        'Do Nothing
    End If
 
    '�z��v�f�폜
    If UBound(gatExcelFileInfo) = 0 Then
        gatExcelFileInfo = atExcelFileInfoInit
    Else
        For lArrayIdx = lExcelFileIdx To UBound(gatExcelFileInfo) - 1
            gatExcelFileInfo(lArrayIdx).bIsAlreadyOpen = gatExcelFileInfo(lArrayIdx + 1).bIsAlreadyOpen
            gatExcelFileInfo(lArrayIdx).sFilePath = gatExcelFileInfo(lArrayIdx + 1).sFilePath
            Set gatExcelFileInfo(lArrayIdx).wTrgtBook = gatExcelFileInfo(lArrayIdx + 1).wTrgtBook
        Next lArrayIdx
        ReDim Preserve gatExcelFileInfo(UBound(gatExcelFileInfo) - 1)
    End If
End Function

