Attribute VB_Name = "Macro_SrcCpy"
Option Explicit

Public Sub �\�[�X�R�[�h���t�H���_�\�����ƃR�s�[()
    Const KEYWORD_RELATIVE_PATH As String = "���΃p�X"
    Const KEYWORD_CPYSRC_PATH As String = "�R�s�[��"
    Const KEYWORD_CPYDST_PATH As String = "�R�s�[��"
    
    '�{�^���ʒu(��)�擾
    Dim objBtn As Object
    Set objBtn = ActiveSheet.Shapes(Application.Caller)
    Dim lTrgtClm As Long
    lTrgtClm = objBtn.TopLeftCell.Column
    
    With ActiveSheet
        Dim sSrchKeyword As String
        Dim lStrtRow As Long
        Dim lLastRow As Long
        Dim rFindResult As Range
        Set rFindResult = .Cells.Find(KEYWORD_RELATIVE_PATH, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            MsgBox sSrchKeyword & "��������܂���ł���"
            MsgBox "�����𒆒f���܂�"
            Exit Sub
        End If
        lStrtRow = rFindResult.Row
        lLastRow = .Cells(.Rows.Count, lTrgtClm).End(xlUp).Row
        
        Dim sCpySrcRootPath As String
        Set rFindResult = .Cells.Find(KEYWORD_CPYSRC_PATH, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            MsgBox sSrchKeyword & "��������܂���ł���"
            MsgBox "�����𒆒f���܂�"
            Exit Sub
        End If
        sCpySrcRootPath = .Cells(rFindResult.Row, lTrgtClm).Value
        
        Dim sCpyDstRootPath As String
        Set rFindResult = .Cells.Find(KEYWORD_CPYDST_PATH, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            MsgBox sSrchKeyword & "��������܂���ł���"
            MsgBox "�����𒆒f���܂�"
            Exit Sub
        End If
        sCpyDstRootPath = .Cells(rFindResult.Row, lTrgtClm).Value
    
        Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        Dim lRowIdx As Long
        Dim sRltvPath As String
        For lRowIdx = lStrtRow To lLastRow
            sRltvPath = .Cells(lRowIdx, lTrgtClm).Value
            sRltvPath = Replace(sRltvPath, "/", "\")
            Dim sSrcPath As String
            Dim sDstPath As String
            
            '�t�H���_�쐬
            sSrcPath = sCpySrcRootPath & "\" & sRltvPath
            sDstPath = sCpyDstRootPath & "\" & sRltvPath
            
            Dim sDstParDir As String
            sDstParDir = GetDirPath(sDstPath)
            Call CreateDirectry(sDstParDir)
            
            If objFSO.FileExists(sSrcPath) Then
                objFSO.CopyFile sSrcPath, sDstPath, True
            End If
            Debug.Print sSrcPath & " " & sDstPath
        Next lRowIdx
    
    End With
    
    MsgBox "�R�s�[�����I"
    
End Sub

Public Function CreateDirectry( _
    ByVal sDirPath As String _
)
    Dim sParentDir As String
    Dim oFileSys As Object
 
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
 
    sParentDir = oFileSys.GetParentFolderName(sDirPath)
 
    '�e�f�B���N�g�������݂��Ȃ��ꍇ�A�ċA�Ăяo��
    If oFileSys.FolderExists(sParentDir) = False Then
        Call CreateDirectry(sParentDir)
    End If
 
    '�f�B���N�g���쐬
    If oFileSys.FolderExists(sDirPath) = False Then
        oFileSys.CreateFolder sDirPath
    End If
 
    Set oFileSys = Nothing
End Function

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�H���_�p�X�𒊏o����
' = ����    sFilePath       String  [in]  �t�@�C���p�X
' = ����    bErrorEnable    Boolean [in]  �G���[�����L��/����(��)
' = �ߒl                    Variant       �t�H���_�p�X
' = �o��    ���[�J���t�@�C���p�X�i��Fc:\test�j�� URL �i��Fhttps://test�j
' =         ���w��\�B
' =         (��) bErrorEnable �ɂăt�@�C���p�X�ȊO���w�肳�ꂽ���̕ԋp�l��
' =         �ς��邱�Ƃ��o����
' =            True  : sFilePath ��ԋp
' =            False : �G���[�l�ixlErrNA�j��ԋp
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath As String, _
    Optional ByVal bErrorEnable As Boolean = False _
) As Variant
    If InStr(sFilePath, "\") Then
        GetDirPath = RemoveTailWord(sFilePath, "\")
    ElseIf InStr(sFilePath, "/") Then
        GetDirPath = RemoveTailWord(sFilePath, "/")
    Else
        If bErrorEnable = True Then
            GetDirPath = CVErr(xlErrNA)  '�G���[�l
        Else
            GetDirPath = sFilePath
        End If
    End If
End Function
    Private Sub Test_GetDirPath()
        Dim Result As String
        Dim vRet As Variant
        Result = "[Result]"
        vRet = GetDirPath("C:\test\a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' C:\test
        vRet = GetDirPath("http://test/a", True): Result = Result & vbNewLine & CStr(vRet)  ' http://test
        vRet = GetDirPath("C:_test_a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' �G���[ 2042
        Result = Result & vbNewLine                                                         '
        vRet = GetDirPath("C:\test\a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' C:\test
        vRet = GetDirPath("http://test/a", False): Result = Result & vbNewLine & CStr(vRet) ' http://test
        vRet = GetDirPath("C:_test_a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' C:_test_a.txt
        Result = Result & vbNewLine                                                         '
        vRet = GetDirPath("C:\test\a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' C:\test
        vRet = GetDirPath("http://test/a"): Result = Result & vbNewLine & CStr(vRet)        ' http://test
        vRet = GetDirPath("C:_test_a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' C:_test_a.txt
        Debug.Print Result
    End Sub

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕��������������B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ����������
' = �o��    �Ȃ�
' ==================================================================
Public Function RemoveTailWord( _
    ByVal sStr As String, _
    ByVal sDlmtr As String _
) As String
    Dim sTailWord As String
    Dim lRemoveLen As Long
    
    If sStr = "" Then
        RemoveTailWord = ""
    Else
        If sDlmtr = "" Then
            RemoveTailWord = sStr
        Else
            If InStr(sStr, sDlmtr) = 0 Then
                RemoveTailWord = sStr
            Else
                sTailWord = ExtractTailWord(sStr, sDlmtr)
                lRemoveLen = Len(sDlmtr) + Len(sTailWord)
                RemoveTailWord = Left$(sStr, Len(sStr) - lRemoveLen)
            End If
        End If
    End If
End Function
    Private Sub Test_RemoveTailWord()
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & RemoveTailWord("", "\")               '
        Result = Result & vbNewLine & RemoveTailWord("c:\a", "\")           ' c:
        Result = Result & vbNewLine & RemoveTailWord("c:\a\", "\")          ' c:\a
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b", "\")         ' c:\a
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\", "\")        ' c:\a\b
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\c.txt", "\")   ' c:\a\b
        Result = Result & vbNewLine & RemoveTailWord("c:\\b\c.txt", "\")    ' c:\\b
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\c.txt", "")    ' c:\a\b\c.txt
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\c.txt", "\\")  ' c:\a\b\c.txt
        Result = Result & vbNewLine & RemoveTailWord("c:\a\\b\c.txt", "\\") ' c:\a
        Result = Result & vbNewLine & "*** test finished! ***"
        Debug.Print Result
    End Sub

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕������ԋp����B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ���o������
' = �o��    �Ȃ�
' ==================================================================
Public Function ExtractTailWord( _
    ByVal sStr As String, _
    ByVal sDlmtr As String _
) As String
    Dim asSplitWord() As String
    
    If Len(sStr) = 0 Then
        ExtractTailWord = ""
    Else
        ExtractTailWord = ""
        asSplitWord = Split(sStr, sDlmtr)
        ExtractTailWord = asSplitWord(UBound(asSplitWord))
    End If
End Function
    Private Sub Test_ExtractTailWord()
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & ExtractTailWord("", "\")               '
        Result = Result & vbNewLine & ExtractTailWord("c:\a", "\")           ' a
        Result = Result & vbNewLine & ExtractTailWord("c:\a\", "\")          '
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b", "\")         ' b
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\", "\")        '
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\c.txt", "\")   ' c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\\b\c.txt", "\")    ' c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\c.txt", "")    ' c:\a\b\c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\c.txt", "\\")  ' c:\a\b\c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\a\\b\c.txt", "\\") ' b\c.txt
        Result = Result & vbNewLine & "*** test finished! ***"
        Debug.Print Result
    End Sub

