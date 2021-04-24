Option Explicit

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" ) 'DownloadFile()

'===============================================================================
'= �ݒ�
'===============================================================================
Const sDOWNLOAD_URL = "https://github.com/draemonash2/other/archive/master.zip"
Const sLOCAL_OBJECT_NAME = "other-master"

'===============================================================================
'= �{����
'===============================================================================
Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'=== ���O���� ===
Dim sDiffTrgtDirPath
Dim sOutputMsg
Dim sDownloadTrgtDirPath
Dim sDownloadTrgtFilePath
Dim sDiffSrcDirPath
Dim sUnzipProgramPath
Dim sDiffProgramPath
sDiffTrgtDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )
sOutputMsg = WScript.ScriptName & " �u" & sLOCAL_OBJECT_NAME & "�v"
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sLOCAL_OBJECT_NAME & ".zip"
sDiffSrcDirPath = sDownloadTrgtDirPath & "\" & sLOCAL_OBJECT_NAME
sUnzipProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_7Z%")
If InStr(sUnzipProgramPath, "%") > 0 then
    MsgBox "���ϐ��uMYEXEPATH_7Z�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sOutputMsg
    WScript.Quit
End If
sDiffProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_WINMERGE%")
If InStr(sDiffProgramPath, "%") > 0 then
    MsgBox "���ϐ��uMYEXEPATH_WINMERGE�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sOutputMsg
    WScript.Quit
end if

Dim vAnswer
vAnswer = MsgBox("�_�E�����[�h���J�n���܂��B", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
    MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sOutputMsg
    WScript.Quit
End If

'=== �_�E�����[�h ===
Call DownloadFile(sDOWNLOAD_URL, sDownloadTrgtFilePath)

'=== �� ===
objWshShell.Run """" & sUnzipProgramPath & """ x -o""" & sDownloadTrgtDirPath & """ """ & sDownloadTrgtFilePath & """", 0, True

'=== �t�H���_��r ===
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 10, True

'=== �t�H���_�폜 ===
vAnswer = MsgBox("�_�E�����[�h�t�H���_���폜���܂����H", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    objFSO.DeleteFile sDownloadTrgtFilePath, True
    objFSO.DeleteFolder sDiffSrcDirPath, True
End If

MsgBox("�������������܂����I", vbYesNo, sOutputMsg)

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

