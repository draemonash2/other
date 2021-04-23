Option Explicit

Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Const sDOWNLOAD_URL = "https://github.com/draemonash2/other/archive/master.zip"
Const sLOCAL_OBJECT_NAME = "other-master"

Dim sDiffTrgtDirPath
sDiffTrgtDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )
Dim sOutputMsg
sOutputMsg = WScript.ScriptName & " �u" & sLOCAL_OBJECT_NAME & "�v"
Dim sDownloadTrgtDirPath
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
Dim sDownloadTrgtFilePath
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sLOCAL_OBJECT_NAME & ".zip"
Dim sDiffSrcDirPath
sDiffSrcDirPath = sDownloadTrgtDirPath & "\" & sLOCAL_OBJECT_NAME

Dim vAnswer
'=== �_�E�����[�h ===
vAnswer = MsgBox("�_�E�����[�h���J�n���܂��B", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
	MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sOutputMsg
	WScript.Quit
End If
CreateObject("Shell.Application").ShellExecute "microsoft-edge:" & sDOWNLOAD_URL
vAnswer = MsgBox("�_�E�����[�h������������uOK�v�������Ă��������B" & vbNewLine & "�𓀂��J�n���܂��B", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
	MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sOutputMsg
	WScript.Quit
End If

'=== �� ===
Dim sUnzipProgramPath
sUnzipProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_7Z%")
If sUnzipProgramPath = "" then
	MsgBox "���ϐ��uMYEXEPATH_7Z�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sOutputMsg
	WScript.Quit
End If
objWshShell.Run """" & sUnzipProgramPath & """ x -o""" & sDownloadTrgtDirPath & """ """ & sDownloadTrgtFilePath & """"
vAnswer = MsgBox("�𓀂�����������uOK�v�������Ă��������B" & vbNewLine & "�t�H���_��r���J�n���܂��B", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
	MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sOutputMsg
	WScript.Quit
End If

'=== �t�H���_��r ===
Dim sDiffProgramPath
sDiffProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_WINMERGE%")
If sDiffProgramPath = "" then
	MsgBox "���ϐ��uMYEXEPATH_WINMERGE�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sOutputMsg
	WScript.Quit
end if
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 10, True

vAnswer = MsgBox("�_�E�����[�h�t�H���_���폜���܂����H", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
	objFSO.DeleteFile sDownloadTrgtFilePath, True
	objFSO.DeleteFolder sDiffSrcDirPath, True
End If

