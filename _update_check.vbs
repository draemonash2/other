Option Explicit

Const sDownloadUrl = "https://github.com/draemonash2/other/archive/master.zip"
Const sLocalObjectName = "other-master"
Const sDiffTrgtDirPath = "C:\other"

Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sScriptName
sScriptName = WScript.ScriptName
Dim sDownloadTrgtDirPath
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
Dim sDownloadTrgtFilePath
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sLocalObjectName & ".zip"
Dim sDiffSrcDirPath
sDiffSrcDirPath = sDownloadTrgtDirPath & "\" & sLocalObjectName

Dim vAnswer
'=== �_�E�����[�h ===
vAnswer = MsgBox("�_�E�����[�h���J�n���܂��B", vbOkCancel, sScriptName)
If vAnswer = vbCancel Then
	MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sScriptName
	WScript.Quit
End If
CreateObject("Shell.Application").ShellExecute "microsoft-edge:" & sDownloadUrl
vAnswer = MsgBox("�_�E�����[�h������������uOK�v�������Ă��������B" & vbNewLine & "�𓀂��J�n���܂��B", vbOkCancel, sScriptName)
If vAnswer = vbCancel Then
	MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sScriptName
	WScript.Quit
End If

'=== �� ===
Dim sUnzipProgramPath
sUnzipProgramPath = objWshShell.Environment("System").Item("MYPATH_7Z")
If sUnzipProgramPath = "" then
	MsgBox "���ϐ��uMYPATH_7Z�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sScriptName
	WScript.Quit
End If
objWshShell.Run """" & sUnzipProgramPath & """ x -o""" & sDownloadTrgtDirPath & """ """ & sDownloadTrgtFilePath & """"
vAnswer = MsgBox("�𓀂�����������uOK�v�������Ă��������B" & vbNewLine & "�t�H���_��r���J�n���܂��B", vbOkCancel, sScriptName)
If vAnswer = vbCancel Then
	MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sScriptName
	WScript.Quit
End If

'=== �t�H���_��r ===
Dim sDiffProgramPath
sDiffProgramPath = objWshShell.Environment("System").Item("MYPATH_WINMERGE")
If sDiffProgramPath = "" then
	MsgBox "���ϐ��uMYPATH_WINMERGE�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sScriptName
	WScript.Quit
end if
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 0, True

vAnswer = MsgBox("�_�E�����[�h�t�H���_���폜���܂����H", vbYesNo, sScriptName)
If vAnswer = vbYes Then
	objFSO.DeleteFile sDownloadTrgtFilePath, True
	objFSO.DeleteFolder sDiffSrcDirPath, True
End If

