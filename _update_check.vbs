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
'=== ダウンロード ===
vAnswer = MsgBox("ダウンロードを開始します。", vbOkCancel, sScriptName)
If vAnswer = vbCancel Then
	MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sScriptName
	WScript.Quit
End If
CreateObject("Shell.Application").ShellExecute "microsoft-edge:" & sDownloadUrl
vAnswer = MsgBox("ダウンロードが完了したら「OK」を押してください。" & vbNewLine & "解凍を開始します。", vbOkCancel, sScriptName)
If vAnswer = vbCancel Then
	MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sScriptName
	WScript.Quit
End If

'=== 解凍 ===
Dim sUnzipProgramPath
sUnzipProgramPath = objWshShell.Environment("System").Item("MYPATH_7Z")
If sUnzipProgramPath = "" then
	MsgBox "環境変数「MYPATH_7Z」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sScriptName
	WScript.Quit
End If
objWshShell.Run """" & sUnzipProgramPath & """ x -o""" & sDownloadTrgtDirPath & """ """ & sDownloadTrgtFilePath & """"
vAnswer = MsgBox("解凍が完了したら「OK」を押してください。" & vbNewLine & "フォルダ比較を開始します。", vbOkCancel, sScriptName)
If vAnswer = vbCancel Then
	MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sScriptName
	WScript.Quit
End If

'=== フォルダ比較 ===
Dim sDiffProgramPath
sDiffProgramPath = objWshShell.Environment("System").Item("MYPATH_WINMERGE")
If sDiffProgramPath = "" then
	MsgBox "環境変数「MYPATH_WINMERGE」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sScriptName
	WScript.Quit
end if
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 0, True

vAnswer = MsgBox("ダウンロードフォルダを削除しますか？", vbYesNo, sScriptName)
If vAnswer = vbYes Then
	objFSO.DeleteFile sDownloadTrgtFilePath, True
	objFSO.DeleteFolder sDiffSrcDirPath, True
End If

