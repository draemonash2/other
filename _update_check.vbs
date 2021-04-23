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
sOutputMsg = WScript.ScriptName & " 「" & sLOCAL_OBJECT_NAME & "」"
Dim sDownloadTrgtDirPath
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
Dim sDownloadTrgtFilePath
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sLOCAL_OBJECT_NAME & ".zip"
Dim sDiffSrcDirPath
sDiffSrcDirPath = sDownloadTrgtDirPath & "\" & sLOCAL_OBJECT_NAME

Dim vAnswer
'=== ダウンロード ===
vAnswer = MsgBox("ダウンロードを開始します。", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
	MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sOutputMsg
	WScript.Quit
End If
CreateObject("Shell.Application").ShellExecute "microsoft-edge:" & sDOWNLOAD_URL
vAnswer = MsgBox("ダウンロードが完了したら「OK」を押してください。" & vbNewLine & "解凍を開始します。", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
	MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sOutputMsg
	WScript.Quit
End If

'=== 解凍 ===
Dim sUnzipProgramPath
sUnzipProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_7Z%")
If sUnzipProgramPath = "" then
	MsgBox "環境変数「MYEXEPATH_7Z」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sOutputMsg
	WScript.Quit
End If
objWshShell.Run """" & sUnzipProgramPath & """ x -o""" & sDownloadTrgtDirPath & """ """ & sDownloadTrgtFilePath & """"
vAnswer = MsgBox("解凍が完了したら「OK」を押してください。" & vbNewLine & "フォルダ比較を開始します。", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
	MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sOutputMsg
	WScript.Quit
End If

'=== フォルダ比較 ===
Dim sDiffProgramPath
sDiffProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_WINMERGE%")
If sDiffProgramPath = "" then
	MsgBox "環境変数「MYEXEPATH_WINMERGE」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sOutputMsg
	WScript.Quit
end if
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 10, True

vAnswer = MsgBox("ダウンロードフォルダを削除しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
	objFSO.DeleteFile sDownloadTrgtFilePath, True
	objFSO.DeleteFolder sDiffSrcDirPath, True
End If

