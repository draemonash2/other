Option Explicit

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" ) 'DownloadFile()

'===============================================================================
'= 設定
'===============================================================================
Const sDOWNLOAD_URL = "https://github.com/draemonash2/other/archive/master.zip"
Const sLOCAL_OBJECT_NAME = "other-master"

'===============================================================================
'= 本処理
'===============================================================================
Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'=== 事前処理 ===
Dim sDiffTrgtDirPath
Dim sOutputMsg
Dim sDownloadTrgtDirPath
Dim sDownloadTrgtFilePath
Dim sDiffSrcDirPath
Dim sUnzipProgramPath
Dim sDiffProgramPath
sDiffTrgtDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )
sOutputMsg = WScript.ScriptName & " 「" & sLOCAL_OBJECT_NAME & "」"
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sLOCAL_OBJECT_NAME & ".zip"
sDiffSrcDirPath = sDownloadTrgtDirPath & "\" & sLOCAL_OBJECT_NAME
sUnzipProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_7Z%")
If InStr(sUnzipProgramPath, "%") > 0 then
    MsgBox "環境変数「MYEXEPATH_7Z」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
End If
sDiffProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_WINMERGE%")
If InStr(sDiffProgramPath, "%") > 0 then
    MsgBox "環境変数「MYEXEPATH_WINMERGE」が設定されていません。" & vbNewLine & "処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
end if

Dim vAnswer
vAnswer = MsgBox("ダウンロードを開始します。", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
    MsgBox "キャンセルが押されたため、処理を中断します。", vbExclamation, sOutputMsg
    WScript.Quit
End If

'=== ダウンロード ===
Call DownloadFile(sDOWNLOAD_URL, sDownloadTrgtFilePath)

'=== 解凍 ===
objWshShell.Run """" & sUnzipProgramPath & """ x -o""" & sDownloadTrgtDirPath & """ """ & sDownloadTrgtFilePath & """", 0, True

'=== フォルダ比較 ===
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 10, True

'=== フォルダ削除 ===
vAnswer = MsgBox("ダウンロードフォルダを削除しますか？", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    objFSO.DeleteFile sDownloadTrgtFilePath, True
    objFSO.DeleteFolder sDiffSrcDirPath, True
End If

MsgBox("処理が完了しました！", vbYesNo, sOutputMsg)

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

