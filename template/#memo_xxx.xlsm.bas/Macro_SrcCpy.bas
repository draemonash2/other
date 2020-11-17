Attribute VB_Name = "Macro_SrcCpy"
Option Explicit

Public Sub ソースコードをフォルダ構造ごとコピー()
    Const KEYWORD_RELATIVE_PATH As String = "相対パス"
    Const KEYWORD_CPYSRC_PATH As String = "コピー元"
    Const KEYWORD_CPYDST_PATH As String = "コピー先"
    
    'ボタン位置(列)取得
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
            MsgBox sSrchKeyword & "が見つかりませんでした"
            MsgBox "処理を中断します"
            Exit Sub
        End If
        lStrtRow = rFindResult.Row
        lLastRow = .Cells(.Rows.Count, lTrgtClm).End(xlUp).Row
        
        Dim sCpySrcRootPath As String
        Set rFindResult = .Cells.Find(KEYWORD_CPYSRC_PATH, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            MsgBox sSrchKeyword & "が見つかりませんでした"
            MsgBox "処理を中断します"
            Exit Sub
        End If
        sCpySrcRootPath = .Cells(rFindResult.Row, lTrgtClm).Value
        
        Dim sCpyDstRootPath As String
        Set rFindResult = .Cells.Find(KEYWORD_CPYDST_PATH, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            MsgBox sSrchKeyword & "が見つかりませんでした"
            MsgBox "処理を中断します"
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
            
            'フォルダ作成
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
    
    MsgBox "コピー完了！"
    
End Sub

Public Function CreateDirectry( _
    ByVal sDirPath As String _
)
    Dim sParentDir As String
    Dim oFileSys As Object
 
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
 
    sParentDir = oFileSys.GetParentFolderName(sDirPath)
 
    '親ディレクトリが存在しない場合、再帰呼び出し
    If oFileSys.FolderExists(sParentDir) = False Then
        Call CreateDirectry(sParentDir)
    End If
 
    'ディレクトリ作成
    If oFileSys.FolderExists(sDirPath) = False Then
        oFileSys.CreateFolder sDirPath
    End If
 
    Set oFileSys = Nothing
End Function

' ==================================================================
' = 概要    指定されたファイルパスからフォルダパスを抽出する
' = 引数    sFilePath       String  [in]  ファイルパス
' = 引数    bErrorEnable    Boolean [in]  エラー発生有効/無効(※)
' = 戻値                    Variant       フォルダパス
' = 覚書    ローカルファイルパス（例：c:\test）や URL （例：https://test）
' =         が指定可能。
' =         (※) bErrorEnable にてファイルパス以外が指定された時の返却値を
' =         変えることが出来る｡
' =            True  : sFilePath を返却
' =            False : エラー値（xlErrNA）を返却
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
            GetDirPath = CVErr(xlErrNA)  'エラー値
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
        vRet = GetDirPath("C:_test_a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' エラー 2042
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
' = 概要    末尾区切り文字以降の文字列を除去する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        除去文字列
' = 覚書    なし
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
' = 概要    末尾区切り文字以降の文字列を返却する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        抽出文字列
' = 覚書    なし
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

