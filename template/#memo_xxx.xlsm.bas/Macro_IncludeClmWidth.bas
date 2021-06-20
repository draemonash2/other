Attribute VB_Name = "Macro_IncludeClmWidth"
Option Explicit

'include column width v1.0.0

' ====================================================================
' = 概要：選択シートの列幅を別ブックから取り込む。
' = 覚書：・列順は考慮しない。
' ====================================================================
Public Sub 列幅取り込み()
    Const sMACRO_NAME = "列幅取り込み"
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
        
    '取り込み対象ブック名取得
    Dim sInBookPath As String
    sInBookPath = ShowFileSelectDialog( _
        ActiveWorkbook.Path & "\" & ActiveWorkbook.Name, _
        sMACRO_NAME, _
        "" _
    ) '★TODO:フィルタ指定追加
    
    '取り込み対象ブックオープン
    Dim wInBook As Workbook
    Set wInBook = ExcelFileOpen(sInBookPath)
    
    'アクティブシート切り替え
    '  取り込み対象ブックを開いたことでアクティブシートが変わるため、
    '  ActiveWindow.SelectedSheetsを実行する前にThisWorkbookをactivateする。
    ThisWorkbook.Sheets(1).Activate
    
    '選択シート一覧取得
    Dim cSelectSheetNames As Variant
    Set cSelectSheetNames = CreateObject("System.Collections.ArrayList")
    Dim shMySheet As Worksheet
    For Each shMySheet In ActiveWindow.SelectedSheets
        cSelectSheetNames.Add shMySheet.Name
    Next
    
    '列幅コピー＆ペースト
    Dim vSelectSheetName As Variant
    For Each vSelectSheetName In cSelectSheetNames
        Set shMySheet = ThisWorkbook.Sheets(vSelectSheetName)
        
        'シート存在確認
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
            '列幅コピー
            shInSheet.Range( _
                shInSheet.Cells(1, 1), _
                shInSheet.Cells(1, shInSheet.UsedRange.Columns.Count + 1) _
            ).Copy
            '列幅貼り付け
            shMySheet.Range( _
                shMySheet.Cells(1, 1), _
                shMySheet.Cells(1, shInSheet.UsedRange.Columns.Count + 1) _
            ).PasteSpecial (xlPasteColumnWidths)
        Else
            'シートが取り込み元ブックに存在しない場合、デバッグメッセージを出して無視する
            Debug.Print "[Error]" & sMACRO_NAME & " : " & shMySheet & "シートが取り込み元ブックにありません"
        End If
    Next
    
    '選択解除
    For Each shMySheet In ThisWorkbook.Sheets
        shMySheet.Activate
        shMySheet.Cells(1, 1).Select
    Next
    ThisWorkbook.Sheets(1).Activate
    
    '取り込み対象ブッククローズ
    Call ExcelFileClose(wInBook, False)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "取り込み完了！", vbOKOnly, sMACRO_NAME
End Sub


' ==================================================================
' = 概要    ファイル（単一）選択ダイアログを表示する
' = 引数    sInitPath       String  [in]  デフォルトファイルパス（省略可）
' = 引数    sTitle          String  [in]  タイトル名（省略可）
' = 引数    sFilters        String  [in]  選択時のフィルタ（省略可）(※)
' = 戻値                    String        選択ファイル
' = 覚書    (※)ダイアログのフィルタ指定方法は以下。
' =              ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =                    ・拡張子が複数ある場合は、";"で区切る
' =                    ・ファイル種別と拡張子は"/"で区切る
' =                    ・フィルタが複数ある場合、","で区切る
' =         sFilters が省略もしくは空文字の場合、フィルタをクリアする。
' = 依存    Mng_FileSys.bas/SetDialogFilters()
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function ShowFileSelectDialog( _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sTitle As String = "", _
    Optional ByVal sFilters As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    If sTitle = "" Then
        fdDialog.Title = "ファイルを選択してください"
    Else
        fdDialog.Title = sTitle
    End If
    fdDialog.AllowMultiSelect = False
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) 'フィルタ追加
    
    'ダイアログ表示
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
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
' = 概要    ShowFileSelectDialog() と ShowFilesSelectDialog() 用の関数
' =         ダイアログのフィルタを追加する。指定方法は以下。
' =           ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =               ・拡張子が複数ある場合は、";"で区切る
' =               ・ファイル種別と拡張子は"/"で区切る
' =               ・フィルタが複数ある場合、","で区切る
' = 引数    sFilters    String      [in]    フィルタ
' = 引数    fdDialog    FileDialog  [out]   ファイルダイアログ
' = 戻値    なし
' = 覚書    sFilters が空文字の場合、フィルタをクリアする。
' = 依存    なし
' = 所属    Mng_FileSys.bas
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
                            "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                            """/"" は一つだけ指定してください" & vbNewLine & _
                            "  " & vFilters(lFilterIdx)
                        MsgBox "処理を中断します。"
                        End
                    End If
                Else
                    MsgBox _
                        "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                        "種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
                        "  " & vFilters(lFilterIdx)
                    MsgBox "処理を中断します。"
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
                        "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                        """/"" は一つだけ指定してください" & vbNewLine & _
                        "  " & sFilters
                    MsgBox "処理を中断します。"
                    End
                End If
            Else
                MsgBox _
                    "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                    "種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
                    "  " & sFilters
                MsgBox "処理を中断します。"
                End
            End If
        End If
    End If
End Function



