Attribute VB_Name = "Macro_FilterAtKeyword"
Option Explicit

'filter at keyword v1.2.0

' ====================================================================
' = 概要：設定されているオートフィルタの範囲に対して、
' =       オートフィルタ操作を行う。操作はコマンドボタンの
' =       文字列を読み取って、以下の操作を行う。
' =         ・「全解除」の場合
' =             全列の絞込みを解除する
' =         ・「解除」の場合
' =             コマンドボタンが配置された列に対する絞込みを
' =             解除する
' =         ・上記以外の場合
' =             コマンドボタンが配置された列に対して、
' =             コマンドボタンの文字列で絞込みを行う
' =                 例）コマンドボタンの文字列が「未完了」の場合
' =                       コマンドボタンが配置された列の中で、
' =                       「未完了」に一致する行を抽出する。
' =
' = 注意：必ずコマンドボタンもしくはオートシェイプから本マクロを
' =       呼び出すこと
' =
' = 覚書：・部分一致させたい場合、ワイルドカードを使用すること
' =           例）ボタンの文字列が「*TODO*」を指定すると「TODO」を
' =               含む行を抽出する
' =       ・「空白セル」を抽出したい場合は空行を入れること
' =       ・ＯＲ検索したい場合は改行を挿入すること
' =           例）ボタンの文字列が「未完了(改行)完了」の場合、
' =               「未完了」もしくは「完了」に一致する行を抽出する
' =
' ====================================================================
Public Sub オートフィルタ操作withボタン()
    Dim objBtn
    Set objBtn = ActiveSheet.Shapes(Application.Caller)
    Dim sBtnText As String
    If objBtn.AutoShapeType = msoShapeMixed Then
        sBtnText = objBtn.AlternativeText
    Else
        sBtnText = objBtn.TextEffect.Text
    End If
    MsgBox sBtnText
    Call FilterAtKeyword(sBtnText, objBtn.TopLeftCell.Column)
End Sub

' ====================================================================
' = 概要：設定されているオートフィルタの範囲に対して、
' =       カテゴリ行のキーワードでフィルタする。
' ====================================================================
Public Sub オートフィルタ操作at現在行カテゴリ()
Attribute オートフィルタ操作at現在行カテゴリ.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim lClmIdx As Long
    
    Dim rFindResult As Range
    Dim sFindKeyword As String
    Dim shTrgtSht As Worksheet
    Set shTrgtSht = ActiveSheet
    
    sFindKeyword = "カテゴリ"
    Set rFindResult = shTrgtSht.Cells.Find(sFindKeyword, LookAt:=xlWhole)
    If rFindResult Is Nothing Then
        MsgBox _
            "セルが見つからなかったため、処理を中断します。" & vbNewLine & _
            "　検索対象シート：" & shTrgtSht.Name & vbNewLine & _
            "　検索対象キーワード：" & sFindKeyword, _
            vbCritical
        End
    End If
    lClmIdx = rFindResult.Column

    If ActiveCell.Row = ActiveSheet.ListObjects(1).HeaderRowRange.Row Then
        'アクティブセルがタイトル行の場合、フィルタを解除する
        Call FilterAtKeyword("解除", lClmIdx)
    Else
        Call FilterAtKeyword(ActiveSheet.Cells(ActiveCell.Row, lClmIdx).Value, lClmIdx)
    End If
End Sub

' ====================================================================
' = 概要：設定されているオートフィルタの範囲に対して、
' =       未着手or保留でフィルタする。
' ====================================================================
Public Sub オートフィルタ操作at未着手保留()
    Const sFILTER_KEYWORD As String = "*未*" & vbLf & "*保*"
    Dim lClmIdx As Long
    
    Dim rFindResult As Range
    Dim sFindKeyword As String
    Dim shTrgtSht As Worksheet
    Set shTrgtSht = ActiveSheet
    
    sFindKeyword = "カテゴリ"
    Set rFindResult = shTrgtSht.Cells.Find(sFindKeyword, LookAt:=xlWhole)
    If rFindResult Is Nothing Then
        MsgBox _
            "セルが見つからなかったため、処理を中断します。" & vbNewLine & _
            "　検索対象シート：" & shTrgtSht.Name & vbNewLine & _
            "　検索対象キーワード：" & sFindKeyword, _
            vbCritical
        End
    End If
    lClmIdx = rFindResult.Column

    If ActiveCell.Row = ActiveSheet.ListObjects(1).HeaderRowRange.Row Then
        'アクティブセルがタイトル行の場合、フィルタを解除する
        Call FilterAtKeyword("解除", lClmIdx)
        Debug.Print "解除"
    Else
        Call FilterAtKeyword(sFILTER_KEYWORD, lClmIdx)
        Debug.Print sFILTER_KEYWORD
    End If
End Sub

' ==================================================================
' = 概要    オートフィルタを操作する
' = 引数    sFilterKeyword  [in]    String  フィルタキーワード
' = 引数    lTrgtClm        [in]    Long    対象列
' = 戻値    なし
' = 覚書    ・フィルタキーワードに"解除"を指定した場合、対象列のフィルタを解除する
' =         ・フィルタキーワードに"全解除"を指定した場合、フィルタ範囲内のフィルタを全て解除する
' =         ・sDELIMITER区切りでフィルタキーワードを指定した場合、OR条件でフィルタする
' = 依存    なし
' = 所属    Macro_FilterAtKeyword.bas
' ==================================================================
Private Sub FilterAtKeyword( _
    ByVal sFilterKeyword, _
    ByVal lTrgtClm _
)
    Const sKEYWORD_RELEASE As String = "解除"
    Const sKEYWORD_ALL_RELEASE As String = "全解除"
    Const sDELIMITER As String = vbLf
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    Dim rFltTrgtRng As Range
    Set rFltTrgtRng = ActiveSheet.AutoFilter.Range
    'テーブル範囲外でフィルタ範囲取得時は、タイトル列に移動してから再度取得
    If Err.Number <> 0 Then
        Dim lTitleRow As Long
        lTitleRow = ActiveSheet.ListObjects(1).HeaderRowRange.Row
        ActiveSheet.Cells(lTitleRow, ActiveCell.Column()).Activate
        Set rFltTrgtRng = ActiveSheet.AutoFilter.Range
    End If
    On Error GoTo 0
    
    Select Case sFilterKeyword
        '*** 全解除 ***
        Case sKEYWORD_ALL_RELEASE:
            rFltTrgtRng.AutoFilter '全解除
            rFltTrgtRng.AutoFilter '再設定
        
        '*** 解除 ***
        Case sKEYWORD_RELEASE:
            rFltTrgtRng.AutoFilter _
                Field:=lTrgtClm
        
        '*** それ以外（フィルタ設定） ***
        Case Else:
            If InStr(sFilterKeyword, sDELIMITER) Then
                Dim vBtnTexts As Variant
                vBtnTexts = Split(sFilterKeyword, sDELIMITER)
                
                '空白セル抽出のために成型する
                Dim lIdx As Long
                For lIdx = LBound(vBtnTexts) To UBound(vBtnTexts)
                    If vBtnTexts(lIdx) = "" Then
                        vBtnTexts(lIdx) = "="
                    Else
                        'Do Nothing
                    End If
                Next lIdx
                
                'フィルタリング実行
                rFltTrgtRng.AutoFilter _
                    Field:=lTrgtClm, _
                    Criteria1:=vBtnTexts, _
                    Operator:=xlFilterValues
            Else
                'フィルタリング実行
                rFltTrgtRng.AutoFilter _
                    Field:=lTrgtClm, _
                    Criteria1:=sFilterKeyword, _
                    Operator:=xlFilterValues
            End If
    End Select
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

