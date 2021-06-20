Attribute VB_Name = "Mng_ExcelFile"
Option Explicit
 
' excel file manage library v1.01

'************************************************************
'* 構造体定義
'************************************************************
Type T_EXCEL_FILE_INFO
    wTrgtBook As Workbook
    sFilePath As String
    bIsAlreadyOpen As Boolean
End Type

'************************************************************
'* モジュール内 変数定義
'************************************************************
Dim gatExcelFileInfo() As T_EXCEL_FILE_INFO

'************************************************************
'* 関数定義
'************************************************************
' ==================================================================
' = 概要    初期化処理
' = 引数    なし
' = 戻値    なし
' = 依存    なし
' = 所属    Mng_ExcelFile.bas
' ==================================================================
Public Function ExcelFileInfoInit()
    Dim atExcelFileInfoInit() As T_EXCEL_FILE_INFO
    gatExcelFileInfo = atExcelFileInfoInit '初期化
End Function

' ==================================================================
' = 概要    エクセルファイルオープン
' = 引数    なし
' = 戻値    なし
' = 依存    なし
' = 所属    Mng_ExcelFile.bas
' ==================================================================
Public Function ExcelFileOpen( _
    ByVal sTrgtBookPath As String _
) As Workbook
    Dim wTrgtBook As Workbook
    Dim bIsAlreadyOpen As Boolean
    Dim lExcelFileIdx As Long
 
    '配列再定義
    If Sgn(gatExcelFileInfo) = 0 Then '未初期化配列
        lExcelFileIdx = 0
    Else '要素数１以上の配列
        lExcelFileIdx = UBound(gatExcelFileInfo) + 1
    End If
    ReDim Preserve gatExcelFileInfo(lExcelFileIdx)
 
    'ブック存在チェック
    If Dir(sTrgtBookPath) = "" Then
        MsgBox "ブック「" & sTrgtBookPath & "」" & vbNewLine & "が存在しません。", vbExclamation
        Exit Function
    Else
        'Do Nothing
    End If
 
    'ブックが既に開かれているかチェック
    bIsAlreadyOpen = False
    For Each wTrgtBook In Workbooks
        If wTrgtBook.Name = Dir(sTrgtBookPath) Then
            bIsAlreadyOpen = True
        Else
            'Do Nothing
        End If
    Next wTrgtBook
 
    'ファイルオープン
    If bIsAlreadyOpen = True Then
        Set wTrgtBook = Workbooks(Dir(sTrgtBookPath))
    Else
        Set wTrgtBook = Workbooks.Open(sTrgtBookPath)
    End If
 
    'ブック情報格納
    gatExcelFileInfo(lExcelFileIdx).bIsAlreadyOpen = bIsAlreadyOpen
    gatExcelFileInfo(lExcelFileIdx).sFilePath = sTrgtBookPath
    Set gatExcelFileInfo(lExcelFileIdx).wTrgtBook = wTrgtBook
 
    Set ExcelFileOpen = wTrgtBook
End Function

' ==================================================================
' = 概要    エクセルファイルクローズ
' = 引数    なし
' = 戻値    なし
' = 依存    なし
' = 所属    Mng_ExcelFile.bas
' ==================================================================
Public Function ExcelFileClose( _
    ByVal wTrgtBook As Workbook, _
    ByVal bIsSave As Boolean _
)
    Dim lExcelFileIdx As Long
    Dim lArrayIdx As Long
    Dim atExcelFileInfoInit() As T_EXCEL_FILE_INFO
 
    '開いているブックを検索
    For lExcelFileIdx = 0 To UBound(gatExcelFileInfo)
        If gatExcelFileInfo(lExcelFileIdx).sFilePath = wTrgtBook.FullName Then
            Exit For
        Else
            'Do Nothing
        End If
    Next lExcelFileIdx
 
    Debug.Assert UBound(gatExcelFileInfo) >= lExcelFileIdx
 
    '閉じる
    If gatExcelFileInfo(lExcelFileIdx).bIsAlreadyOpen = False Then
        If bIsSave = True Then
            gatExcelFileInfo(lExcelFileIdx).wTrgtBook.Close SaveChanges:=True
        Else
            gatExcelFileInfo(lExcelFileIdx).wTrgtBook.Close SaveChanges:=False
        End If
    Else
        'Do Nothing
    End If
 
    '配列要素削除
    If UBound(gatExcelFileInfo) = 0 Then
        gatExcelFileInfo = atExcelFileInfoInit
    Else
        For lArrayIdx = lExcelFileIdx To UBound(gatExcelFileInfo) - 1
            gatExcelFileInfo(lArrayIdx).bIsAlreadyOpen = gatExcelFileInfo(lArrayIdx + 1).bIsAlreadyOpen
            gatExcelFileInfo(lArrayIdx).sFilePath = gatExcelFileInfo(lArrayIdx + 1).sFilePath
            Set gatExcelFileInfo(lArrayIdx).wTrgtBook = gatExcelFileInfo(lArrayIdx + 1).wTrgtBook
        Next lArrayIdx
        ReDim Preserve gatExcelFileInfo(UBound(gatExcelFileInfo) - 1)
    End If
End Function

