	#NoEnv							; 通常、値が割り当てられていない変数名を参照しようとしたとき、システムの環境変数に同名の変数がないかを調べ、
									; もし存在すればその環境変数の値が参照される。スクリプト中に #NoEnv を記述することにより、この動作を無効化できる。
;	#Warn							; Enable warnings to assist with detecting common errors.
	SendMode Input					; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。
;	SetWorkingDir %A_ScriptDir%		; スクリプトの作業ディレクトリを本スクリプトの格納ディレクトリに変更。
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え

;[参考URL]
;	https://sites.google.com/site/autohotkeyjp/home
;		無変換）vk1Dsc07B
;		変換）	vk1Csc079
;		+）		Shift
;		^）		Control
;		!）		Alt
;		#）		Windowsロゴキー
;	http://ahkwiki.net/Top

;* ***************************************************************
;* Keys
;* ***************************************************************
;*** Global ***
	#1::
		sExePath = "C:\prg_exe\Vim\gvim.exe"
		sFilePath = "%A_MyDocuments%\Dropbox\100_Documents\900_【その他】\100_ToDo.txt"
		StartProgramAndActivate( sExePath, sFilePath )
		return
	#2::
		sFilePath = "%A_MyDocuments%\Dropbox\100_Documents\141_【生活】＜衣食住＞家計\100_家計簿.xlsm"
		StartProgramAndActivate( "", sFilePath )
		return
	#3::
		sExePath = "C:\prg_exe\Vim\gvim.exe"
		sFilePath = "%A_MyDocuments%\Dropbox\100_Documents\132_【生活】＜趣味＞音楽\音楽ストック.txt"
		StartProgramAndActivate( sExePath, sFilePath )
		return
	#4::
		sExePath = "C:\prg_exe\Vim\gvim.exe"
		sFilePath = "%A_MyDocuments%\Dropbox\100_Documents\900_【その他】\999_その他.txt"
		StartProgramAndActivate( sExePath, sFilePath )
		return
	#\::
		sExePath = "C:\prg_exe\Vim\gvim.exe"
		sFilePath = "%A_ScriptFullPath%"
		StartProgramAndActivate( sExePath, sFilePath )
		return
	#F1::
		sExePath = "C:\prg_exe\Vim\gvim.exe"
		sFilePath = "%A_Desktop%\temp.txt"
		StartProgramAndActivate( sExePath, sFilePath )
		return
	#F2::
		sExePath = "C:\codes\vbs\500_CreateExcelFile.vbs"
		sFilePath = "%A_Desktop%\temp.xlsm"
		StartProgramAndActivate( sExePath, sFilePath )
		return
	#F5::Run "C:\prg_exe\Rapture\rapture.exe"
	#F6::RunSuppressMultiStart( "C:\prg_exe\cCalc\cCalc.exe", "" )
	Pause::
		WinSet, AlwaysOnTop, TOGGLE, A
		MsgBox, 0x43000, ウィンドウ最前面切り替え, アクティブウィンドウ最前面化の有効/無効を切り替えます, 5
		Return

;*** Software local ***
	#IfWinActive ahk_exe EXCEL.EXE
		;F1ヘルプ無効化
		F1::return
	#IfWinActive
	
	#IfWinActive AHK_Exe kinza.exe
		;The Great Suspender 用
		F8::^+s
		F9::^+u
		;Shift＋クリックで新規タブ（バックグラウンド）で開く
		+LButton::Send, ^{LButton}
		;Ctrl＋クリックで新規タブで開く
		^LButton::Send, +^{LButton}
		+^i::
			Send, ^c
			Sleep 100
			Send, !d
			Sleep 100
			Send, ^v
			Sleep 100
			Send, {Home}
			Sleep 100
			Send, ^{Right}
			Sleep 100
			Send, {Delete}
			Send, {Delete}
			Sleep 100
			Send, {Enter}
			return
	#IfWinActive
	
	#IfWinActive ahk_class MPC-BE
		]::Send, {Space}
	#IfWinActive

;* ***************************************************************
;* Functions
;* ***************************************************************
	RunSuppressMultiStart( path, argument )
	{
		IfInString, path, \
		{
			Loop, Parse, path , \
			{
				sFileName = %A_LoopField%
			}
			Process, Exist, % sFileName
			If ErrorLevel<>0
			{
				WinActivate,ahk_pid %ErrorLevel%
			}
			else
			{
				Run % path . " " . argument
			}
		}
		else
		{
			MsgBox argument error!
		}
		return
	}
	
	WinSizeChange( size, maxwinx, maxwiny )
	{
		if size = up
		{
			WinGetActiveStats, A, WinWidth, WinHeight, WinX, WinY
			if ( WinX = maxwinx && WinY = maxwiny )
			{
				WinMaximize
			}
			else
			{
				WinMaximize
			}
		}
		else if size = down
		{
			WinGetActiveStats, A, WinWidth, WinHeight, WinX, WinY
			if ( WinX = maxwinx && WinY = maxwiny )
			{
				WinRestore
			}
			else
			{
				WinMinimize
			}
		}
		else if size = max
		{
			WinMaximize
		}
		else if size = restore
		{
			WinRestore
		}
		else if size = min
		{
			WinMinimize
		}
		else
		{
			MsgBox "[error] please select up / down / max / restore / min."
		}
		return
	}
	
	; 既定のショートカットキーとの干渉によりプログラム起動後に
	; ウィンドウがアクティベートされないことがある。(※)
	; 上記問題を対処するため、本関数ではプログラム起動後に
	; ウィンドウをアクティベートする処理を実行する。
	; 
	; (※)例
	; 「Windows キー + 1」はタスクパーに１つ目にピン止め
	; されているプログラムをアクティベートするショートカットキーで
	; あるため、Run 関数を使用してそのまま実行すると、非アクティブ
	; 状態でプログラムが起動してしまう。
	StartProgramAndActivate( sExePath, sFilePath )
	{
		IfInString, sFilePath, \
		{
			;*** extract file name ***
			Loop, Parse, sFilePath , \
			{
				sFileName = %A_LoopField%
			}
			StringReplace, sFileName, sFileName, ", , All
			
			;*** start program ***
			If ( sExePath == "" )
			{
				;MsgBox A ;for debug
				Run, %sFilePath%
			}
			else
			{
				;MsgBox B ;for debug
				Run, %sExePath% %sFilePath%
			}
			
			;*** activate started program ***
			SetTitleMatchMode, 2
			WinWait, %sFileName%
			WinActivate, %sFileName%
			
			;MsgBox %sExePath% ;for debug
			;MsgBox %sFilePath% ;for debug
			;MsgBox %sFileName% ;for debug
		}
		else
		{
			MsgBox argument error!
		}
		return
	}
	
