' グローバル関数定義
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Global MSGFILENAME As String        '　メッセージファイルネーム
Global MYCALL As String             '　コールサイン
Global HOMERBBS As String           '  ＦＷＤ相手のコールサイン
Global LINK As String               '  リンク方法
Global BAUD As String               '  ボーレート
Global MSGDIR As String             '  メッセージ格納ディレクトリ
Global EDITOR As String             '  起動エディタ名
Global PORT As Integer              '　通信ポート（１から指定）
Global MODE As String               ' 動作モード(ex.FWD,AUTO)
Global MAILHEADER As String         ' 電子メールヘッダーの付加
Global COM_MONITOR As String        ' 通信モニタの表示
Global MYNAME As String             ' 使用者名
Global USERDIR As String            ' ユーザー設定ファイル格納ディレクトリ
Global XFLOWFLAG As Integer         ' ソフトウェアフロー制御フラグ
'
Global SYSDIR As String              '  カレントディレクトリ
Global VERSION As String             '  バージョン
Global WRITEMODE As String           '  書き込みモード　Ｐ　ｏｒ　Ｂ
Global WRITE_TFL As String
Global WRITE_ATBBS As String
'
Global TOFIELDCOL As Integer
Global ATBBSCOL As Integer
Global CALLCOL As Integer
Global TITLECOL As Integer
'
Global MESSAGERECORDNUM As Integer
Global DEBUGFLAG As Integer
'
Global MESSAGE() As MSGD2
'
Global FONTNAMELISTMENU As String * 40
Global FONTSIZELISTMENU As Single
Global FONTNAMETEXTVIEW As String * 40
Global FONTSIZETEXTVIEW As Single
'
Global DOWNSETMODE As String
Global SERVERSETMODE As String
'
Global TIMEZONE As String

' リードコマンド
Global READCMD As String

'タイマーアウトフラグ
Global TIMEOUTFLAG As Integer
Global TIMECOUNT As Integer

' 固定文字列
Global Const DISC$ = "*** DISCONNECTED"

' ファイル管理情報ロードフラグ
Global FILELOADFLAG As Integer      ' FILE.IDXメモリ取り込み済みフラグ
Global FILECHANGEFLAG As Integer    ' FILE.IDX内容書換えフラグ

' ファイルダウンロード時のダイアログ結果
Global FILEEX As Integer

' オートパイロット設定
Global AUTOTOFIELD() As String
Global AUTOATBBS() As String
Global AUTOCALL() As String

'オートパイロット文字列検索位置
Global AUTOPOSTOFIELD As Integer
Global AUTOPOSATBBS As Integer
Global AUTOPOSCALL As Integer
Global AUTOPOSTITLE As Integer

' 通信起動時間の設定
Global TIMERENABLE As String
Global TIMESTARTHOUR As String * 2
Global TIMEENDHOUR As String * 2
Global TIMEMINUTES As String * 2
Global TIMERCOM As Integer
Global TIMERSET() As String * 5
Global TIMERMODE As Integer
Global TIMERINTERVAL As Integer

' 環境設定ファイルファイル名
Global INIFILE As String

' サーバー設定
Global SERVERENABLE As String           'サーバーのON/OFF
Global SERVERCYCLE As String * 2        ' サーバー起動サイクル
Global SERVERFILE() As String * 40      ' サーバーファイル名配列
Global SERVERTOFIELD() As String * 10   ' サーバーToField
Global SERVERATBBS() As String * 10     ' サーバー@BBS
Global SERVERCALL() As String * 10      ' サーバーFrom
Global SERVERMSGNUM As Integer          ' サーバー記録メッセージ番号
Global SVNAME$(), SVTFL$(), SVATBBS$(), SVCALL$()
                                        ' サーバー設定項目

' タイムアウト時間
Global RCVWAITTIME As Integer
Global SENDWAITTIME As Integer

' 送信ディレイ
Global SENDDELAY As Long

' データベースファイル番号
Global DATABASEFILENUM As Integer

' Windows状況
Global R_MODE As Integer
Global WINDOW_MODE As Integer

' 言語モード
Global L_MODE As Integer

' 通信ポートオープンフラグ
Global COMFLAG As Integer

' ＩＭＥモード
Global IMEVALUE As Integer

' DJL-MB Multi DIr
Global DJLMAILDIR As String
Global DJLMULTIDIR As Integer
Global DJLSPOOLDIR As String

' ダウンロードフラグ
Global DOWNLOAD As Integer

' プロンプト
Global PROMPT1 As String

' ターミナルソフト使用フラグ
Global TERMFLAG As Integer

' タイトル入力プロンプト
Global TITLEINPUT As String

' スプーラディレクトリ
Global SPOOLDIR As String
Global SPOOLUP As Integer

' 通信モニタの表示
Global MONITORON As Integer

' 引用符
Global RE_MARK As String

' リスト表示カスタム設定
Global MENUSET$(100, 3)
Global MENUSETINDEX As Integer

' オートパイロット制限値
Global TIMELIMIT As Long
Global RCVLIMIT As Long
Global RCVCOUNT As Long

' 初回起動時のフラグ
Global FIRST As Integer

' LaPutaモードフラグ
Global LAPUTAMODE As Integer
Global NEWSGROUP As String

' 外部テキストエディタ作成ファイル名
Global EXTFILENAME As String

' HTMLエディタ
Global HTMLEDITOR As String

' ウェッブブラウザ
Global WEBBROWSER As String

' HTMLフラグ
Global HTMLFLAG As Integer

' Windowの状態
Global WINDOW_STATE As Integer

