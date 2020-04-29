Option Explicit

Sub ReadWinPos (f As Form, fn As String)
    Dim w%, h%, l%, t%

    Trace "ReadWinPos"

    w% = GetPrivateProfileInt("Window", fn + "_Width", 9600, INIFILE)
    h% = GetPrivateProfileInt("Window", fn + "_Height", 6000, INIFILE)
    l% = GetPrivateProfileInt("Window", fn + "_left", 0, INIFILE)
    t% = GetPrivateProfileInt("Window", fn + "_Top", 0, INIFILE)

    f.Width = w%
    f.Height = h%
    f.Left = l%
    f.Top = t%
End Sub

Sub SPGNEWS_End ()
    Dim i%, r%, a%
    Dim s$, s1$

    If FileExist(USERDIR + "LOG.TXT") = True Then
        If FileLen(USERDIR + "LOG.TXT") > 1000000 Then
            a% = MsgBox("ログ情報が大きくなりすぎました。削除しますか？", MB_ICONQUESTION Or MB_YESNO, "SPG-NEWS")
            If a% = IDYES Then
                Kill USERDIR + "LOG.TXT"
            End If
        End If
    End If

    s$ = "データベースの更新を行いますか？"
    s1$ = "データベース更新中"

    If FILECHANGEFLAG = True Then
'        If MODE <> "DJLMB" Then
            FileCSVSave
'        End If
        If MsgBox(s$, MB_ICONQUESTION Or MB_YESNO) = IDYES Then
            screen.MousePointer = 11
            frmInfo!lblInfo.Caption = s1$
            frmInfo.Show
            If FileExist(USERDIR + "TOFIELD.DAT") = True Then
                Kill USERDIR + "TOFIELD.DAT"
            End If
            If FileExist(USERDIR + "ATBBS.DAT") = True Then
                Kill USERDIR + "ATBBS.DAT"
            End If
            If FileExist(USERDIR + "CALL.DAT") = True Then
                Kill USERDIR + "CALL.DAT"
            End If
            If FileExist(USERDIR + "BID.DAT") = True Then
                Kill USERDIR + "BID.DAT"
            End If
            For i% = UBound(MESSAGE) To 0 Step -1
                If Mid$(MESSAGE(i).stat, 3, 1) <> "K" Then
                    If Left$(MESSAGE(i).stat, 1) <> "P" Then
                        r% = FileDataBaseAdd(USERDIR + "TOFIELD.DAT", Trim$(MESSAGE(i).field))
                    End If
                    DoEvents
                    If Left$(MESSAGE(i).stat, 1) <> "P" Then
                        r% = FileDataBaseAdd(USERDIR + "ATBBS.DAT", Trim$(MESSAGE(i).atbbs))
                    End If
                    DoEvents
                    r% = FileDataBaseAdd(USERDIR + "CALL.DAT", Trim$(MESSAGE(i).Call))
                    If MESSAGE(i).bid <> "" Then
                        r% = FileDataBaseAdd(USERDIR + "BID.DAT", Trim$(MESSAGE(i).bid))
                    End If
                End If
                DoEvents
            Next i%
        End If
        screen.MousePointer = 0
        Unload frmInfo
    End If

    ' メインウィンドゥがアイコン化されているときは座標の記録を行わない
    If frmMain.WindowState <> 1 Then
        WrtWinPos frmMain, "Main"
    End If
    End
End Sub

Sub SPGNEWSInit ()
    Dim Filenum As Integer
    
    Trace "SPGNEWSInit"

    SYSDIR = CurDir$ + "\"
    VERSION = "(Build 240)"
    FILELOADFLAG = False
    FILECHANGEFLAG = False
    COMFLAG = False
    TERMFLAG = False
    
    ' INIファイル設定
    INIFILE = SYSDIR + "SPGNEWS.INI"
    If Command$ = "" Then
        INIFILE = SYSDIR + "SPGNEWS.INI"
    Else
        INIFILE = Command$
    End If

    ConfigRead
End Sub

Sub Splash ()
    ' オープニング画面表示
    frmOpen.Show 0
    TIMEOUTFLAG = 0
    frmTimer.Enabled = True
    frmTimer!tmrTimer.Interval = 3000   ' 3秒間
    frmTimer!tmrTimer.Enabled = True     'タイマースタート
    Do While (TIMEOUTFLAG = 0)
        DoEvents
    Loop
    frmTimer!tmrTimer.Enabled = False
    TIMEOUTFLAG = 0
End Sub

'
' SPG-NEWS for Windows起動モジュール
'
Sub Main ()
    Dim hModule As Integer
    Dim count As Integer
    Dim SoundName As String * 80

    DEBUGFLAG = False
    
    Trace "Main"

    ' 既に「SPG-NEWS」が起動しているかチェックする

    hModule = GetModuleHandle("SPGNEWS.EXE")
    count = GetModuleUsage(hModule)
    If count > 1 Then
'        MsgBox "SPG-NEWSは複数起動できません"
'        End
    End If
    

    If UCase$(Command$) = "E" Then
        L_MODE = 1
    Else
        L_MODE = 0
    End If

    'Debug
'    L_MODE = 1

    SPGNEWSInit
    FileListMake "!K"
    frmPanel.Show
    ReadWinPos frmMain, "Main"
    frmMain.Show

    frmMain.Caption = "SPG-NEWS for Windows [ " + MYCALL + " ］"
    
    If L_MODE = 0 Then
        frmMain!pnl3DStatBar1.Caption = "非通信中"
    Else
        frmMain!pnl3DStatBar1.Caption = ""
    End If
    frmMain!pnl3DTime.Caption = Format$(Now, "yy/mm/dd hh:mm")
    TIMERCOM = False
    frmTimer!tmrScadule.Enabled = True
End Sub

Sub WrtWinPos (f As Form, formname As String)
    Dim a%
    Dim l$, t$, w$, h$
          
    Trace "WrtWinPos"
    
    l$ = Format$(f.Left) + Chr$(0)
    t$ = Format$(f.Top) + Chr$(0)
    w$ = Format$(f.Width) + Chr$(0)
    h$ = Format$(f.Height) + Chr$(0)
    
    a% = WritePrivateProfileString("Window", formname + "_left", l$, INIFILE)
    a% = WritePrivateProfileString("Window", formname + "_Top", t$, INIFILE)
    a% = WritePrivateProfileString("Window", formname + "_Width", w$, INIFILE)
    a% = WritePrivateProfileString("Window", formname + "_Height", h$, INIFILE)
End Sub

