'ファイル処理関数
' Copyright (c) 1993-1997 by JO1SPG/T.KIZAWA
Option Explicit

Sub FileAdd (fname As String, s As String)
    Dim FileNum As Integer

    FileNum = FreeFile
    Open fname For Append As FileNum
    Print #FileNum, s
    Close FileNum
End Sub

'
' 終了時のファイル書き込み
'
Sub FileCSVSave ()
    Dim FileNum As Integer
    Dim Count As Integer
    Dim i, c As Integer
    Dim MaxNum As Integer
    Dim w%, h%, x%, lst$, kfile$, kcmd$

    Trace "FileCSVSave"

'    On Error GoTo ERRPROC_CSV_S

    frmCount.Show 0
    If L_MODE = 0 Then
        frmCount.Caption = "SPG-NEWS メッセージデータ書き込み中"
    Else
        frmCount.Caption = "Writing Message data"
    End If
    Screen.MousePointer = 11
    c = 0
    w% = frmCount.picCount.Width
    h% = frmCount.picCount.Height
    frmCount.picCount.Line (w% * .25, 0)-(w% * .25, h%), RGB(255, 0, 0)
    frmCount.picCount.Line (w% * .5, 0)-(w% * .5, h%), RGB(255, 0, 0)
    frmCount.picCount.Line (w% * .75, 0)-(w% * .75, h%), RGB(255, 0, 0)
    Count = 0
    FileNum = FreeFile
    If MODE = "AUTO" Or MODE = "FWD" Then
        FileCopy MSGDIR + "\" + "FILE.IDX", MSGDIR + "\" + "FILE.BAK"
1011    Open MSGDIR + "\" + "FILE.IDX" For Output As FileNum
    End If
    MaxNum = UBound(MESSAGE)
    lst$ = ""
    If MaxNum > 0 Then
        For i = 0 To MaxNum
            If MESSAGE(i).fname <> "" Then
                ' メッセージ番号
                lst$ = MESSAGE(i).fname + Chr$(9)
                ' メッセージタイプ
                lst$ = lst$ + "TXT" + Chr$(9)
                ' 保存方法
                lst$ = lst$ + "NON" + Chr$(9)
                ' ステータス
                lst$ = lst$ + MESSAGE(i).stat + Chr$(9)
                ' ファイルサイズ
                lst$ = lst$ + MESSAGE(i).size + Chr$(9)
                ' フィールド
                lst$ = lst$ + MESSAGE(i).field + Chr$(9)
                ' ＠ＢＢＳ
                lst$ = lst$ + MESSAGE(i).atbbs + Chr$(9)
                ' コールサイン
                lst$ = lst$ + MESSAGE(i).Call + Chr$(9)
                ' 受信日付
                lst$ = lst$ + MESSAGE(i).date + Chr$(9)
                ' 受信時刻
                lst$ = lst$ + MESSAGE(i).time + Chr$(9)
                ' Bulletin-ID
                lst$ = lst$ + MESSAGE(i).bid + Chr$(9)
                ' タイトル
                lst$ = lst$ + MESSAGE(i).Title
                Select Case MODE
                Case "AUTO", "FWD"
                    ' ファイル追加
                    If Mid$(MESSAGE(i).stat, 3, 1) <> "K" Then
                        Print #FileNum, lst$
                    Else
                        FileKill (Val(MESSAGE(i).fname))
                    End If
                Case "DJLMB"
                    If Mid$(MESSAGE(i).stat, 3, 1) = "K" Then
                        ' DJL-MB Ｋコマンド発行
                        kfile$ = "K" + Left$(Time$, 2) + Mid$(Time$, 4, 2) + Right$(Time$, 2) + ".TXT"
                        Open DJLSPOOLDIR + "\" + kfile$ For Append As FileNum
                        kcmd$ = "K " + MESSAGE(i).fname
                        Print #FileNum, kcmd$
                        Close #FileNum
                    End If
                End Select
                x% = (i / MaxNum) * w%
                frmCount.picCount.Line (0, 0)-(x%, h%), , BF
                lst$ = ""
                DoEvents
            End If
        Next i
    End If
    If MODE = "AUTO" Or MODE = "FWD" Then
        Close FileNum
    End If
    frmCount.picCount.Line (0, 0)-(w%, h%), , BF
    Screen.MousePointer = 0
    frmCount.Hide
End Sub

'
' データベース追加
'
Function FileDataBaseAdd (fname As String, s As String)
    
    Dim s1 As String
    Dim FileNum As Integer

    Trace "FileDataBaseAdd"
    
    If FileExist(fname) = False Then
        FileNum = FreeFile
        Open fname For Output As FileNum
        Close #FileNum
    End If
        
' データベースファイルオープン
    FileNum = FreeFile
    Open fname For Input As FileNum

' 検索
    Do While Not EOF(FileNum)
        Line Input #FileNum, s1
        If s = s1 Then                 ' 重複
            Close #FileNum
            FileDataBaseAdd = -1
            Exit Function
        End If
    Loop
    Close #FileNum

' 追加
    FileNum = FreeFile
    Open fname For Append As FileNum
        Print #FileNum, s
    Close #FileNum
    FileDataBaseAdd = 0
End Function

'
' DJL-MB MAIL.MB読み込み
'
' f:True  オープニング画面を表示する
'   False オープニング画面を表示しない
Sub FileDJLMsgLoad (f As Integer)
    Dim FileNum As Integer
    Dim Count As Integer
    Dim i, r As Integer
    Dim l, c As Long
    Dim p!, Percent$
    Dim w%, h%, a$, n%
    Dim mbsize&
    Dim mbnum%
    Dim djlm As DJLMSG

    ReDim MESSAGE(0) As MSGD2

'    On Error GoTo FileDJLMsgLoad_err

    Screen.MousePointer = 11
    l = 0
    c = 0
    w% = frmOpen.picCount.Width
    h% = frmOpen.picCount.Height
    If f = True Then
        frmOpen.picCount.Line (w% * .25, 0)-(w% * .25, h%), RGB(255, 0, 0)
        frmOpen.picCount.Line (w% * .5, 0)-(w% * .5, h%), RGB(255, 0, 0)
        frmOpen.picCount.Line (w% * .75, 0)-(w% * .75, h%), RGB(255, 0, 0)
    End If

    mbsize& = GetDJLMBSize()
    mbnum% = mbsize& / 512

    If FileExist(DJLMAILDIR + "\" + "MAIL.MB") = False Then
        Screen.MousePointer = 0
        r = MsgBox("DJL-MBのMAIL.MBがないか、現在MAIL.MBをアクセスできる状況ではありません。MAIL.MBのディレクトリを設定していない場合は設定しますか?", MB_ICONEXCLAMATION Or MB_YESNO)
        If r = IDYES Then
            SetUpConfig (True)
            End
        Else
            If f = True Then
                End
            Else
                Exit Sub
            End If
        End If
    End If
    
    n% = 1
    Count = mbnum% - 1
    FileNum = FreeFile
    ReDim Preserve MESSAGE(mbnum% - 1)
    Open DJLMAILDIR + "\" + "MAIL.MB" For Random Access Read As FileNum Len = 512
    For i = mbnum% To 2 Step -1
        Get #FileNum, i, djlm
        MESSAGE(Count).fname = Format$(Asc(Right$(djlm.num, 1)) * 256 + Asc(Left$(djlm.num, 1)))
        MESSAGE(Count).mediatype = "TXT"
        MESSAGE(Count).savemode = "NON"
        MESSAGE(Count).stat = djlm.stat1 + "F"
        MESSAGE(Count).size = Format$(Asc(Right$(djlm.size, 1)) * 256 + Asc(Left$(djlm.size, 1)))
        MESSAGE(Count).field = LNullCut(djlm.field)
        MESSAGE(Count).atbbs = LNullCut(djlm.atbbs)
        MESSAGE(Count).Call = LNullCut(djlm.call1)
        MESSAGE(Count).date = LNullCut(djlm.rdate)
        MESSAGE(Count).time = LNullCut(djlm.rtime)
        MESSAGE(Count).bid = LNullCut(djlm.bid)
        MESSAGE(Count).Title = LNullCut(djlm.Title)
        ' Killフラグが立っているものは読み込まない
        If (Asc(djlm.stat2) And &H4) <> 4 Then
            Count = Count - 1
        End If
        p! = n% / mbnum%
        If f = True Then
            frmOpen.picCount.Line (0, 0)-(frmOpen.picCount.Width * p!, frmOpen.picCount.Height), , BF
        End If
        DoEvents
        n% = n% + 1
        If n% = 2520 Then
            MsgBox "これ以上メッセージを読み込むことができません" + Chr$(13) + Chr$(10), MB_ICONEXCLAMATION Or MB_OK
            Exit For
        End If
    Next i
    Close #FileNum
    
    If f = True Then
        frmOpen.picCount.Line (0, 0)-(w%, 615), , BF
    End If
    Screen.MousePointer = 0
    FILELOADFLAG = True
    frmOpen.Hide
    Exit Sub

FileDJLMsgLoad_err:
    Screen.MousePointer = 0
    Beep
    r = MsgBox("DJL-MBのメッセージを格納するディレクトリがありません", MB_ICONEXCLAMATION)
    SetUpConfig (True)
    End
End Sub

'　ファイルの存在チェック
'  あり＝Ｔｒｕｅ
'  なし＝Ｆａｌｓｅ
'
Function FileExist (s$) As Integer
    Dim FileNum As Integer

    On Error GoTo FE

    FileNum = FreeFile
    Open s$ For Input Access Read As FileNum
    Close FileNum
    FileExist = True
    Exit Function

FE:
    Close FileNum
    FileExist = False
    Exit Function
End Function

'
' メッセージを削除する
'
Sub FileKill (num&)
    Dim i As Integer
    Dim FileNum As Integer

    On Error GoTo ERR_KILL
    
    i = DataGetNum(num&)
    If i = -1 Then
        Exit Sub
    End If
    Kill MSGDIR + "\" + Trim$(Str$(num))
    Exit Sub
ERR_KILL:
    Resume Next
End Sub

' ファイル一行読み込み
'
Function FileLineRead (fnum As Integer) As String
    Trace "FileLineRead"
    While Not EOF(fnum)
                
    Wend
    FileLineRead = Null
End Function

Function FileListLineMake (i As Integer) As String
    Dim s$, s1$, s2$, s3$, s4$, s5$, s6$, lst$
    
    Trace "FileListLineMake"
    
    s3$ = MESSAGE(i).field
    s4$ = MESSAGE(i).atbbs
    s5$ = MESSAGE(i).Call
    s$ = MESSAGE(i).fname
    s1$ = MESSAGE(i).stat
    s2$ = MESSAGE(i).size
    s6$ = MESSAGE(i).Title
    lst$ = Right$("000000" + s$, 5) + " "
    lst$ = lst$ + Left$(s1$ + "   ", 3) + " "
    lst$ = lst$ + Right$("      " + s2$, 5) + " "
    lst$ = lst$ + Left$(s3$ + "      ", 6)
    lst$ = lst$ + "@" + Left$(s4$ + "      ", 6) + " "
    lst$ = lst$ + Left$(s5$ + "      ", 6) + " "
    lst$ = lst$ + Left$(MESSAGE(i).date, 2) + "/" + Mid$(MESSAGE(i).date, 3, 2) + "/" + Mid$(MESSAGE(i).date, 5, 2) + " "
    lst$ = lst$ + Left$(MESSAGE(i).time, 2) + ":" + Mid$(MESSAGE(i).time, 3, 2) + " "
    lst$ = lst$ + s6$
    FileListLineMake = lst$
End Function

'
' ファイルリスト作成
'
Sub FileListMake (ss$)
    Dim MaxNum As Integer
    Dim MaxNum1 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim row1 As Integer
    Dim st, e, stp As Integer
    Dim sf$, sa$, sc$, s$, s1$, s2$, s3$, s4$, s5$, s6$, lst$

    If FILELOADFLAG = False Then
'        FileCSVLoad
        FileMsgLoad
    End If
    MaxNum = UBound(MESSAGE)

    ' コンボボックスの検索データ読み出し
    sf$ = UCase$(frmMain!cboToField.Text)
    sa$ = UCase$(frmMain!cboAtbbs.Text)
    sc$ = UCase$(frmMain!cboCall.Text)

    frmMain!lstMenu.Clear

    If L_MODE = 0 Then
        frmMain!Label3.Caption = "Msg#  " + "Stat" + " Size " + "宛先  " + "@BBS    " + "送信者 " + "日付     " + "時刻  " + "題名"
    Else
        frmMain!Label3.Caption = "Msg#  " + "Stat" + " Size " + "To    " + "@BBS    " + "From   " + "Day      " + "Time  " + "Title"
    End If
    j = 1
'    st = MaxNum
'    e = 0
'    If st > e Then
'        stp = -1
'    Else
'        stp = 1
'    End If
    For i = MaxNum To 0 Step -1  'stp
        If StatSelect(MESSAGE(i).stat, ss$) = True Then
            s3$ = MESSAGE(i).field
            s4$ = MESSAGE(i).atbbs
            s5$ = MESSAGE(i).Call
            If s3$ Like sf$ = True Then
                If s4$ Like sa$ = True Or s4$ = "" Then
                    If s5$ Like sc$ = True Then
                        s$ = MESSAGE(i).fname
                        s1$ = MESSAGE(i).stat
                        s2$ = MESSAGE(i).size
                        s6$ = MESSAGE(i).Title
                        lst$ = Right$("000000" + s$, 5) + " " ' + "□ "
                        lst$ = lst$ + Left$(s1$ + "   ", 3) + " "
                        lst$ = lst$ + Right$("      " + s2$, 5) + " "
                        lst$ = lst$ + Left$(s3$ + "      ", 6) + ""
                        lst$ = lst$ + "@" + Left$(s4$ + "      ", 6) + " "
                        lst$ = lst$ + Left$(s5$ + "      ", 6) + " "
                        lst$ = lst$ + Left$(MESSAGE(i).date, 2) + "/" + Mid$(MESSAGE(i).date, 3, 2) + "/" + Mid$(MESSAGE(i).date, 5, 2) + " "
                        lst$ = lst$ + Left$(MESSAGE(i).time, 2) + ":" + Mid$(MESSAGE(i).time, 3, 2) + " "' + ":" + mid$(Message(i).time, 5, 2) + " "
                        lst$ = lst$ + s6$
                        ' リスト追加
                        If s$ <> "" Then
                            frmMain!lstMenu.AddItem lst$ ' ****
                        End If
                        lst$ = ""
                        j = j + 1
                        If j >= 400 Then
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next i
'    Close #FileNum
    frmMain!lstMenu.ListIndex = 0 ' ****
    If frmMain!lstMenu.ListCount <> 0 Then
        frmMain!lstMenu.Selected(0) = True
    End If
End Sub

Sub FileMsgLoad ()
    ' オープニング画面表示
    frmOpen.Show 0
    TIMEOUTFLAG = 0
    frmTimer.Enabled = True
    frmTimer!tmrTimer.Interval = 2000   ' 2秒間
    frmTimer!tmrTimer.Enabled = True     'タイマースタート
    Do While (TIMEOUTFLAG = 0)
        DoEvents
    Loop
    frmTimer!tmrTimer.Enabled = False
    TIMEOUTFLAG = 0
    
    Select Case MODE
    Case "AUTO", "FWD"
        FileCSVLoad (True)
    Case "DJLMB"
        FileDJLMsgLoad (True)
    End Select
End Sub

Function FileNameMake (s$)
    Dim a As Integer

    Select Case MODE
    Case "AUTO", "FWD"
        ' SPG-NEWSモード
        FileNameMake = MSGDIR + "\" + s$
    Case "DJLMB"
        If DJLMULTIDIR = True Then
            a = Val(s$) Mod 10
            FileNameMake = DJLMAILDIR + "\" + "MAIL" + Format$(a) + "\" + s$
        Else
            FileNameMake = DJLMAILDIR + "\" + s$
        End If
    End Select
End Function

' エディタの内容を記録する
'
Sub FileSave (Filename)
    Dim FileNum As Integer
    Dim FileNum1 As Integer
    Dim Contents As String
    Dim s$

    Trace "FileSave"
    
    FileNum = FreeFile

    ' ファイルを開きます。
    Open Filename For Output As FileNum
    ' メモ帳の内容を変数に設定します。
    Contents = frmEdit!txtEdit.Text
    ' 砂時計を表示します。
    Screen.MousePointer = 11
    ' 内容を保存するファイルに書きます。
    Print #FileNum, Contents;

    If FileExist(USERDIR + EXTFILENAME) = True Then
        FileNum1 = FreeFile
        Open USERDIR + EXTFILENAME For Input As FileNum1
        While Not EOF(FileNum1)
            Line Input #FileNum1, s$
            Print #FileNum, s$
        Wend
        Close #FileNum1
        Kill USERDIR + EXTFILENAME
    End If
    Close #FileNum
    ' マウス ポインタをリセットします。
    Screen.MousePointer = 0
End Sub

'
' テキスト作成
'
Sub FileTextMake (msgnum&)
    Dim FileNum As Integer
    Dim num As Integer
    Dim i, j, ttf, rhf As Integer
    Dim Filename As String ' * 80
    Dim Title As String
    Dim Path As String
    Dim root As String
    Dim strline$, t$, d$, c2$, tfl$, fr$, sb$, s$, tt$
    
    On Error GoTo FileTextMake_Error

    Trace "FileTextMake"
    
    Filename = FileNameMake(Trim$(Str$(msgnum)))
    If msgnum = 0 Then
        Exit Sub
    End If

    num = DataGetNum(msgnum&)

    FileNum = FreeFile
    i = 0                               ' ヘッダーカウント
'    Path = "Path="
10  Open Filename For Input As FileNum
'    Do
'        Line Input #Filenum, strline$
'        If Left$(strline$, 1) = "R" Then
'            i = i + 1
'            Path = Path + HeadCut(strline$) + "!"
'            If (i Mod 10) = 0 Then
'                Path = Path + Chr$(13) + Chr$(10) + "     "
'            End If
'            root = strline$
'        Else
'            Exit Do
'        End If
'    Loop
'
'    If i > 0 Then                       '　ルートヘッダーあり
'        t$ = HeadTime(root)             ' 時刻取得
'        d$ = HeadDate(root)             ' 日付取得
'        c2$ = HeadCut2(root)            ' コールサイン取得
'        t$ = "Date: " + Getrfctime(d$, t$)
'        tfl$ = "To: " + MESSAGE(num).field + "@" + MESSAGE(num).atbbs
'        fr$ = "From: " + MESSAGE(num).Call + "@" + c2$
'        sb$ = "Subject: " + MESSAGE(num).Title
'        s$ = s$ + t$ + Chr$(13) + Chr$(10) + tfl$ + Chr$(13) + Chr$(10) + fr$ + Chr$(13) + Chr$(10)
'        s$ = s$ + sb$ + Chr$(13) + Chr$(10) + Path + Chr$(13) + Chr$(10)
'    End If

'    s$ = s$ + strline$ + Chr$(13) + Chr$(10)
    ttf = False
    rhf = False
    While (Not EOF(FileNum)) Or (strline$ = "/EX")
        Line Input #FileNum, strline$
        i = i + 1
        If strline$ = "" And rhf = False Then
            frmTextView!Label1.Caption = Format$(i - 1)
            rhf = True
        End If
        If InStr(strline$, "Subject: ") = 1 And ttf = False Then 'B226
            tt$ = Mid$(strline$, 10)
            ttf = True
        End If
        s$ = s$ + strline$ + Chr$(13) + Chr$(10)
    Wend
    Close #FileNum

    ' もしSubject:が検出できなかったら既存のタイトルを使う
    If tt$ = "" Then
        tt$ = MESSAGE(num).Title
    End If
    frmTextView.Caption = MESSAGE(num).Title
    frmTextView!txtTextView.Text = HTMLCut(s$)
    t$ = "To: " + MESSAGE(num).field + "@"
    If InStr(MESSAGE(num).atbbs, ".") <> 0 Then
        t$ = t$ + Left$(MESSAGE(num).atbbs, InStr(MESSAGE(num).atbbs, ".") - 1)
    Else
        t$ = t$ + MESSAGE(num).atbbs
    End If
    If Left$(MESSAGE(num).stat, 1) = "P" Then
        frmTextView!mnuReplyBulletin.Enabled = False
    Else
        frmTextView!mnuReplyBulletin.Enabled = True
    End If
    frmTextView!pnlToField.Caption = t$
    frmTextView!pnlFrom.Caption = "From: " + MESSAGE(num).Call
    frmTextView!pnlTitle.Caption = "Subject: " + tt$
    frmTextView!Label2.Caption = ""
    frmTextView.Show
    FileNum = FreeFile
    Open USERDIR + "READhtml.TXT" For Output As FileNum
        Print #FileNum, s$
    Close #FileNum
    Exit Sub

FileTextMake_Error:
    MsgBox "メッセージはありません", MB_ICONEXCLAMATION
    Exit Sub
End Sub

Sub FileTimeKill (sf$, sa$, sc$, dd As Integer)
    Dim i, j, MaxNum As Integer
    Dim days, days1 As Long
    Dim d$, sd$, st$, s3$, s4$, s5$

    If FILELOADFLAG = False Then
'        FileCSVLoad
        FileMsgLoad
    End If
    MaxNum = UBound(MESSAGE)

    j = 1
    For i = 0 To MaxNum
        If InStr(MESSAGE(i).stat, "K") = 0 Then
            s3$ = MESSAGE(i).field
            s4$ = MESSAGE(i).atbbs
            s5$ = MESSAGE(i).Call
            sd$ = MESSAGE(i).date
            st$ = MESSAGE(i).time
            If (sf$ = "*" Or sf$ = s3$) Then
                If (sa$ = "*" Or sa$ = s4$) Then
                    If (sc$ = "*" Or sc$ = s5$) Then
                        d$ = Mid$(Date$, 9, 2) + Mid$(Date$, 1, 2) + Mid$(Date$, 4, 2)
'                        d$ = Mid$(Date$, 3, 2) + Mid$(Date$, 6, 2) + Mid$(Date$, 9, 2)
                        If DayCount(d$) - dd >= DayCount(sd$) Then
                            If Len(MESSAGE(i).stat) = 2 Then
                                MESSAGE(i).stat = MESSAGE(i).stat + "K"
                            Else
                                Mid$(MESSAGE(i).stat, 3, 1) = "K"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

'
' ファイル書き込み（ＳＹＳＯＰ用）
'
Function FileWrite (fname As String, stat As String, field As String, atbbs As String, Title As String, fstat As String, flag As Integer)
    Dim FileNum As Integer  'ファイル番号
    Dim FileNum1 As Integer  'ファイル番号
    Dim Filenum2 As Integer  'ファイル番号
    Dim InFilenum As Integer
    Dim OutFilenum As Integer
    Dim MaxNum As Integer
    Dim s As String
    Dim num&               'メッセージ番号
    Dim linetext As String
    Dim bid As String
    Dim HEAD$, t$, lst$
    Dim a, Count As Integer
    Dim fndjl$, msgid$

    Trace "FileWrite"
    
'    On Error GoTo err_sig

' 大文字変換
    field = UCase$(field)
    atbbs = UCase$(atbbs)

' メッセージ番号取得
    num& = GetPrivateProfileInt("Mail", "Nextnumber", -1, INIFILE)
    If num& = -1 Then
        num& = 1
    End If
    s = Mid$(Str$(num&), 2)
    bid = s + "_" + MYCALL
    HEAD$ = "SPG-NEWS for Windows Version 4.5 "

    Select Case MODE
    Case "AUTO", "FWD"
        ' SPG-NEWSオリジナルファイル作成
        FileNum1 = FreeFile
        Open fname For Input As #FileNum1
        Filenum2 = FreeFile
        Open MSGDIR + "\" + s For Output As #Filenum2
        Print #Filenum2, "X-Posting-Front-End: " + HEAD$ + " " + VERSION
        Print #Filenum2, "Date: " + Format$(Now, "ddd, dd mmm yy hh:mm:ss") + " " + TIMEZONE
        If Trim$(atbbs) = "" Then
            Print #Filenum2, "To: " + Trim$(field)
        Else
            Print #Filenum2, "To: " + Trim$(field) + "@" + Trim$(atbbs)
        End If
        Print #Filenum2, "From: " + MYCALL + "@" + HOMERBBS;
        If MYNAME <> "" Then
            Print #Filenum2, " (" + MYNAME + ")"
        Else
            Print #Filenum2, ""
        End If
        msgid$ = "<" + Mid$(Date$, 7, 4) + Mid$(Date$, 1, 2) + Mid$(Date$, 4, 2) + Mid$(Time$, 1, 2) + Mid$(Time$, 4, 2) + Mid$(Time$, 7, 2) + "." + Format$(num&) + "@" + MYCALL + ">"
        Print #Filenum2, "Message-ID: " + msgid$
'        Print #Filenum2, "Organization: SPG SYSTEMS / SPG-NEWS開発部電子メールシステム開発グループ"
'        If LAPUTAMODE = True Then
'            Print #Filenum2, "X-LaPuta-Version: 0.01"
'            Print #Filenum2, "X-LaPuta-Newsgroups: " + LCase$(Trim$(atbbs)) + "." + LCase$(Trim$(field)) + NEWSGROUP
'            If Mid$(stat, 2, 1) = "$" Then
'                Print #Filenum2, "X-LaPuta-MessageID: " + bid
'            End If
'        End If
        Print #Filenum2, "Subject: " + Trim$(Title)
        '
        Line Input #FileNum1, linetext
        If InStr(linetext, "In-reply-to:") <> 0 Then     ' Reply
            Print #Filenum2, linetext
        Else
            Print #Filenum2, ""
            Print #Filenum2, linetext
        End If
        Do While Not EOF(FileNum1)
            Line Input #FileNum1, linetext
            Print #Filenum2, linetext
        Loop
        Close #FileNum1
    
        If flag = True And FileExist(USERDIR + "SIGNATUR.TXT") = True Then
            FileNum1 = FreeFile
            Open USERDIR + "SIGNATUR.TXT" For Input As #FileNum1
            Do While Not EOF(FileNum1)
                Line Input #FileNum1, linetext
                Print #Filenum2, linetext
            Loop
            Close #FileNum1
        End If
'skip_sig:
        Close #Filenum2
    
        Count = UBound(MESSAGE) + 1
        ReDim Preserve MESSAGE(Count)
    
        ' ＣＳＶ形式ファイルバックアップ
        t$ = Chr$(9)
        If Mid$(stat, 2, 1) <> "$" Then
            bid = ""
        End If
        lst$ = s + t$ + "TXT" + t$ + "NON" + t$ + stat + t$ + Trim$(Str$(FileLen(MSGDIR + "\" + s))) + t$ + field + t$ + atbbs + t$ + MYCALL + t$ + Format$(Now, "yymmdd") + t$ + Format$(Now, "hhmmss") + t$ + bid + t$ + Title
        FileNum = FreeFile
        Open MSGDIR + "\" + "FILE.IDX" For Append Lock Write As #FileNum
            Print #FileNum, lst$
        Close #FileNum
        MessageCut lst$, MESSAGE(Count)
    Case "DJLMB"
        ' DJL-MB SPOOLファイル作成
        FileNum1 = FreeFile
        Open fname For Input As #FileNum1
        Filenum2 = FreeFile
        fndjl$ = DJLSPOOLDIR + "\S" + Left$(Time$, 2) + Mid$(Time$, 4, 2) + Right$(Time$, 2)
        Open fndjl$ For Append As #Filenum2
        
        Print #Filenum2, "S" + Left$(stat, 1) + " " + field + " @ " + atbbs + " < " + MYCALL;
        If Mid$(stat$, 2, 1) = "$" Then
            Print #Filenum2, " $" + bid;
        End If
        Print #Filenum2, ""
        Print #Filenum2, Title

        Print #Filenum2, "Posting-Front-End: " + HEAD$ + " " + VERSION
        Print #Filenum2, "Date: " + Format$(Now, "ddd, dd mmm yy hh:mm:ss") + " " + TIMEZONE
        If Trim$(atbbs) = "" Then
            Print #Filenum2, "To: " + Trim$(field)
        Else
            Print #Filenum2, "To: " + Trim$(field) + "@" + Trim$(atbbs)
        End If
        Print #Filenum2, "From: " + MYCALL + "@" + HOMERBBS;
        If MYNAME <> "" Then
            Print #Filenum2, " (" + MYNAME + ")"
        Else
            Print #Filenum2, ""
        End If
        If Mid$(stat, 2, 1) = "$" Then
        ' Print #Filenum2, "X-Bulletin-ID: " + bid
        End If
        Print #Filenum2, "Subject: " + Trim$(Title)
        '
        Line Input #FileNum1, linetext
        If InStr(linetext, "In-reply-to:") <> 0 Then     ' Reply
            Print #Filenum2, linetext
        Else
            Print #Filenum2, ""
            Print #Filenum2, linetext
        End If
        Do While Not EOF(FileNum1)
            Line Input #FileNum1, linetext
            Print #Filenum2, linetext
        Loop
        Close #FileNum1
    
        If flag = True And FileExist(USERDIR + "SIGNATUR.TXT") = True Then
            FileNum1 = FreeFile
            Open USERDIR + "SIGNATUR.TXT" For Input As #FileNum1
            Do While Not EOF(FileNum1)
                Line Input #FileNum1, linetext
                Print #Filenum2, linetext
            Loop
            Close #FileNum1
        End If
'skip_sig1:
        Print #Filenum2, "/EX"
        Close #Filenum2
    End Select

    If MODE = "AUTO" Or MODE = "FWD" Then
        ServerOut num&, stat, field, atbbs, MYCALL, Title, bid
    End If

' データベース更新
    If Mid$(stat, 2, 1) = "$" Then
        a = FileDataBaseAdd(USERDIR + "BID.DAT", Trim$(bid))
    End If
    If Left$(stat, 1) = "P" Then
        If Trim$(atbbs) = "" Then
            a = FileDataBaseAdd(USERDIR + "ADDR.DAT", Trim$(field))
        Else
            a = FileDataBaseAdd(USERDIR + "ADDR.DAT", Trim$(field) + "@" + Trim$(atbbs))
        End If
    Else
        a = FileDataBaseAdd(USERDIR + "TOFIELD.DAT", Trim$(field))
        a = FileDataBaseAdd(USERDIR + "ATBBS.DAT", Trim$(atbbs))
    End If

' メッセージ番号更新
    num& = num& + 1
    If num& > 32767 Then
        num& = 1
    End If
    a = WritePrivateProfileString("Mail", "NextNumber", Format$(num&), INIFILE)
    Exit Function

'err_sig:
'    GoTo skip_sig
End Function

'
' ファイル書き込み（ＦＷＤ受信用）
'
Function FileWrite2 (fname As String, stat As String, field As String, atbbs As String, callsign As String, Title As String, bid As String, fstat As String, flag As Integer)
    Dim FileNum As Integer  'ファイル番号
    Dim FileNum1 As Integer  'ファイル番号
    Dim Filenum2 As Integer  'ファイル番号
    Dim InFilenum As Integer
    Dim OutFilenum As Integer
    Dim MaxNum As Integer
    Dim s As String
    Dim num&                'メッセージ番号
    Dim linetext As String
    Dim a, Count As Integer
    Dim t$, lst$, HEAD$

    Trace "FileWrite2"

' 大文字変換
    field = UCase$(field)
    atbbs = UCase$(atbbs)
    callsign = UCase$(callsign)

' メッセージ番号取得
    num& = GetPrivateProfileInt("Mail", "NextNumber", -1, INIFILE)
    If num& = -1 Then
        num& = 1
    End If
    s = Mid$(Str$(num&), 2)

' ファイル作成
    FileNum1 = FreeFile
    Open fname For Input As FileNum1
    Filenum2 = FreeFile
    Open MSGDIR + "\" + s For Output As Filenum2
'    If flag = True Then
'        HEAD$ = "SPG-NEWS for Windows Version 4.3 "
'        Print #Filenum2, "X-Sender: " + HEAD$ + " " + VERSION
'    End If
    Do While Not EOF(FileNum1)
        Line Input #FileNum1, linetext
        Print #Filenum2, linetext
    Loop
    Close #FileNum1
    Close #Filenum2

    Count = UBound(MESSAGE) + 1
    ReDim Preserve MESSAGE(Count)

' ＣＳＶ形式ファイルバックアップ
    t$ = Chr$(9)
    lst$ = s + t$ + "TXT" + t$ + "NON" + t$ + stat + t$ + Trim$(Str$(FileLen(MSGDIR + "\" + s))) + t$ + field + t$ + atbbs + t$ + callsign + t$ + Format$(Now, "yymmdd") + t$ + Format$(Now, "hhmmss") + t$ + bid + t$ + Title
    FileNum = FreeFile
    Open MSGDIR + "\" + "FILE.IDX" For Append Lock Write As #FileNum
        Print #FileNum, lst$
    Close #FileNum
    MessageCut lst$, MESSAGE(Count)

    ServerOut num, "BF", field, atbbs, callsign, Title, bid

' データベース更新
    If Mid$(stat, 2, 1) = "$" Then
        a = FileDataBaseAdd(USERDIR + "BID.DAT", Trim$(bid))
    End If
    If Left$(stat, 1) = "P" Then
        a = FileDataBaseAdd(USERDIR + "ADDR.DAT", Trim$(field) + " @ " + Trim$(atbbs))
    Else
        a = FileDataBaseAdd(USERDIR + "TOFIELD.DAT", Trim$(field))
        a = FileDataBaseAdd(USERDIR + "ATBBS.DAT", Trim$(atbbs))
        a = FileDataBaseAdd(USERDIR + "CALL.DAT", Trim$(callsign))
    End If

' メッセージ番号更新
    num& = num& + 1
    If num& > 32767 Then
        num& = 1
    End If
    a = WritePrivateProfileString("Mail", "NextNumber", Format$(num&), INIFILE)
End Function

Sub FileWriteDJLMB (num&, stat$, field$, atbbs$, MYCALL, Title$, bid$)
' FileWriteDJLMB num&, stat, field, atbbs, MYCALL, Title, bid
    Dim FileNum1, Filenum2 As Integer
    Dim l$, oname$

    oname$ = "DJLUP.TXT"
    FileNum1 = FreeFile
    Open MSGDIR + "\" + Format$(num&) For Input As #FileNum1
    Filenum2 = FreeFile
    Open "F:\SPGNEWS\TEST\" + oname$ For Append As #Filenum2
    Print #Filenum2, "S" + Left$(stat$, 1) + " " + field$ + " @ " + atbbs$ + " < " + MYCALL;
    If Mid$(stat$, 2, 1) = "$" Then
        Print #Filenum2, " $" + bid$;
    End If
    Print #Filenum2, ""
    Print #Filenum2, Title

    Do While Not EOF(FileNum1)
        Line Input #FileNum1, l$
        Print #Filenum2, l$
        Loop
    Print #Filenum2, "/EX"
    Close #FileNum1
    Close #Filenum2
End Sub

Function MakeDir (DestPath$) As Integer
    Dim BackPos, forePos As Integer
    Dim temp$

    Screen.MousePointer = 11

    '---------------------------------------------
    ' Add slash to end of path if not there already
    '---------------------------------------------
    If Right$(DestPath$, 1) <> "\" Then
        DestPath$ = DestPath$ + "\"
    End If
          

    '-----------------------------------
    ' Change to the root dir of the drive
    '-----------------------------------
    On Error Resume Next
    ChDrive DestPath$
    If Err <> 0 Then GoTo errorOut
    ChDir "\"

    '-------------------------------------------------
    ' Attempt to make each directory, then change to it
    '-------------------------------------------------
    BackPos = 3
    forePos = InStr(4, DestPath$, "\")
    Do While forePos <> 0
        temp$ = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)

        Err = 0
        MkDir temp$
        If Err <> 0 And Err <> 75 Then GoTo errorOut

        Err = 0
        ChDir temp$
        If Err <> 0 Then GoTo errorOut

        BackPos = forePos
        forePos = InStr(BackPos + 1, DestPath$, "\")
    Loop
                 
    MakeDir = True
    Screen.MousePointer = 0
    ChDir SYSDIR
    Exit Function
                 
errorOut:
    If L_MODE = 0 Then
        MsgBox "ディレクトリ " + Chr$(34) + DestPath$ + Chr$(34) + " が作成できません", 48
    Else
        MsgBox "Can't make diretroty " + Chr$(34) + DestPath$ + Chr$(34), 48
    End If
    MakeDir = False
    ChDir SYSDIR
    Screen.MousePointer = 0
End Function

Function FileBidScan (s As String) As Integer
    Dim s1 As String
    Dim FileNum As Integer
    Dim FileNum1 As Integer

    Trace "FileBidScan"
    
    If FileExist(USERDIR + "BID.DAT") = False Then
        FileNum1 = FreeFile
        Open USERDIR + "BID.DAT" For Output As FileNum1
        Print #FileNum1, "SPG-NEWS V4.3 BID.DAT"
        Close #FileNum1
    End If

    FileNum = FreeFile
' データベースファイルオープン
    Open USERDIR + "BID.DAT" For Input As FileNum

' 検索
    Do While Not EOF(FileNum)
        Line Input #FileNum, s1
        If s = s1 Then                 ' 重複
            Close #FileNum
            FileBidScan = -1
            Exit Function
        End If
    Loop
    Close #FileNum

    FileBidScan = 0
End Function

'
' 起動時のFILE.IDX読み込み
'
' f:True  オープニング画面を表示する
'   False オープニング画面を表示しない
'
Sub FileCSVLoad (f As Integer)
    Dim FileNum As Integer
    Dim Count As Integer
    Dim i, r As Integer
    Dim l, c As Long
    Dim p!, Percent$
    Dim w%, h%, a$

    ReDim MESSAGE(0) As MSGD2

    On Error GoTo FileCSVLoad_err

    Screen.MousePointer = 11
    l = 0
    c = 0
    w% = frmOpen.picCount.Width
    h% = frmOpen.picCount.Height
    If f = True Then
        frmOpen.picCount.Line (w% * .25, 0)-(w% * .25, h%), RGB(255, 0, 0)
        frmOpen.picCount.Line (w% * .5, 0)-(w% * .5, h%), RGB(255, 0, 0)
        frmOpen.picCount.Line (w% * .75, 0)-(w% * .75, h%), RGB(255, 0, 0)
    End If
    Count = 0
    FileNum = FreeFile
    If FileExist(MSGDIR + "\" + "FILE.IDX") = False Then
200     Open MSGDIR + "\" + "FILE.IDX" For Output As FileNum
        Close FileNum
    End If
    Open MSGDIR + "\" + "FILE.IDX" For Input As FileNum
    l = LOF(FileNum)
    If l > 0 Then
        Do While (Not EOF(FileNum))
            Line Input #FileNum, a$
            If a$ <> "" Then
                ReDim Preserve MESSAGE(Count)
                c = c + Len(a$)
                p! = c / l
                If f = True Then
                    frmOpen.picCount.Line (0, 0)-(frmOpen.picCount.Width * p!, frmOpen.picCount.Height), , BF
                End If
                DoEvents
                MessageCut a$, MESSAGE(Count)
                Count = Count + 1
                If Count = 2000 Then
                    If L_MODE = 0 Then
                        MsgBox "メッセージ数が規定値を超えました" + Chr$(13) + Chr$(10) + "古いメッセージを削除してください", MB_ICONEXCLAMATION Or MB_OK
                    Else
                        MsgBox "The number of messages exceeded a predetermined value." + Chr$(13) + Chr$(10) + "Please delete old massege.", MB_ICONEXCLAMATION Or MB_OK
                    End If
                End If
                If Count = 2520 Then
                    If L_MODE = 0 Then
                        MsgBox "これ以上メッセージを記録することができません" + Chr$(13) + Chr$(10) + "古いメッセージを削除して再起動してください", MB_ICONEXCLAMATION Or MB_OK
                    Else
                        MsgBox "これ以上メッセージを記録することができません" + Chr$(13) + Chr$(10) + "The message can't be recorded any further.Please reactive after deleting an old message.", MB_ICONEXCLAMATION Or MB_OK
                    End If
                    Exit Do
                End If
            End If
        Loop
    End If
    Close FileNum
    If f = True Then
        frmOpen.picCount.Line (0, 0)-(w%, 615), , BF
    End If
    Screen.MousePointer = 0
    FILELOADFLAG = True
    frmOpen.Hide
    Exit Sub

FileCSVLoad_err:
    Select Case Erl
    Case 100
        Beep
        If L_MODE = 0 Then
            r = MsgBox("初期設定を終了します。SPG-NEWSを再起動してください。", MB_ICONEXCLAMATION)
        Else
            r = MsgBox("Initialize done.Please restart SPG-NEWS", MB_ICONEXCLAMATION)
        End If
        End
    Case 200
        Close #FileNum
        If L_MODE = 0 Then
            r = MsgBox("メッセージを格納するディレクトリ " + Chr$(34) + MSGDIR + Chr$(34) + " がありません" + Chr$(13) + Chr$(10) + "ディレクトリを作成しますか？", MB_ICONEXCLAMATION Or MB_YESNO)
        Else
            r = MsgBox("There are no Directory of store message." + Chr$(34) + MSGDIR + Chr$(34) + Chr$(13) + Chr$(10) + "Are you make directory?", MB_ICONEXCLAMATION Or MB_YESNO)
        End If
        If r = IDYES Then
            r = MakeDir(MSGDIR)
            If r = False Then
                End
            End If
        Else
            End
        End If
    Case Else
        Beep
        ErrorProc "FileCSVLoad", Err, Erl, Error
        End
    End Select
    Resume
End Sub

' テキストビュアーの内容をファイルにセーブする
'
Sub FileTextSave (Filename)
    Dim FileNum As Integer
    Dim Contents As String
    
    FileNum = FreeFile

    ' ファイルを開きます。
    Open Filename For Output As FileNum
    ' メモ帳の内容を変数に設定します。
    Contents = frmEdit!txtEdit.Text
    ' 砂時計を表示します。
    Screen.MousePointer = 11
    ' 内容を保存するファイルに書きます。
    Print #FileNum, Contents;
    Close #FileNum
    ' マウス ポインタをリセットします。
    Screen.MousePointer = 0
End Sub

Function HTMLCut (s$) As String
    Dim s1$, s2$, ss$, c$
    Dim i, st, en, flag, flaghead As Integer
    
    st = InStr(UCase$(s$), "<HTML>")
    en = InStr(UCase$(s$), "</HTML>")

    If st = 0 Or en = 0 Then
        HTMLFLAG = False
        HTMLCut = s$
        Exit Function
    End If
    
    HTMLFLAG = True
    s1$ = Left$(s$, st - 1)
    s2$ = Mid$(s$, en + 7)

    flag = False
    For i = st + 6 To en
        c$ = Mid$(s$, i, 1)
        Select Case c$
        Case "<"
            flag = True
            If UCase$(Mid$(s$, i, 3)) = "<P>" Then
                ss$ = ss$ + Chr$(13) + Chr$(10)
                i = i + 2
                flag = False
            End If
            If UCase$(Mid$(s$, i, 4)) = "<BR>" Then
                ss$ = ss$ + Chr$(13) + Chr$(10)
                i = i + 3
                flag = False
            End If
            If UCase$(Mid$(s$, i, 6)) = "<HEAD>" Then
                i = i + 5
                flaghead = True
            End If
            If UCase$(Mid$(s$, i, 7)) = "</HEAD>" Then
                i = i + 6
                flaghead = False
            End If
        Case ">"
            flag = False
        Case "&"
            If Mid$(s$, i, 4) = "&gt;" Then
                i = i + 3
                ss$ = ss$ + ">"
            End If
        Case "q"
            If Mid$(s$, i, 5) = "quot;" Then
                i = i + 4
                ss$ = ss$ + Chr$(34)
            End If
        Case Chr$(13)
            If Mid$(s$, i, 2) = Chr$(13) + Chr$(10) Then
                i = i + 1
            End If
        Case Else
            If flag = False And flaghead = False Then
                ss$ = ss$ + c$
            End If
        End Select
    Next i

    HTMLCut = s1$ + ss$ + s2$
End Function

Function StatSelect (sc$, a$) As Integer
    If Left$(a$, 1) = "!" Then
        If Mid$(sc$, 3, 1) <> Mid$(a$, 2, 1) Then
            StatSelect = True
            Exit Function
        End If
    Else
        If Mid$(sc$, 3, 1) = Left$(a$, 1) Then
            StatSelect = True
            Exit Function
        End If
    End If
    StatSelect = False
End Function

Sub ConfigRead ()
    Dim a As Integer
    Dim mycall1 As String * 10
    
    a = GetPrivateProfileString("Config", "MyCall", "NOCALL", mycall1, 7, INIFILE)
    If Left$(mycall1, a) = "NOCALL" Then
        FIRST = True
        MsgBox "ＳＰＧ－ＮＥＷＳをご利用いただきありがとうございます。" + Chr$(13) + Chr$(10) + "環境設定メニューで設定をして下さい。", MB_ICONINFORMATION
        SetUpConfig True
    Else
        FIRST = False
        SetUpConfig False
    End If
End Sub

Sub FileBidAdd (s As String)
    Dim FileNum As Integer
    
    Trace "FileBidAdd"
    
    FileNum = FreeFile

' 追加
    Open USERDIR + "BID.DAT" For Append As FileNum
        Print #FileNum, s
    Close FileNum
End Sub

