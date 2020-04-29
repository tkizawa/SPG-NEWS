'共通呼出プロシージャ
' Copyright (c) 1993-1997 by JO1SPG/T.KIZAWA
Option Explicit

Sub AttrChange ()
    Dim ListNum&
    Dim i As Integer

    ListNum = Val(Left$(frmMain!lstMenu.Text, 5))
    i = DataGetNum(ListNum&)
End Sub

Sub AutoConfig ()
    frmAutoConfig.Show 1
End Sub

'
' オートパイロット起動
'
Sub AutoPilotStart ()
    Dim a As Integer
    
    If FILEEXist(USERDIR + "CONNECT.SPG") = False Then
        MsgBox "ＲＢＢＳへのコネクト方法が指定されていません。通信設定で指定してください。", 64, "SPG-NEWS"
        Exit Sub
    End If
    a = GetPrivateProfileInt("AutoPilot", "ToFieldColum", -1, INIFILE)
    If a = -1 Then
        MsgBox "オートパイロットの設定がありません。オートパイロット設定で設定してください。", 64, "SPG-NEWS"
        frmAutoConfig.Show 1
        Exit Sub
    End If
    Auto    ' A_PILOT.BAS
    frmTimer!tmrAuto.Enabled = False     'タイマー停止
End Sub

' 通信システム起動（送信・受信）
'
Sub Forward ()
    If FILEEXist(USERDIR + "CONNECT.SPG") = False Then
        If L_MODE = 0 Then
            MsgBox "ＲＢＢＳへのコネクト方法が指定されていません。接続設定で指定してください。", 64, "SPG-NEWS"
        Else
            MsgBox "The connection method to RBBS is not specified.", 64, "SPG-NEWS"
        End If
        frmFWDConfig.Show 1
        Exit Sub
    End If
    ' FWD転送開始
    FWD
End Sub

Sub MenuFileKill ()
    Dim ListNum&
    Dim idx As Integer
    Dim i, j As Integer
    Dim lst$

    screen.MousePointer = 11
    For i = 0 To frmMain!lstMenu.ListCount - 1
        If i >= frmMain!lstMenu.ListCount Then
            Exit For
        End If
        If frmMain!lstMenu.Selected(i) = True Then
            ListNum& = Val(Left$(frmMain!lstMenu.List(i), 5))
            j = DataGetNum(ListNum&)
            If Len(MESSAGE(j).stat) < 3 Then
                MESSAGE(j).stat = MESSAGE(j).stat + "K"
            Else
                Mid$(MESSAGE(j).stat, 3, 1) = "K"
            End If
            lst$ = FileListLineMake(j)
            frmMain!lstMenu.List(i) = lst$
        End If
    Next i
    FILECHANGEFLAG = True
    screen.MousePointer = 0
End Sub

Sub MenuFileRevival ()
    Dim ListNum&
    Dim idx As Integer
    Dim i, j As Integer
    Dim lst$

    screen.MousePointer = 11
    For i = 0 To frmMain!lstMenu.ListCount - 1
        If i >= frmMain!lstMenu.ListCount Then
            Exit For
        End If
        If frmMain!lstMenu.Selected(i) = True Then
            ListNum = Val(Left$(frmMain!lstMenu.List(i), 5))
            j = DataGetNum(ListNum&)
            MESSAGE(j).stat = Left$(MESSAGE(j).stat, 2)
            lst$ = FileListLineMake(j)
            frmMain!lstMenu.List(i) = lst$
        End If
    Next i
    FILECHANGEFLAG = True
    screen.MousePointer = 0
End Sub

Sub OpenAbout ()
    frmAbout.Hide
    frmAbout.Left = (screen.Width - frmAbout.Width) / 2
    frmAbout.Top = (screen.Height - frmAbout.Height) / 2
    frmAbout.Show 1
End Sub

Sub OpenHelp ()
    Dim a As Integer
    
    If L_MODE = 0 Then
        a = Shell("WINHELP SPGNEWS.HLP", 1)
    Else
        a = Shell("WINHELP SPG_ENG.HLP", 1)
    End If
End Sub

Sub ReadMail ()
    Dim ListNum&
    Dim i As Integer
    Dim lst$

    ListNum& = Val(Left$(frmMain!lstMenu.Text, 5))
    ' メッセージ通し番号設定
    i = DataGetNum(ListNum&)
    SetMessageNum i
'    If Len(MESSAGE(i).stat) < 3 Then
'        MESSAGE(i).stat = MESSAGE(i).stat + "*"
'    Else
'        Mid$(MESSAGE(i).stat, 3, 1) = "*"
'    End If
    FileTextMake ListNum&
End Sub

Sub ReplyBulletin ()
    Dim FileNum As Integer
    Dim FileNum1 As Integer
    Dim ListNum&
    Dim num As Integer
    Dim Filename As String * 80
    Dim Title As String
    Dim s1 As String
    Dim s2 As String
    Dim root As String
    Dim i%
    Dim t$, strline$, fr$, tt$, dd$, ss$, s$, msgid$
    Dim rhf, ssf, hf As Integer

    rhf = False
    hf = False
    RE_MARK = ">"
    ToolBarDisable
    If frmMain!lstMenu.Text = "" Then
        Exit Sub
    End If
    ListNum = Val(Left$(frmMain!lstMenu.Text, 5))
    Filename = FileNameMake(Trim$(Str$(ListNum)))

    num = DataGetNum(ListNum&)
    
    On Error Resume Next
    frmEdit!cboToField.Clear
    frmEdit!cboAtbbs.Clear

'ライトモードＢ
    WRITEMODE = "B"
    R_MODE = False

    FileNum = FreeFile
    Open USERDIR + "SEND.TXT" For Output As FileNum
    Close #FileNum

    ' 宛て先指定コンボボックス設定
    FileNum = FreeFile
    Open USERDIR + "TOFIELD.DAT" For Input As #FileNum
    FileNum1 = FreeFile
    Open USERDIR + "ATBBS.DAT" For Input As #FileNum1

    Do While Not EOF(FileNum)
        Line Input #FileNum, s1
        frmEdit!cboToField.AddItem s1
    Loop
   
    Do While Not EOF(FileNum1)
        Line Input #FileNum1, s2
        frmEdit!cboAtbbs.AddItem s2
    Loop
    
    i% = InStr(MESSAGE(num).Title, Chr$(13))
    t$ = MESSAGE(num).Title
    If i% <> 0 Then
        t$ = Left$(MESSAGE(num).Title, i% - 1)
    End If

    Close #FileNum
    Close #FileNum1

    SetReMark MESSAGE(num).Call

    FileNum = FreeFile
    Open Filename For Input As #FileNum
    Do While (EOF(FileNum) <> True)
        Line Input #FileNum, strline$
        If Left$(strline$, 2) = "R:" Then
            root = strline$
        End If
        If strline$ = "" Then
            hf = True
            Exit Do
        End If
        If InStr(1, strline$, "From: ") = 1 Then
            fr$ = Mid$(strline$, 7)
        End If
        If InStr(1, strline$, "Subject: ") = 1 Then
            tt$ = Mid$(strline$, 10)
        End If
        If InStr(1, strline$, "Date: ") = 1 Then
            dd$ = Mid$(strline$, 7)
        End If
        If InStr(1, strline$, "Message-ID: ") = 1 Then
            msgid$ = Mid$(strline$, 13)
        End If
        ss$ = ss$ + RE_MARK + strline$ + Chr$(13) + Chr$(10)
    Loop

    If hf = True Then
        ss$ = ""
    End If

    '最初の一行のヘッダーのチェックをする
'    If InStr(1, strline$, "From: ") = 1 Then
'        fr$ = Mid$(strline$, 7)
'    End If
'    If InStr(1, strline$, "Subject: ") = 1 Then
'        tt$ = Mid$(strline$, 10)
'    End If
'    If InStr(1, strline$, "Date: ") = 1 Then
'        dd$ = Mid$(strline$, 7)
'    End If
'    If InStr(1, strline$, "Message-ID: ") = 1 Then
'        msgid$ = Mid$(strline$, 13)
'    End If

'    While Not EOF(FileNum)
    Do
        Line Input #FileNum, strline$
        If InStr(1, strline$, "From: ") = 1 Then
            fr$ = Mid$(strline$, 7)
            ssf = True
        End If
        If InStr(1, strline$, "Subject: ") = 1 Then
            tt$ = Mid$(strline$, 10)
            ssf = True
        End If
        If InStr(1, strline$, "Date: ") = 1 Then
            dd$ = Mid$(strline$, 7)
            ssf = True
        End If
        If InStr(1, strline$, "Message-ID: ") = 1 Then
            msgid$ = Mid$(strline$, 13)
            ssf = True
        End If
        If hf = True Then
            ss$ = ss$ + RE_MARK + strline$ + Chr$(13) + Chr$(10)
        End If
    Loop While Not EOF(FileNum)
'    Wend
    Close #FileNum
    
    ' Subjectからタイトルが抽出できなかった場合はファイル管理情報にあるタイトルを使う
    If tt$ = "" Then
        tt$ = t$
    End If

'    If root <> "" Then
'        s$ = "In-reply-to: " + MESSAGE(num).Call + "@" + HeadCut(root) + " 's message" + Chr$(13) + Chr$(10)  ' of " + Getrfctime(HeadDate(root), HeadTime(root)) + Chr$(13) + Chr$(10)
'        s$ = s$ + "References: <" + HeadMsgNum(root) + "@" + HeadCut(root) + "> by " + MESSAGE(num).Call + "@" + HeadCut(root) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
'    Else
        If InStr(fr$, ".") <> 0 Then
            fr$ = Left$(fr$, InStr(fr$, ".") - 1)
        End If
        If fr$ <> "" And dd$ <> "" And msgid$ <> "" Then
            s$ = "In-reply-to: " + fr$ + " 's message" + " of " + dd$ + Chr$(13) + Chr$(10)
            s$ = s$ + "References: " + msgid$ + " " + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
        End If
'    End If
    ss$ = s$ + ss$

    frmEdit!cboToField.Text = MESSAGE(num).field
    frmEdit!cboAtbbs.Text = MESSAGE(num).atbbs
    frmEdit!txtTitle = ReplyTitleMake(tt$)
    frmEdit!txtEdit = ss$
    frmEdit.Caption = "掲示板返信書き込み"
    frmEdit.Show 0
End Sub

Sub ReplyMail ()
    Dim FileNum As Integer
    Dim ListNum&
    Dim num As Integer
    Dim Filename As String * 80
    Dim Title As String
    Dim s1 As String
    Dim s2 As String
    Dim root As String
    Dim strline$, fr$, t$, dd$, tt$, ss$, s$, a$, msgid$
    Dim i%, frl%
    Dim hf As Integer
    
    RE_MARK = ">"
    ToolBarDisable
    If frmMain!lstMenu.Text = "" Then
        Exit Sub
    End If
    ListNum = Val(Left$(frmMain!lstMenu.Text, 5))
    Filename = FileNameMake(Trim$(Str$(ListNum)))

    num = DataGetNum(ListNum&)

'    On Error Resume Next
    frmEdit!cboToField.Clear
    frmEdit!cboAtbbs.Clear
    frmEdit!cboAtbbs.Visible = False
    frmEdit!lblAtBBS.Visible = False

' ライトモードＰ
    WRITEMODE = "P"
    R_MODE = False

    FileNum = FreeFile
    Open USERDIR + "SEND.TXT" For Output As #FileNum
    Close #FileNum

    ' 宛て先指定コンボボックス設定
    If FILEEXist(USERDIR + "ADDR.DAT") = True Then
        FileNum = FreeFile
        Open USERDIR + "ADDR.DAT" For Input As #FileNum

        Do While Not EOF(FileNum)
            Line Input #FileNum, s1
            frmEdit!cboToField.AddItem s1
        Loop
        Close #FileNum
    End If

    SetReMark MESSAGE(num).Call
    FileNum = FreeFile
    Open Filename For Input As #FileNum
    Do While (EOF(FileNum) <> True)
        Line Input #FileNum, strline$
        If Left$(strline$, 2) = "R:" Then
            root = strline$
        End If
        If strline$ = "" Then
            ' 空行があればフラグON
            hf = True
            Exit Do
        End If
        If InStr(strline$, "From: ") = 1 Then
            fr$ = strline$
        End If
        If InStr(strline$, "Subject: ") = 1 Then
            t$ = Mid$(strline$, 10)
        End If
        If InStr(strline$, "Date: ") = 1 Then
            dd$ = Mid$(strline$, 7)
        End If
        If InStr(strline$, "Message-ID: ") = 1 Then
            msgid$ = Mid$(strline$, 13)
        End If
        ss$ = ss$ + RE_MARK + strline$ + Chr$(13) + Chr$(10)
    Loop
    
    ' 空行があった場合ヘッダー部は引用しない
    If hf = True Then
        ss$ = ""
    End If

    '最初の一行のヘッダーのチェックをする
'    If InStr(strline$, "From: ") = 1 Then
'        fr$ = Mid$(strline$, 7)
'    End If
'    If InStr(strline$, "Subject: ") = 1 Then
'        t$ = Mid$(strline$, 10)
'    End If
'    If InStr(strline$, "Date: ") = 1 Then
'        dd$ = Mid$(strline$, 7)
'    End If

'    While Not EOF(FileNum)
    Do While (EOF(FileNum) <> True)
        Line Input #FileNum, strline$
        If InStr(1, strline$, "From: ") = 1 Then
            fr$ = strline$
        End If
        If InStr(1, strline$, "Subject: ") = 1 Then
            tt$ = Mid$(strline$, 10)
        End If
        If InStr(strline$, "Date: ") = 1 Then
            dd$ = Mid$(strline$, 7)
        End If
        If InStr(strline$, "Message-ID: ") = 1 Then
            msgid$ = Mid$(strline$, 13)
        End If
        ' 空行があった場合引用には追加しない
        If hf = True Then
            ss$ = ss$ + RE_MARK + strline$ + Chr$(13) + Chr$(10)
        End If
    Loop ' While Not EOF(FileNum)
'    Wend
    Close #FileNum
    
    i% = InStr(MESSAGE(num).Title, Chr$(13))
    t$ = MESSAGE(num).Title
    If i% <> 0 Then
        t$ = Left$(MESSAGE(num).Title, i% - 1)
    End If

    ' Subjectからタイトルが抽出できなかった場合はファイル管理情報にあるタイトルを使う
    If tt$ = "" Then
        tt$ = t$
    End If

    ' ルートヘッダーがある場合
'    If root <> "" Then
'        frmEdit!cboToField.Text = MESSAGE(num).Call + "@" + HeadCut2(root)
'        s$ = "In-reply-to: " + MESSAGE(num).Call + "@" + HeadCut(root) + " 's message" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)' of " + Getrfctime(HeadDate(root), HeadTime(root)) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
    ' ルートヘッダーが無い場合
'    Else
        frl% = InStr(7, fr$, " ")
        If frl% <> 0 Then
            fr$ = Mid$(fr$, 7, frl% - 7)
        Else
            fr$ = Mid$(fr$, 7)
        End If
        ' 原文からFrom:抽出できなかった場合はファイル管理情報にあるタイトルを使う
        If fr$ = "" Then
            fr$ = MESSAGE(num).Call
        End If
        frmEdit!cboToField.Text = fr$
        If InStr(fr$, ".") <> 0 Then
            fr$ = Left$(fr$, InStr(fr$, ".") - 1)
        End If
        If dd$ <> "" Then
            s$ = "In-reply-to: " + fr$ + " 's message" + " of " + dd$ + Chr$(13) + Chr$(10)
            If msgid$ <> "" Then
                s$ = s$ + "References: " + msgid$ + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
            End If
        Else
            s$ = "In-reply-to: " + fr$ + " 's message" + Chr$(13) + Chr$(10)
            If msgid$ <> "" Then
                s$ = s$ + "References: " + msgid$ + " " + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
            End If
        End If
'   End If

    ss$ = s$ + ss$
'    frmEdit!txtTitle = "RE:" + tt$
    frmEdit!txtTitle = ReplyTitleMake(tt$)
    frmEdit!txtEdit = ss$
    frmEdit.Caption = "電子メール返信書き込み"
    frmEdit.Show 0
End Sub

Sub ReWrite ()
    Dim FileNum As Integer
    Dim FileNum1 As Integer
    Dim ListNum&
    Dim num As Integer
    Dim Filename As String * 80
    Dim Title As String
    Dim s1 As String
    Dim s2 As String
    Dim i%, f%, frl%
    Dim a$, t$, fr$, tt$, dd$, ss$, s$
    
    ToolBarDisable
    If frmMain!lstMenu.Text = "" Then
        Exit Sub
    End If
    ListNum = Val(Left$(frmMain!lstMenu.Text, 5))
'    Filename = MSGDIR + "\" + Trim$(Str$(ListNum))
    Filename = FileNameMake(Trim$(Str$(ListNum)))

    num = DataGetNum(ListNum&)
'    If MESSAGE(num).Call <> MYCALL And (mid$(MESSAGE(num).stat, 2, 1) = "$" Or mid$(MESSAGE(num).stat, 2, 1) = " ") Then
'        Exit Sub
'    End If

'    On Error Resume Next
    frmEdit!cboToField.Clear
    frmEdit!cboAtbbs.Clear

'ライトモード設定
    WRITEMODE = Left$(MESSAGE(num).stat, 1)

    FileNum = FreeFile
    Open USERDIR + "SEND.TXT" For Output As FileNum
    Close #FileNum

    If WRITEMODE = "B" Then
        ' 宛て先指定コンボボックス設定
        FileNum = FreeFile
        Open USERDIR + "TOFIELD.DAT" For Input As #FileNum
        FileNum1 = FreeFile
        Open USERDIR + "ATBBS.DAT" For Input As #FileNum1
    
        Do While Not EOF(FileNum)
            Line Input #FileNum, s1
            frmEdit!cboToField.AddItem s1
        Loop
       
        Do While Not EOF(FileNum1)
            Line Input #FileNum1, s2
            frmEdit!cboAtbbs.AddItem s2
        Loop
        Close #FileNum
        Close #FileNum1
    Else
        frmEdit!cboAtbbs.Visible = False
        frmEdit!lblAtBBS.Visible = False
        ' 宛て先指定コンボボックス設定
        If FILEEXist(USERDIR + "ADDR.DAT") = True Then
            FileNum = FreeFile
            Open USERDIR + "ADDR.DAT" For Input As #FileNum
            Do While Not EOF(FileNum)
                Line Input #FileNum, s1
                frmEdit!cboToField.AddItem s1
            Loop
            Close #FileNum
        End If
    End If

    i% = InStr(MESSAGE(num).Title, Chr$(13))
    t$ = MESSAGE(num).Title
    If i% <> 0 Then
        t$ = Left$(MESSAGE(num).Title, i% - 1)
    End If

    s$ = ""
    ss$ = ""
    FileNum = FreeFile
    Open Filename For Input As #FileNum
    f% = False
    While InStr(1, s$, "Subject: ") <> 1
        Line Input #FileNum, s$
        If InStr(1, s$, "To: ") = 1 Then
            fr$ = s$
        End If
    Wend
    ' Subject:の次の行が空行だったら取り込まない
    Line Input #FileNum, s$
    If s$ <> "" Then
        ss$ = s$ + Chr$(13) + Chr$(10)
    Else
        ss$ = ""
    End If
    While Not EOF(FileNum)
        Line Input #FileNum, s$
        ss$ = ss$ + s$ + Chr$(13) + Chr$(10)
    Wend
    Close #FileNum
    frl% = InStr(5, fr$, " ")
    If frl% <> 0 Then
        a$ = Mid$(fr$, 5, frl% - 5)
    Else
        a$ = Mid$(fr$, 5)
    End If
    If WRITEMODE = "P" Then
        frmEdit!cboToField.Text = a$
    Else
        frmEdit!cboToField.Text = MESSAGE(num).field
    End If
    frmEdit!cboAtbbs.Text = MESSAGE(num).atbbs
    frmEdit!txtTitle = t$
    frmEdit!txtEdit = ss$
    frmEdit.Caption = "書き込み再編集"
    R_MODE = True
    frmEdit.Show 0
End Sub

Sub SetReMark (s$)
    Dim i%, s1$

    For i% = Len(s$) To 1 Step -1
        If Asc(Mid$(s$, i%, 1)) <= &H39 Then
            s1$ = Mid$(s$, i% + 1)
            Exit For
        End If
    Next i%
    
    frmReMark!optReMark(0).Caption = ">"
    frmReMark!optReMark(1).Caption = s$ + ">"
    frmReMark!optReMark(2).Caption = s1$ + ">"
    frmReMark!optReMark(3).Caption = "引用符なし"
    frmReMark.Show 1
End Sub

Sub WriteBulletin ()
    Dim FileNum1, Filenum2 As Integer
    Dim s1 As String
    Dim s2 As String
    Dim dummy As String

    On Error Resume Next
'    frmEdit!cboToField.clear
'    frmEdit!cboAtbbs.clear
'　ライトモードＢ
    WRITEMODE = "B"
    R_MODE = False

    FileNum1 = FreeFile
    Open USERDIR + "SEND.TXT" For Output As #FileNum1
    Close #FileNum1

    FileNum1 = FreeFile
    Open USERDIR + "TOFIELD.DAT" For Input As #FileNum1
    Filenum2 = FreeFile
    Open USERDIR + "ATBBS.DAT" For Input As #Filenum2

    Do While Not EOF(FileNum1)
        Line Input #FileNum1, s1
        frmEdit!cboToField.AddItem s1
    Loop
   
    Do While Not EOF(Filenum2)
        Line Input #Filenum2, s2
        frmEdit!cboAtbbs.AddItem s2
    Loop
    
    Close #FileNum1
    Close #Filenum2

    frmEdit.Caption = "掲示板書き込み"
    frmEdit.Show 0
End Sub

Sub WriteMail ()
    Dim FileNum As Integer
    Dim s1 As String
    Dim s2 As String
    Dim dummy As String

    On Error GoTo ERRORPROC

    frmEdit!cboToField.Clear
    frmEdit!cboAtbbs.Clear
    frmEdit!cboAtbbs.Visible = False
    frmEdit!lblAtBBS.Visible = False
'    frmEdit!fraToField.Caption = "メール宛て先"
'    frmEdit!fraAtbbs.Visible = False
'ライトモードＰ
    WRITEMODE = "P"
    R_MODE = False

    FileNum = FreeFile
    Open USERDIR + "SEND.TXT" For Output As #FileNum
    Close #FileNum

    FileNum = FreeFile
1000 Open USERDIR + "ADDR.DAT" For Input As #FileNum
    
    Do While Not EOF(FileNum)
        Line Input #FileNum, s1
        frmEdit!cboToField.AddItem s1
    Loop
   
    Close #FileNum

1010 frmEdit.Caption = "電子メール書き込み"
    frmEdit.Show 0
    Exit Sub

ERRORPROC:
    If Erl = 1000 Then
        GoTo 1010
    End If
    Error Err
End Sub

