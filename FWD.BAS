' ＦＷＤ転送システム
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA
Option Explicit

'
' Ｓコマンド解析
'
Sub CmdAna (s As String, msg As MSGD2)
    Dim i As Integer
    Dim j As Integer
    ReDim s1(20) As String
    Dim l As Integer

    Trace "CmdAna"

    ' コマンドライン分割
    i = 1
    j = 0
    l = Len(s)
    Do
        Do While (Mid$(s, i, 1) = " ")
            i = i + 1
        Loop
        Do While (Mid$(s, i, 1) <> " ")
            If i = l Then
                s1(j) = s1(j) + Mid$(s, i, 1)
                GoTo loop_out
            End If
            s1(j) = s1(j) + Mid$(s, i, 1)
            i = i + 1
        Loop
        j = j + 1
    Loop
loop_out:
    ' ステータス（Ｂ／Ｐ）
    ' To Field設定
    msg.bid = ""
    msg.stat = UCase$(Mid$(s1(0), 2, 1) + "F")
    msg.field = UCase$(s1(1))

    ' @BBS設定
    For i = 2 To j
        If s1(i) = "@" Then
            msg.atbbs = UCase$(s1(i + 1))
            Exit For
        End If
    Next i

    ' From設定
    For i = 2 To j
        If s1(i) = "<" Then
            msg.Call = UCase$(s1(i + 1))
            Exit For
        End If
    Next i

    ' BID設定
    For i = 2 To j
        If Left$(s1(i), 1) = "$" Then
            msg.bid = Mid$(s1(i), 2)
            Exit For
        End If
    Next i
End Sub

Sub Disconnect ()
    Dim i As Integer
    Dim a%, b$

    Trace "Disconnect"
    If LINK = "TNC" Then
        b$ = SendStr("D", 1, True)
    Else
        b$ = SendStr(DISC$, 0, True)
    End If
    i = Dtr_off()
    a% = SerialClose()
End Sub

Sub FWD ()
    Dim a2%, i%, MaxNum%
    Dim a$, stat$, r$

    On Error GoTo FWD_Error
    
    ' ターミナルソフト使用中は転送しない
    If TERMFLAG = True Then
        Exit Sub
    End If

    ' コネクト
    frmPanel!pnlInfo.Caption = "ＦＷＤ転送 コネクト中"
    Idol_off
10  If Connect() <> 0 Then
        Exit Sub
    End If
    ' ＳＩＤ受信待ち
    frmPanel!pnlInfo.Caption = "ＦＷＤ転送 ＳＩＤ受信待ち"
20  If (WaitComm("[", RCVWAITTIME) = DISC$) Then
        Disconnect
        Idol_on
        Exit Sub
    End If
    ' プロンプト受信待ち
    frmPanel!pnlInfo.Caption = "ＦＷＤ転送 プロンプト待ち"
30  If (WaitComm(">", RCVWAITTIME) = DISC$) Then
        Disconnect
        Idol_on
        Exit Sub
    End If
    
    ' ＳＩＤ送信
    frmPanel!pnlInfo.Caption = "ＦＷＤ転送 ＳＩＤ送信"
40  a$ = SendStr("[SPG-4.3-$]", 0, True)
    If InStr(a$, DISC$) <> 0 Then
        Disconnect
        Idol_on
        Exit Sub
    End If

    ' プロンプト受信待ち
    frmPanel!pnlInfo.Caption = "ＦＷＤ転送 プロンプト待ち"
50  If (WaitComm(">", RCVWAITTIME) = DISC$) Then
        Disconnect
        Idol_on
        Exit Sub
    End If
    
    ' 送信
    MaxNum% = UBound(Message)

    For i% = 0 To MaxNum%
        stat$ = Mid$(Message(i%).stat, 2, 1)
        If (stat$ <> "") And (stat$ <> "K") And (stat$ <> "F") And Mid$(Message(i%).stat, 3, 1) <> "K" And Mid$(Message(i%).stat, 3, 1) <> "H" Then
            ' 送信コマンド送信
            frmPanel!pnlInfo.Caption = "ＦＷＤ転送 送信コマンド送信"
            frmPanel!pnlToField.Caption = Message(i%).field
            frmPanel!pnlAtbbs.Caption = Message(i%).atbbs
            frmPanel!pnlCall.Caption = Message(i%).Call
            frmPanel!pnlTitle.Caption = Message(i%).Title
60          If (FWDSendCmd(i%, True) <> 0) Then
                Disconnect
                Idol_on
                Exit Sub
            End If
            ' OK/NO待ち
            frmPanel!pnlInfo.Caption = "ＦＷＤ転送 OK/NO待ち"
            r$ = RcvStr(RCVWAITTIME, True)
            If Left$(r$, 1) <> "N" Then
                ' ファイル送信
                frmPanel!pnlInfo.Caption = "ＦＷＤ転送 メッセージ送信"
70              If (FWDSendFile(i%) <> 0) Then
                    Disconnect
                    Idol_on
                    Exit Sub
                End If
            End If
            ' プロンプト受信待ち
            frmPanel!pnlInfo.Caption = "ＦＷＤ転送 プロンプト待ち"
80          If (WaitComm(">", RCVWAITTIME) = DISC$) Then
                Disconnect
                Idol_on
                Exit Sub
            End If
            ' ステータス更新
            Mid$(Message(i%).stat, 2) = "F"
        End If
        ClearPanel
    Next i%

    ' 受信
    frmPanel!pnlInfo.Caption = "ＦＷＤ転送 受信中"
    Do
        r$ = SendStr("F>", 0, True)
        If InStr(r$, DISC$) <> 0 Then
            Exit Do
        End If
        r$ = RcvStr(RCVWAITTIME, True)
        If InStr(r$, DISC$) <> 0 Then
            Exit Do
        End If
        If Left$(r$, 1) <> "S" Then
            Exit Do
        End If
        r$ = Left$(r$, InStr(r$, Chr$(13)) - 1)
90      a2% = FWDRcv(r$)
        If a2% = -1 Then
            Exit Do
        End If
    Loop
    Disconnect
    Idol_on
    Exit Sub

FWD_Error:
    ErrorProc "FWD", Err, Erl, Error
    If L_MODE = 0 Then
        MsgBox "FWD転送中にエラーが発生しました。" + Chr$(13) + Chr$(10) + "通信を強制終了します", MB_ICONEXCLAMATION, "SPG-NEWS"
    Else
        MsgBox "FWD Error." + Chr$(13) + Chr$(10) + "FWD Abort.", MB_ICONEXCLAMATION, "SPG-NEWS"
    End If
    FileAdd USERDIR + "LOG.TXT", "FWD>Error " + Format$(Err) + ":" + Format$(Erl)
    Disconnect
    Idol_on
    Exit Sub
End Sub

'
' ＦＷＤ受信
'
Function FWDRcv (s As String) As Integer
    Dim msg As MSGD2
    Dim FileNum As Integer
    Dim a$, b$, r$

    Trace "FWDRcv"
    
    CmdAna s, msg

    If msg.bid <> "" Then
        b$ = SpaceCut(msg.bid)
        If FileBidScan(b$) <> 0 Then
            a$ = SendStr("NO", 0, True)
            If InStr(a$, DISC$) <> 0 Then
                FWDRcv = -1
                Exit Function
            End If
            FWDRcv = -2
            Exit Function
        End If
    End If
    a$ = SendStr("OK", 0, True)
    If InStr(a$, DISC$) <> 0 Then
        FWDRcv = -1
        Exit Function
    End If
    'タイトル受信
    r$ = RcvStr(RCVWAITTIME, True)
    If InStr(r$, DISC$) <> 0 Then
        FWDRcv = -1
        Exit Function
    End If
    msg.Title = Mid$(r$, 1, Len(r$) - 2)

    frmPanel!pnlToField.Caption = msg.field
    frmPanel!pnlAtbbs.Caption = msg.atbbs
    frmPanel!pnlCall.Caption = msg.Call
    frmPanel!pnlTitle.Caption = msg.Title

    ' 受信ファイルオープン
    FileNum = FreeFile
    Open "RCV.TXT" For Output As FileNum
    ' メール本体受信
    Do
        r$ = RcvStr(RCVWAITTIME, False)
        If InStr(r$, DISC$) <> 0 Then
            Close #FileNum
            FWDRcv = -1
            Exit Function
        End If
        If InStr(r$, "/EX") = 1 Or InStr(r$, Chr$(&H1A)) Then
            Close #FileNum
            Exit Do
        End If
        Print #FileNum, r$;
    Loop

    ' ファイル書き込み
    If (FileWrite2("RCV.TXT", msg.stat, msg.field, msg.atbbs, msg.Call, msg.Title, msg.bid, "***", False) <> 0) Then
        FWDRcv = -1
        Exit Function
    End If
    FWDRcv = 0
End Function

'
' ＦＷＤ送信コマンド送信
'
Function FWDSendCmd (i As Integer, flag As Integer) As Integer
    Dim a$, s$

    Trace "FWDSendCmd"

    s$ = "S" + Left$(Message(i).stat, 1) + " "
    s$ = s$ + Message(i).field
    If Message(i).atbbs <> "" Then
        s$ = s$ + " @ " + Message(i).atbbs
    End If
    If flag = True Then
        s$ = s$ + " < " + Message(i).Call
    End If
    If Mid$(Message(i).stat, 2, 1) = "$" Then
        s$ = s$ + " $" + Message(i).bid
    End If
    a$ = SendStr(s$, 0, True)
    If InStr(a$, DISC$) <> 0 Then
        FWDSendCmd = -1
        Exit Function
    End If
    FWDSendCmd = 0
End Function

'
' ＦＷＤファイル送信
'
Function FWDSendFile (i As Integer) As Integer
    Dim FileNum As Integer
    Dim Filename As String
    Dim a$, send$

    Trace "FWDSendFile"
    
    ' タイトル送信
    a$ = SendStr(StrCut(Message(i).Title), 0, True)
    If InStr(a$, DISC$) <> 0 Then
        FWDSendFile = -1
        Exit Function
    End If

    Filename = SpaceCut(MSGDIR) + "\" + StrCut(Message(i).fname)
    FileNum = FreeFile
    Open Filename For Input As FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, send$
        a$ = SendStr(send$, 0, False)
        If InStr(a$, DISC$) <> 0 Then
            FWDSendFile = -1
            Close #FileNum
            Exit Function
        End If
    Wend
    send$ = Chr$(&H1A)
'    send$ = "/EX"
    a$ = SendStr(send$, 0, False)
    If InStr(a$, DISC$) <> 0 Then
        FWDSendFile = -1
        Close #FileNum
        Exit Function
    End If
    FWDSendFile = 0
    Close #FileNum
End Function

'
' コネクト
' 接続からプロンプトが返ってくるまで
'
Function Connect () As Integer
    Dim ss As String * 100
    Dim sr As String * 100
    Dim i, FileNum As Integer
    Dim a$, b$, rr$, s1$, s2$, a2$
    Dim a1%

    frmPanel!pnlInfo.Caption = "通信回線オープン"
    s1$ = "[送信中] "
    s2$ = "[受信待ち] "
    If (OpenPort() <> 0) Then
        frmPanel!pnlInfo.Caption = "通信回線オープン失敗"
        Connect = -1
        Exit Function
    End If
    
    i = FlushComm(nCid, 1)
    i = Dtr_on()

    ' コネクト処理
    If LINK = "TNC" Then
        frmPanel!pnlInfo.Caption = "ＴＮＣ設定"
        a2$ = SendStr("ECHO OFF", 1, True)
        a2$ = SendStr("AUTOLF OFF", 1, True)
        a2$ = SendStr("c", 1, True)
        If (WaitComm("Link", RCVWAITTIME) = DISC$) Then
            a1% = SerialClose()
            Connect = -1
            Exit Function
        End If
    Else
        a2$ = SendStr("", 0, True)
        a2$ = SendStr("", 0, True)
        a2$ = SendStr("", 0, True)
    End If
    
    If FILEEXist(USERDIR + "CONNECT.SPG") = True Then
        FileNum = FreeFile
        Open USERDIR + "CONNECT.SPG" For Input As #FileNum
        i = 0
        Do While (EOF(FileNum) = False)
            Line Input #FileNum, a$
            Line Input #FileNum, b$
            If b$ = "" Then Exit Do
            If a$ = "R" Then
                frmPanel!pnlInfo.Caption = s2$ + b$
                If (WaitComm(b$, RCVWAITTIME) = DISC$) Then
                    Disconnect
                    Connect = -1
                    Close #FileNum
                    Exit Function
                End If
            Else
                frmPanel!pnlInfo.Caption = s1$ + b$
                rr$ = SendStr(b$, 0, True)
                If InStr(rr$, DISC$) <> 0 Then
                    Disconnect
                    Connect = -1
                    Close #FileNum
                    Exit Function
                End If
            End If
        Loop
        Close #FileNum
    Else
        Disconnect
        Connect = -1
        Exit Function
    End If

    frmPanel!pnlInfo.Caption = "コネクト成功"
    Connect = 0
End Function

