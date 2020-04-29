' オートパイロットシステム通信モジュール
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA
Option Explicit

'
' オートパイロットメインルーチン
'
Sub Auto ()
    Dim a  As Integer
    Dim msg As MSGD2
    Dim p$, r$
    Dim i%, j%, st%, stat$
    Dim sm As String * 10
    Dim smf$

    Dim ln As String * 20
    Dim lc As String * 20
    Dim listcommand As Integer
    Dim ln1$
    Dim lc1$
    Dim FileNum As Integer

    RCVCOUNT = 0

    Trace "Auto"
    
    On Error GoTo Auto_Error

    ' ターミナルソフト使用中はオートパロットを実行しない
    If TERMFLAG = True Then
        Exit Sub
    End If
    
    ' 文字列検索位置設定
10    SetSerchPos

    ' コネクト
    frmPanel!pnlInfo.Caption = "オートパイロット コネクト中"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>コネクト開始"
    Idol_off
20  If Connect() <> 0 Then
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>コネクト失敗"
        Disconnect
        Idol_on
        Exit Sub
    End If
    ' ＳＩＤ受信待ち
    frmPanel!pnlInfo.Caption = "オートパイロット ＳＩＤ受信待ち"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>ＳＩＤ受信待ち"

30  If WaitComm("[", RCVWAITTIME) = DISC$ Then
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>SID受信失敗"
        Disconnect
        Idol_on
        Exit Sub
    End If

    ' プロンプト取得
    frmPanel!pnlInfo.Caption = "オートパイロット プロンプト待ち"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>プロンプト待ち"

    ' プロンプトのチェックと生成
    p$ = PromptMake()
    If p$ = DISC$ Then
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>プロンプト取得失敗"
        Disconnect
        Idol_on
        Exit Sub
    End If
    PROMPT1 = p$

    ' タイムアウト設定
    frmTimer!tmrAuto.Interval = 60000   ' 一分間
    frmTimer!tmrAuto.Enabled = True     'タイマースタート
    
    ' ＳＹＳＯＰコマンド発行
    If AutoSendSysopCmd() = False Then
        Exit Sub
    End If

    ' 送信予定メッセージの送信
    If AutoSendMessage() = False Then
        Exit Sub
    End If

    ' 残存リストがあれば取得する
    If FileExist(USERDIR + "LISTNEXT.TXT") = True Then
        listcommand = GetPrivateProfileInt("AutoPilot", "ListCommand", 0, INIFILE)
        FileNum = FreeFile
        Open USERDIR + "LISTNEXT.TXT" For Input As #FileNum
            Input #FileNum, ln1$, lc1$
        Close #FileNum
        Kill USERDIR + "LISTNEXT.TXT"
        If AutoSendList(ln1$, lc1$, listcommand) = False Then
            Exit Sub
        End If
        ' リストを取得してダウンロードするメッセージをREADNEXT.TXTに記録する
        If AutoGetList(ln1$, False) = False Then
            Exit Sub
        End If
    End If
    
    ' 予約済メッセージ読み取り
    frmPanel!pnlInfo.Caption = "オートパイロット 予約済メッセージ読み取り中"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>予約済メッセージ読み取り"
50  If AutoReadNextMail() = False Then
        Disconnect
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>予約済メッセージ読み取り失敗"
        Idol_on
        Exit Sub
    End If
    
' -----------------------------------------------------------------------------
    
    ' リストコマンド送信
    a% = GetPrivateProfileString("AutoPilot", "ListNum", "0", ln, 10, INIFILE)
    a% = GetPrivateProfileString("AutoPilot", "ListCount", "20", lc, 10, INIFILE)
    listcommand = GetPrivateProfileInt("AutoPilot", "ListCommand", 0, INIFILE)
    
    ln1$ = Format$(Val(RNullCut(ln)) + 1)
    If ln1$ = "2" Then
        ln1$ = "1"
    End If
    lc1$ = Format$(Val(ln1$) + Val(RNullCut(lc)))
    
    If AutoSendList(ln1$, lc1$, listcommand) = False Then
        Exit Sub
    End If

    ' リストを取得してダウンロードするメッセージをREADNEXT.TXTに記録する
    If AutoGetList(ln1$, True) = False Then
        Exit Sub
    End If

    ' メッセージ読み込み
    ' READNEXT.TXTに従いメッセージをダウンロード
    frmPanel!pnlInfo.Caption = "オートパイロット メッセージ受信中"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>メッセージ受信"
    If AutoReadNextMail() = False Then
        Disconnect
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>メッセージ読み取り失敗"
        Idol_on
        Exit Sub
    End If
    
    ' オートパイロット正常終了
    frmPanel!pnlInfo.Caption = "オートパイロット 回線断"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>オートパイロット正常終了"
    r$ = SendStr("B", 0, True)
    Disconnect
    Idol_on
    Exit Sub

Auto_Error:
    ErrorProc "AutoPilot", Err, Erl, Error
    PrtInfo "不当なリストを取得しました。" + Chr$(13) + Chr$(10) + "リスト文字化けが原因と推測されます。" + Chr$(13) + Chr$(10) + "通信を強制終了します"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>Error " + Format$(Err) + ":" + Format$(Erl)
    Disconnect
    Idol_on
    Exit Sub
End Sub

'
' リストを取得してダウンロードするメッセージをREADNEXT.TXTに記録する
' リスト取得に失敗した場合リストをとっていないリスト番号をLISTNEXT.TXTに記録する
' ln1$ : 取り始めるリストの最小値
' flag : True=正常終了後最新リスト番号を記録する
'　　　　False=記録しない
'
Function AutoGetList (lnst$, flag As Integer)
    Dim a%
    Dim ln As String * 20
    Dim r$, rb$, rn$, s$
    Dim sel%

    ' 前回コネクト時の最新リスト番号を取得
'    a% = GetPrivateProfileString("AutoPilot", "ListNum", "1", ln, 10, INIFILE)
'    ln1$ = RNullCut(ln)
    ' リストをファイルに記録する
    frmPanel!pnlInfo.Caption = "オートパイロット リスト受信待ち"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>リスト受信待ち"
    r$ = RcvStrLine(RCVWAITTIME, True)
    If r$ = DISC$ Then
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>リスト受信待機中に回線断1" + Chr$(34) + r$ + Chr$(34)
        Disconnect
        Idol_on
        AutoGetList = False
        Exit Function
    End If
    ' リスト本体を受信するまで待つ
    Do While (Val(r$) = 0)
        r$ = RcvStrLine(RCVWAITTIME, True)
        If r$ Like PROMPT1 = True Then
            ' リストなし何もダウンロードしない
            frmPanel!pnlRcv.Caption = ""
            AutoGetList = True
            Exit Function
        End If
        If r$ = DISC$ Then
            FileAdd USERDIR + "LOG.TXT", "AutoPilot>リスト受信待機中に回線断2" + Chr$(34) + r$ + Chr$(34)
            Disconnect
            Idol_on
            AutoGetList = False
            Exit Function
        End If
    Loop

    frmPanel!pnlInfo.Caption = "オートパイロット リスト受信中"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>リスト受信中"
    ' リストの最新番号を記録
    rn$ = r$
    rb$ = r$
    ' リスト取得
    ' プロンプトが来るまでリストを取得する
    Do While (r$ Like PROMPT1 = False)
        If InStr(r$, DISC$) <> 0 Then
            ' リスト受信待機中に回線断
            FileAdd USERDIR + "LOG.TXT", "AutoPilot>リスト受信待機中に回線断3" + Chr$(34) + r$ + Chr$(34)
            Disconnect
            Idol_on
            AutoGetList = False
            FileAdd USERDIR + "LISTNEXT.TXT", lnst$ + "," + Format$(Val(rb$) - 1)
            s$ = Format$(Val(rn$))
           a% = WritePrivateProfileString("AutoPilot", "ListNum", UCase$(s$), INIFILE)
            Exit Function
        End If
        rb$ = r$
        ' ダウンロード検索
60      sel% = MessageSelect(r$)
        ' 検索対象だったら読む
        If sel% = True Then
            frmPanel!pnlRcv.Caption = "する"
            FileAdd USERDIR + "READNEXT.TXT", r$
        Else
            frmPanel!pnlRcv.Caption = "しない"
        End If
        ' 記録
        If Val(r$) <> 0 Then
            r$ = Left$(r$, InStr(r$, Chr$(13)) - 1)
            FileAdd USERDIR + "LIST.TXT", r$
        End If
        ' 一行受信するまで待つ
        r$ = ""
        r$ = RcvStrLine(RCVWAITTIME, True)
        DoEvents
    Loop
    frmPanel!pnlRcv.Caption = ""
    AutoGetList = True
    ' リスト正常取得終了最終リスト番号記録
    If flag = True Then
        s$ = Format$(Val(rn$))
        a% = WritePrivateProfileString("AutoPilot", "ListNum", UCase$(s$), INIFILE)
    End If
End Function

'
' オートパイロット用アイテム抽出
'
Sub AutoListCut (r$, msg As MSGD2)
    Dim i As Integer
    Dim at$, f$, t$, tt$, r1$
    
    Trace "AutoListCut"
    
    r1$ = r$
    If Val(r1$) < 0 Then
        r1$ = Mid$(r1$, 2)
    End If
    ' リストより文字列抽出
    't$ = mid$(r$, AUTOPOSTOFIELD, instr(AUTOPOSTOFIELD, r$, " ") - AUTOPOSTOFIELD)
    ' To Fieldの区切り文字はスペースと@の場合があるので両方区切り文字として検索する
    tt$ = ""
    For i = AUTOPOSTOFIELD To Len(r1$)
        tt$ = Mid$(r1$, i, 1)
        If tt$ = " " Or tt$ = "@" Then
            Exit For
        End If
        t$ = t$ + tt$
    Next i
    at$ = Mid$(r1$, AUTOPOSATBBS, InStr(AUTOPOSATBBS, r1$, " ") - AUTOPOSATBBS)
    f$ = Mid$(r1$, AUTOPOSCALL, InStr(AUTOPOSCALL, r1$, " ") - AUTOPOSCALL)
    tt$ = Mid$(r1$, AUTOPOSTITLE, InStr(AUTOPOSTITLE, r1$, Chr$(13)) - AUTOPOSTITLE)
    ' MSGD2設定
    If t$ = MYCALL Then
        msg.stat = "PF"
    Else
        msg.stat = "AF"
    End If
    msg.field = t$
    msg.atbbs = at$
    msg.Call = f$
    msg.Title = tt$
    msg.bid = ""
End Sub

'
' メール読み出し
'
' n$ : <num>
'      文字列先頭に - があると削除する
'
Function AutoReadMail (n$) As Integer
    Dim FileNum As Integer
    Dim mk As String * 10
    Dim a As Integer
    Dim cmd$, p$, r$, c%, mkf$
    Dim num As Long                     ' B124

    Trace "AutoReadMail"

    On Error GoTo AutoReadMail_Err
    
6001    num = Val(n$)
    ' プロンプト取得
    ' Readコマンド生成
6002    cmd$ = READCMD + " " + Format$(Abs(num))
    ' Readコマンド送信
6003    r$ = SendStr(cmd$, 0, True)
    ' ファイルオープン
6004    FileNum = FreeFile
6005    Open USERDIR + "READ.TXT" For Output As FileNum
    ' メール読み出し
6006    r$ = RcvStrLine(RCVWAITTIME, True)
    ' プロンプトを検出するまで読み出す
6007    While (r$ Like PROMPT1 = False)
6008        If InStr(r$, DISC$) <> 0 Then
6009            Close FileNum
6010            Disconnect
6011            AutoReadMail = False
6012            Exit Function
6013        End If
6014        Print #FileNum, r$;
6015        r$ = RcvStrLine(RCVWAITTIME, True) ' B170
6016        DoEvents
6017    Wend
6018    Close FileNum
    ' 自分宛のメールだったら削除する
6019    a = GetPrivateProfileString("AutoPilot", "MailKill", "Y", mk, 5, INIFILE)
6020    mkf$ = Left$(mk, 1)
6021    If Val(n$) < 0 And mkf$ = "Y" Then     ' B185
        ' Killコマンド生成
6022        cmd$ = "K " + Format$(Abs(num))
        ' Readコマンド送信
6023        r$ = SendStr(cmd$, 0, True)
        ' プロンプト受信待ち
6024        c% = 0
6025        r$ = RcvStrLine(RCVWAITTIME, True)
6026        Do While (r$ Like PROMPT1 = False)
6027            c% = c% + 1
6028            If c% = 10 Then
6029                FileAdd USERDIR + "LOG.TXT", "AutoPilot>自分宛メール受信中回線断"
6030                Disconnect
6031                Idol_on
6032                AutoReadMail = False
6033                Exit Function
6034            End If
6035            r$ = RcvStrLine(RCVWAITTIME, True)
6036        Loop
6037    End If
6038    AutoReadMail = True
6039    Exit Function

AutoReadMail_Err:
    ErrorProc "AutoReadMail", Err, Erl, Error
    FileAdd USERDIR + "LOG.TXT", "AutoReadMail>Error " + Format$(Err) + ":" + Format$(Erl)
    Disconnect
    Idol_on
    Exit Function
End Function

'
' READNEXT.TXTにもとづいてメッセージを読み出す
' メッセージ読み取り中エラーが発生したらREADNEXT.TXTを再生する
'
Function AutoReadNextMail ()
    ReDim ReadNum(0) As String
    Dim msg As MSGD2
    Dim FileNum As Integer
    Dim i%, j%, rc%, st%, r$

    On Error GoTo AutoReadNextMail_Err
    
    ' 読み込み予定のメッセージを読み込む
5000    rc% = 0
5001    FileNum = FreeFile
5002    st% = True
    ' READNEXT.TXTがあれば読み込み予定メッセージを読み込む
5003    If (FileExist(USERDIR + "READNEXT.TXT") = True) Then
5004        Open USERDIR + "READNEXT.TXT" For Input As FileNum
5005        ' リスト読み込み
5006        Do While (EOF(FileNum) = False)
5007            ReDim Preserve ReadNum(rc%) As String
5008            Line Input #FileNum, r$
5009            ReadNum(rc%) = r$
5010            rc% = rc% + 1
5011        Loop
5012        Close FileNum
        ' メッセージ読み込み
5013        For i% = rc% - 1 To 0 Step -1
5014            If Val(ReadNum(i%)) <> 0 Then  ' R 0 が送信されるバグの暫定対策(後で本対策する)
5015                AutoListCut ReadNum(i%) + Chr$(13), msg
5016                frmPanel!pnlToField.Caption = msg.field
5017                frmPanel!pnlAtbbs.Caption = msg.atbbs
5018                frmPanel!pnlCall.Caption = msg.Call
5019                frmPanel!pnlTitle.Caption = msg.Title
                    ' 自分宛のメールだったら行頭に"-"を追加する 'B186
                    If msg.field = MYCALL Then
                        ReadNum(i%) = "-" + ReadNum(i%)
                    End If
5020                st% = AutoReadMail(ReadNum(i%))
                ' 読み込み失敗の場合
5021                If st% = False Then
5022                    Exit For
5023                End If
                ' ファイル書き込み
                ' @BBSが指定されていなかったらHomeRBBSを指定する
5024                If msg.atbbs = "" Then
5025                    msg.atbbs = HOMERBBS
5026                End If
5027                If (FileWrite2(USERDIR + "READ.TXT", msg.stat, msg.field, msg.atbbs, msg.Call, msg.Title, msg.bid, "***", False) <> 0) Then
5028                    AutoReadNextMail = False
5029                    Exit Function
5030                End If
5031            End If
5032        Next i%
5033        Kill USERDIR + "READNEXT.TXT"
        ' メッセージを正常にダウンロードできなかった場合
        ' 次回ダウンロード予定ファイルを作成する
5034        If st% = False Then
'5035            For j% = i% To 0 Step -1
5035            For j% = 0 To i%
5036                FileAdd USERDIR + "READNEXT.TXT", ReadNum(j%)
5037            Next j%
5038            AutoReadNextMail = False
5039            Exit Function
5040        End If
5041    End If
5042    AutoReadNextMail = True
5043    Exit Function

AutoReadNextMail_Err:
    ErrorProc "AutoReadNextMail", Err, Erl, Error
    FileAdd USERDIR + "LOG.TXT", "AutoReadNextMail>Error " + Format$(Err) + ":" + Format$(Erl)
    Disconnect
    Idol_on
    Exit Function
End Function

'
' メッセージの送信
' 戻り値: -1 DISCONNECT
'         -2 Sコマンド入力拒否
'
Function AutoSendFile (i As Integer) As Integer
    Dim FileNum As Integer
    Dim Filename As String
    Dim a$, r$, send$

    Trace "AutoSendFile"
    
    r$ = RcvStr(RCVWAITTIME, True)
    ' TITLEINPUTが設定されていない場合はプロンプトのチェックをしない
    If TITLEINPUT <> "" Then
        If r$ <> TITLEINPUT + Chr$(13) + Chr$(10) Then
            ' タイトル入力プロンプトでない
            AutoSendFile = -2
            Exit Function
        End If
    End If
    If InStr(r$, DISC$) <> 0 Then
        AutoSendFile = -1
        Exit Function
    End If

    ' タイトル送信
    a$ = SendStr(StrCut(MESSAGE(i).Title), 0, True)
    If InStr(a$, DISC$) <> 0 Then
        AutoSendFile = -1
        Exit Function
    End If

    r$ = RcvStrLine(RCVWAITTIME, True)
    If InStr(r$, DISC$) <> 0 Then
        AutoSendFile = -1
        Exit Function
    End If
    Filename = SpaceCut(MSGDIR) + "\" + StrCut(MESSAGE(i).fname)
    FileNum = FreeFile
    Open Filename For Input As FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, send$
        a$ = SendStr(send$, 0, True) 'B170
        If InStr(a$, DISC$) <> 0 Then
            AutoSendFile = -1
            Close #FileNum
            Exit Function
        End If
    Wend
'    send$ = Chr$(&H1A)
    send$ = "/EX"
    a$ = SendStr(send$, 0, True) 'B174
    If InStr(a$, DISC$) <> 0 Then  ' 95/01/22 V3.32.1
        AutoSendFile = -1
        Close #FileNum
        Exit Function
    End If
    AutoSendFile = 0
    Close #FileNum
End Function

'
' リストコマンド送信
'
Function AutoSendList (ln1$, lc1$, listcommand As Integer)
    Dim a%
'    Dim ln As String * 20
'    Dim lc As String * 20
'    Dim listcommand As Integer
    Dim sp$, s$, r$

    ' リストコマンド送信
    frmPanel!pnlInfo.Caption = "オートパイロット リストコマンド送信"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>リストコマンド送信"
'    a% = GetPrivateProfileString("AutoPilot", "ListNum", "1", ln, 10, INIFILE)
'    a% = GetPrivateProfileString("AutoPilot", "ListCount", "20", lc, 10, INIFILE)
'    listcommand = GetPrivateProfileInt("AutoPilot", "ListCommand", 0, INIFILE)
    
'    ln1$ = RNullCut(ln)
'    lc1$ = Format$(Val(ln1$) + Val(RNullCut(lc)))
    If listcommand = 0 Then
        sp$ = " "
    Else
        sp$ = "-"
    End If

    s$ = "L " + ln1$ + sp$ + lc1$
    r$ = SendStr(s$, 0, True)
    If InStr(r$, DISC$) <> 0 Then
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>リストコマンド送信失敗"
        Disconnect
        Idol_on
        AutoSendList = False
        Exit Function
    End If
    AutoSendList = True
End Function

'
' 送信予定のメッセージを送信する
'
Function AutoSendMessage ()
    Dim c%, i%, MaxNum%, st%
    Dim r$, stat$

    ' メッセージ書き込み
    frmPanel!pnlInfo.Caption = "オートパイロット メッセージ送信中"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>メッセージ送信"
    MaxNum% = UBound(MESSAGE)

    For i% = 0 To MaxNum%
        stat$ = Mid$(MESSAGE(i%).stat, 2, 1)
        If (stat$ <> "") And (stat$ <> "K") And (stat$ <> "F") And Mid$(MESSAGE(i%).stat, 3, 1) <> "K" And Mid$(MESSAGE(i%).stat, 3, 1) <> "H" Then
            frmPanel!pnlToField.Caption = MESSAGE(i%).field
            frmPanel!pnlAtbbs.Caption = MESSAGE(i%).atbbs
            frmPanel!pnlCall.Caption = MESSAGE(i%).Call
            frmPanel!pnlTitle.Caption = MESSAGE(i%).Title
            ' 送信コマンド送信
100         If (FWDSendCmd(i%, False) <> 0) Then
                FileAdd USERDIR + "LOG.TXT", "AutoPilot>送信コマンド送信中に回線断" + Chr$(34) + r$ + Chr$(34)
                Disconnect
                Idol_on
                AutoSendMessage = False
                Exit Function
            End If
            ' ファイル送信
            st% = AutoSendFile(i%)
110         If st% = -1 Then
                ' 回線断
                FileAdd USERDIR + "LOG.TXT", "AutoPilot>ファイル送信中に回線断" + Chr$(34) + r$ + Chr$(34)
                Disconnect
                Idol_on
                AutoSendMessage = False
                Exit Function
            End If
            If st% = -2 Then
                ' Ｓコマンドが拒否された Fフラグを立てて送信しないようにする
                Mid$(MESSAGE(i%).stat, 2) = "F"
                FILECHANGEFLAG = True
            End If
            ' プロンプト受信待ち
            c% = 0
            r$ = RcvStrLine(RCVWAITTIME, True)
            Do While (r$ Like PROMPT1 = False)
                c% = c% + 1
                If c% = 10 Then
                    FileAdd USERDIR + "LOG.TXT", "AutoPilot>ファイル送信終了後プロンプト待ちで回線断"
                    Disconnect
                    Idol_on
                    AutoSendMessage = False
                    Exit Function
                End If
                r$ = RcvStrLine(RCVWAITTIME, True)
            Loop
            ' ステータス更新
            Mid$(MESSAGE(i%).stat, 2) = "F"
            FILECHANGEFLAG = True
        End If
    Next i%
    AutoSendMessage = True
End Function

'
' ＳＹＳＯＰコマンド発行
'
Function AutoSendSysopCmd ()
    Dim a%
    Dim sm As String * 5
    Dim smf$, r$, p$

    a% = GetPrivateProfileString("AutoPilot", "SysopMode", "N", sm, 3, INIFILE)
    smf$ = Left$(sm, 1)
    If smf$ = "Y" Then
        r$ = SendStr("@", 0, True)
        ' プロンプトのチェックと再生成
        p$ = PromptMake()
        If p$ = DISC$ Then
            FileAdd USERDIR + "LOG.TXT", "AutoPilot>ＳＹＳＯＰモード移行失敗"
            Disconnect
            Idol_on
            AutoSendSysopCmd = False
            Exit Function
        End If
        PROMPT1 = p$
    End If
    AutoSendSysopCmd = True
End Function

Sub AutoTmp ()
'    ' メッセージ読み込み
'    frmPanel!pnlInfo.Caption = "オートパイロット メッセージ受信中"
'    FileAdd SYSDIR + "LOG.TXT", "AutoPilot>メッセージ受信"
'    st% = True
'    For i% = rc% - 1 To 0 Step -1
'70      AutoListCut ReadNum(i%), msg
'        frmPanel!pnlToField.Caption = msg.field
'        frmPanel!pnlAtbbs.Caption = msg.atbbs
'        frmPanel!pnlCall.Caption = msg.Call
'        frmPanel!pnlTitle.Caption = msg.Title
'        ' 自分宛のメールだったら行頭に"-"を追加する
'        If msg.field = MYCALL Then
'            ReadNum(i%) = "-" + ReadNum(i%)
'        End If
'80      st% = AutoReadMail(ReadNum(i%))
'        If st% = False Then
'            Exit For
'        End If
'        ' ファイル書き込み
'        ' @BBSが指定されていなかったらHomeRBBSを指定する
'        If msg.atbbs = "" Then
'            msg.atbbs = HOMERBBS
'        End If
'        If (FileWrite2("READ.TXT", msg.stat, msg.field, msg.atbbs, msg.Call, msg.Title, msg.bid, "***") <> 0) Then
'            FileAdd SYSDIR + "LOG.TXT", "AutoPilot>ファイル書き込み失敗" + Chr$(34) + r$ + Chr$(34)
'            Disconnect
'            Idol_on
'            st% = False
'            Exit For
'        End If
'    Next i%
'    ' メッセージを正常にダウンロードできなかった場合
'    ' 次回ダウンロード予定ファイルを作成する
'    If st% = False Then
'        For j% = i% To 0 Step -1
'90          AutoListCut ReadNum(j%), msg
'            If msg.field = MYCALL And Val(ReadNum(j%)) >= 0 Then
'                ReadNum(j%) = "-" + ReadNum(j%)
'            End If
'            FileAdd SYSDIR + "READNEXT.TXT", ReadNum(j%)
'        Next j%
'        FileAdd SYSDIR + "LOG.TXT", "AutoPilot>メッセージダウンロード中に回線断"
'        Disconnect
'        Idol_on
'        Exit Sub
'    End If
    
'    Do While (EOF(FileNum) = False)
'        Input #FileNum, yn$(imax%), tfl$(imax%), atbbs$(imax%), f1$(imax%), Title$(imax%)
'        imax% = imax% + 1
'    Loop
'    Trace "Select2"
'    For i% = 0 To imax%
'        Trace "Select3"
'        If (t$ Like tfl$(i%) = True) And (at$ Like atbbs$(i%) = True) And (f$ Like f1$(i%) = True) And (titl$ Like Title$(i%) = True) Then
'            Trace "Select4"
'            If yn$(i%) = "@" Then
'                MessageSelect = True
'            Else
'                MessageSelect = False
'                Exit Do
'            End If
'        End If
'    Next i%
End Sub

'
' オートパイロット用検索データ作成
'
Sub MakeSerchData ()
    Dim a, i, j, l As Integer
    ReDim s1(20) As String
    Dim at$, c$, t$

    Dim tfl As String * 600
    Dim atbbs As String * 600
    Dim callsign As String * 600
    
    Trace "MakeSerchData"
    
    a = GetPrivateProfileString("AutoPilot", "ToField", "", tfl, 200, INIFILE)
    a = GetPrivateProfileString("AutoPilot", "Atbbs", "", atbbs, 200, INIFILE)
    a = GetPrivateProfileString("AutoPilot", "CallSign", "", callsign, 200, INIFILE)

    t$ = RNullCut(tfl)
    at$ = RNullCut(atbbs)
        s1(j) = ""
    c$ = RNullCut(callsign)

'---------------------------------------------------------------
    i = 1
    j = 0
    l = Len(t$)
    Do
        Do While (Mid$(t$, i, 1) <> ",")
            If i > l Then
                GoTo loop_out
            End If
            s1(j) = s1(j) + Mid$(t$, i, 1)
            i = i + 1
        Loop
        j = j + 1
        i = i + 1
    Loop
            
loop_out:
    ReDim Preserve AUTOTOFIELD(j + 1)
    For i = 0 To j
        AUTOTOFIELD(i) = s1(i)
        s1(i) = ""
    Next i
    AUTOTOFIELD(j + 1) = "-"
'---------------------------------------------------------------
    i = 1
    j = 0
    l = Len(at$)
    Do
        Do While (Mid$(at$, i, 1) <> ",")
            If i > l Then
                GoTo loop_out1
            End If
            s1(j) = s1(j) + Mid$(at$, i, 1)
            i = i + 1
        Loop
        j = j + 1
        i = i + 1
    Loop
            
loop_out1:
    ReDim Preserve AUTOATBBS(j + 1)
    For i = 0 To j
        AUTOATBBS(i) = s1(i)
        s1(i) = ""
    Next i
    AUTOATBBS(j + 1) = "-"
'---------------------------------------------------------------
    i = 1
    j = 0
    l = Len(c$)
    Do
        Do While (Mid$(c$, i, 1) <> ",")
            If i > l Then
                GoTo loop_out2
            End If
            s1(j) = s1(j) + Mid$(c$, i, 1)
            i = i + 1
        Loop
        j = j + 1
        i = i + 1
    Loop
            
loop_out2:
    ReDim Preserve AUTOCALL(j + 1)
    For i = 0 To j
        AUTOCALL(i) = s1(i)
        s1(i) = ""
    Next i
    AUTOCALL(j + 1) = "-"
End Sub

'
'  リストのダウンロード選択
'
Function MessageSelect (s$) As Integer
    Dim t$, tt$, titl$, at$, f$
    Dim i%, imax%
    Dim FileNum As Integer
'    Dim yn$, tfl$, atbbs$, f1$, Title$
    ReDim yn$(100), tfl$(100), atbbs$(100), f1$(100), Title$(100)
    
    MessageSelect = False

    ' リスト以外の文字列か??
    If Val(s$) = 0 Then
        MessageSelect = False
        Exit Function
    End If

    ' リストより文字列抽出
    ' To Fieldの区切り文字はスペースと@の場合があるので両方区切り文字として検索する
    tt$ = ""
    For i% = AUTOPOSTOFIELD To Len(s$)
        tt$ = Mid$(s$, i%, 1)
        If tt$ = " " Or tt$ = "@" Then
            Exit For
        End If
        t$ = t$ + tt$
    Next i%
    at$ = Mid$(s$, AUTOPOSATBBS, InStr(AUTOPOSATBBS, s$, " ") - AUTOPOSATBBS)
    f$ = Mid$(s$, AUTOPOSCALL, InStr(AUTOPOSCALL, s$, " ") - AUTOPOSCALL)
    titl$ = Mid$(s$, AUTOPOSTITLE)

    ' 自分宛のメールは読み込む
    If t$ = MYCALL Then
        MessageSelect = True
        Exit Function
    End If

    If FileExist(USERDIR + "DOWNLOAD.SPG") = False Then
        MessageSelect = True
        Exit Function
    End If
        
    Trace "Select1"

    FileNum = FreeFile
    imax% = 0
    Open USERDIR + "DOWNLOAD.SPG" For Input As #FileNum
    Do While (EOF(FileNum) = False)
        Input #FileNum, yn$(imax%), tfl$(imax%), atbbs$(imax%), f1$(imax%), Title$(imax%)
        imax% = imax% + 1
        If imax% = 100 Then
            Exit Do
        End If
    Loop
    Close #FileNum

    For i% = 0 To imax%
        If (t$ Like tfl$(i%) = True) And (at$ Like atbbs$(i%) = True) And (f$ Like f1$(i%) = True) And (titl$ Like Title$(i%) = True) Then
            If yn$(i%) = "@" Then
                MessageSelect = True
            Else
                MessageSelect = False
                Exit For
            End If
        End If
    Next i%
    
    frmPanel!pnlToField.Caption = t$
    frmPanel!pnlAtbbs.Caption = at$
    frmPanel!pnlCall.Caption = f$
    frmPanel!pnlTitle.Caption = titl$
End Function

'
' 検出するプロンプトの生成
' プロンプトのチェックと生成を行う
'
Function PromptMake () As String
    Dim r$, p$, p1$, i%

    Do
        r$ = RcvStrLine(RCVWAITTIME, True)
        If r$ = DISC$ Then
            PromptMake = DISC$
            Exit Function
        End If
        If p$ = ">" + Chr$(13) + Chr$(10) Then
            MsgBox "ＲＢＢＳのプロンプトが > のみです。メッセージ中の引用符などをプロンプトとして誤認識する可能性があります。オートパイロット停止後手動でＲＢＢＳに接続しプロンプトを変更することをおすすめします。"
        End If
        If InStr(r$, ">" + Chr$(13) + Chr$(10)) <> 0 Then
        ' プロンプト生成
            For i% = 1 To Len(r$)
                p1$ = Mid$(r$, i%, 1)
                If p1$ = " " Or (Asc(p1$) >= &H30 And Asc(p1$) <= &H39) Then
                    p$ = p$ + "*"
                ElseIf p1$ = "[" Then
                    p$ = p$ + "[[]"
                Else
                    p$ = p$ + p1$
                End If
            Next i%
            Exit Do
        End If
    Loop
    PromptMake = p$
End Function

' 一行受信
' 一行受信するまで待機する
'
' t:タイムアウト時間
' dsp_flag:モニタ表示フラグ(True=表示する)
'
Function RcvStrLine (t As Integer, dsp_flag As Integer) As String
    Dim s$, r$

'    Trace "RcvStrLine"
    
    ' 一行受信するまで待つ
    Do While (InStr(r$, Chr$(13)) = 0)
        r$ = RcvStr(t, dsp_flag)
        If r$ = DISC$ Then
            RcvStrLine = DISC$
            Exit Function
        End If
        s$ = s$ + r$
    Loop
    RcvStrLine = s$
End Function

'
' ダウンロードカラム設定
'
Sub SetSerchPos ()
    Trace "SetSerchPos"
    AUTOPOSTOFIELD = GetPrivateProfileInt("AutoPilot", "ToFieldColum", 17, INIFILE)
    AUTOPOSCALL = GetPrivateProfileInt("AutoPilot", "CallColum", 31, INIFILE)
    AUTOPOSTITLE = GetPrivateProfileInt("AutoPilot", "TitleColum", 48, INIFILE)
    AUTOPOSATBBS = GetPrivateProfileInt("AutoPilot", "AtbbsColum", 24, INIFILE)
End Sub

