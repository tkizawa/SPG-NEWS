Option Explicit

Sub SetUpConfig (flag As Integer)
    Dim a As Integer
    Dim i%, j%, k%, c%
    Dim mycall1 As String * 10
    Dim homerbbs1 As String * 40
    Dim link1 As String * 10
    Dim baud1 As String * 10
    Dim msgdir1 As String * 40
    Dim editor1 As String * 100
    Dim mode1 As String * 10
    Dim MailHead As String * 10
    Dim ComMonitor As String * 10
    Dim FontSizeList As String * 10
    Dim FontSizeText As String * 20
    Dim SoundFilename As String * 80
    Dim myname1 As String * 100
    Dim se As String * 10
    Dim te As String * 10
    Dim tz, tzl As Integer
    Dim tzs As String * 10
    Dim djldir As String * 80
    Dim djldirl As Integer
    Dim djlspool As String * 80
    Dim djlspooll As Integer
    Dim spooldir1 As String * 80
    Dim spooldirl As Integer
    Dim titleinput1 As String * 200
    Dim titleinputl As Integer
    Dim monitor1 As String * 80
    Dim monitorl As Integer
    Dim timelimit1 As Integer
    Dim senddelay1 As String * 20
    Dim htmleditorl As Integer
    Dim htmleditor1 As String * 100
    Dim webbrowserl As Integer
    Dim webbrowser1 As String * 100
    Dim userdir1 As String * 128
    Dim userdirl As Integer
    Dim readcmd1 As String * 10
    Dim readcmdl As Integer

    On Error Resume Next

    a = GetPrivateProfileString("Config", "MyCall", "NOCALL", mycall1, 7, INIFILE)
    a = GetPrivateProfileString("Config", "HomeRBBS", "JM1YOV.11.JNET1.JPN.AS", homerbbs1, 40, INIFILE)
    a = GetPrivateProfileString("Config", "MyName", "", myname1, 80, INIFILE)
    a = GetPrivateProfileString("Config", "Link", "RS232C", link1, 10, INIFILE)
    a = GetPrivateProfileString("Config", "Baud", "9600", baud1, 10, INIFILE)
    a = GetPrivateProfileString("Config", "MsgDir", CurDir$, msgdir1, 40, INIFILE)
    userdirl = GetPrivateProfileString("Config", "UserDir", CurDir$, userdir1, 80, INIFILE)
    a = GetPrivateProfileString("Config", "Editor", "NOTEPAD.EXE", editor1, 80, INIFILE)
    a = GetPrivateProfileString("Config", "Mode", "FWD", mode1, 10, INIFILE)
    a = GetPrivateProfileString("Config", "MailHeader", "ON", MailHead, 5, INIFILE)
    a = GetPrivateProfileString("Config", "ComMonitor", "ON", ComMonitor, 5, INIFILE)
    PORT = GetPrivateProfileInt("Config", "Port", 1, INIFILE)
    tzl = GetPrivateProfileString("Config", "TimeZone", "+0900", tzs, 6, INIFILE)
    a = GetPrivateProfileString("Font", "TextViewFontName", "標準明朝", FONTNAMETEXTVIEW, 20, INIFILE)
    a = GetPrivateProfileString("Font", "TextViewFontSize", "12", FontSizeText, 5, INIFILE)
    a = GetPrivateProfileString("Font", "ListMenuFontName", "標準明朝", FONTNAMELISTMENU, 20, INIFILE)
    a = GetPrivateProfileString("Font", "ListMenuFontSize", "12", FontSizeList, 5, INIFILE)
    a = GetPrivateProfileString("Timer", "Enable", "OFF", te, 5, INIFILE)
    a = GetPrivateProfileString("Timer", "StartHour", "00", TIMESTARTHOUR, 5, INIFILE)
    a = GetPrivateProfileString("Timer", "EndHour", "23", TIMEENDHOUR, 5, INIFILE)
    a = GetPrivateProfileString("Timer", "Minutes", "00", TIMEMINUTES, 5, INIFILE)
    TIMERMODE = GetPrivateProfileInt("Timer", "Mode", 0, INIFILE)
    TIMERINTERVAL = GetPrivateProfileInt("Timer", "IntervalMinutes", 60, INIFILE)
    a = GetPrivateProfileString("Server", "Enable", "OFF", se, 5, INIFILE)
    a = GetPrivateProfileString("Server", "Cycle", "00", SERVERCYCLE, 5, INIFILE)
    RCVWAITTIME = GetPrivateProfileInt("Config", "RcvTimeOut", 3, INIFILE)
    SENDWAITTIME = GetPrivateProfileInt("Config", "SendTimeOut", 1, INIFILE)
    a = GetPrivateProfileString("Config", "SendDelay", "0", senddelay1, 10, INIFILE)
    SENDDELAY = GetPrivateProfileInt("Config", "SendDelay", 0, INIFILE)
    L_MODE = GetPrivateProfileInt("Config", "Language", 0, INIFILE)
    djldirl = GetPrivateProfileString("Option", "djldir", "C:\BBS", djldir, 80, INIFILE)
    djlspooll = GetPrivateProfileString("Option", "djlspool", "C:\DJLSPOOL", djlspool, 80, INIFILE)
    spooldirl = GetPrivateProfileString("Option", "spooldir", "C:\SPGSPOOL", spooldir1, 80, INIFILE)
    titleinputl = GetPrivateProfileString("AutoPilot", "TitlePrompt", "", titleinput1, 160, INIFILE)
    monitorl = GetPrivateProfileString("Config", "Monitor", "ON", monitor1, 10, INIFILE)
    TIMELIMIT = GetPrivateProfileInt("AutoPilot", "TimeLimit", -1, INIFILE)
    htmleditorl = GetPrivateProfileString("Option", "HTMLEditor", "C:\MSOFFICE\WINWORD\WINWORD.EXE", htmleditor1, 80, INIFILE)
    webbrowserl = GetPrivateProfileString("Option", "WEBBrowser", "C:\PROGRA~1\MICROS~3\IEXPLORE.EXE", webbrowser1, 80, INIFILE)
    XFLOWFLAG = GetPrivateProfileInt("Config", "XFlow", 0, INIFILE)
    readcmdl = GetPrivateProfileString("AutoPilot", "ReadCommand", "R", readcmd1, 7, INIFILE)
    SPOOLUP = GetPrivateProfileInt("Option", "spoolup", True, INIFILE)

    READCMD = Left$(readcmd1, readcmdl)
    FONTSIZETEXTVIEW = Val(FontSizeText)
    FONTNAMETEXTVIEW = Trim$(FONTNAMETEXTVIEW)
    FONTSIZELISTMENU = Val(FontSizeList)
    FONTNAMELISTMENU = Trim$(FONTNAMELISTMENU)

    MYCALL = UCase$(SpaceCut(mycall1))
    If MYCALL = "JO1SPG" Then
        LAPUTAMODE = True
    Else
        LAPUTAMODE = False
    End If
    HOMERBBS = UCase$(SpaceCut(homerbbs1))
    MYNAME = RNullCut(myname1)
    LINK = UCase$(SpaceCut(link1))
    BAUD = SpaceCut(baud1)
    MSGDIR = UCase$(SpaceCut(msgdir1))
    EDITOR = UCase$(SpaceCut(editor1))
    MODE = UCase$(SpaceCut(mode1))
    MAILHEADER = UCase$(SpaceCut(MailHead))
    COM_MONITOR = UCase$(SpaceCut(ComMonitor))
    TIMERENABLE = UCase$(SpaceCut(te))
    TIMESTARTHOUR = Trim$(TIMESTARTHOUR)
    TIMEENDHOUR = Trim$(TIMEENDHOUR)
    TIMEMINUTES = Trim$(TIMEMINUTES)
    SERVERENABLE = UCase$(SpaceCut(se))
    SERVERCYCLE = Trim$(SERVERCYCLE)
    TIMEZONE = Left$(tzs, tzl)
    DJLMAILDIR = Left$(djldir, djldirl)
    DJLSPOOLDIR = Left$(djlspool, djlspooll)
    SPOOLDIR = Left$(spooldir1, spooldirl)
    USERDIR = Left$(userdir1, userdirl) + "\"
    TITLEINPUT = Left$(titleinput1, titleinputl)
    HTMLEDITOR = Left$(htmleditor1, htmleditorl)
    WEBBROWSER = Left$(webbrowser1, webbrowserl)
    If Left$(monitor1, monitorl) = "ON" Then
        MONITORON = True
    Else
        MONITORON = False
    End If
'    DJLMAILDIR = "E:\BBS" 'TEST
'    MODE = "DJLMB" ' TEST
    If MODE = "DJLMB" Then
        GetDJLMailDir
    End If

    ' CONFIGダイアログ設定
    frmSetup!txtMyCall.Text = MYCALL
    frmSetup!txtFwdCall.Text = HOMERBBS
    frmSetup!txtMyName.Text = MYNAME
    If LINK = "TNC" Then
        frmSetup!optLink(0).Value = True
    Else
        frmSetup!optLink(1).Value = True
    End If
    Select Case BAUD
        Case "38400"
            frmSetup!cboBaud.ListIndex = 0
        Case "19200"
            frmSetup!cboBaud.ListIndex = 1
        Case "9600"
            frmSetup!cboBaud.ListIndex = 2
        Case "4800"
            frmSetup!cboBaud.ListIndex = 3
        Case "2400"
            frmSetup!cboBaud.ListIndex = 4
        Case "1200"
            frmSetup!cboBaud.ListIndex = 5
        Case Else
            frmSetup!cboBaud.ListIndex = 2
    End Select
    tz = Val(Left$(TIMEZONE, 3))
    If tz < 0 Then
        frmSetup!cboTimeZone.ListIndex = -tz
    Else
        frmSetup!cboTimeZone.ListIndex = tz + 12
    End If
    
    frmSetup!txtMsgDir.Text = MSGDIR
    frmSetup!drvMsg.Drive = Left$(MSGDIR, 2)
    frmSetup!dirMsg.Path = MSGDIR
    frmSetup!txtEditor.Text = EDITOR
    frmSetup!optPort(PORT - 1).Value = True
    Select Case MODE
    Case "FWD"
        frmSetup!optMode(0).Value = True
    Case "AUTO"
        frmSetup!optMode(1).Value = True
    Case "DJLMB"
        frmSetup!optMode(2).Value = True
    Case "ATERM"
        frmSetup!optMode(3).Value = True
    End Select

    ' タイマー設定
    ReDim TIMERSET(24)
    i% = Val(TIMESTARTHOUR)
    j% = Val(TIMEENDHOUR)
    k% = i%
    c% = 0
    While (k% <> j%)
        k% = i% Mod 24
        TIMERSET(c%) = Right$("00" + Mid$(Str$(k%), 2), 2) + ":" + Right$("00" + Mid$(Str$(Val(TIMEMINUTES)), 2), 2)
        c% = c% + 1
        i% = i% + 1
    Wend
    TIMERSET(c%) = "-1"
    
    ' タイムアウト時間設定
    frmSetup!cboRcvTimeOut.ListIndex = RCVWAITTIME - 1
    frmSetup!cboSendTimeOut.ListIndex = SENDWAITTIME - 1
    SENDDELAY = Val(senddelay1)
    frmSetup!txtSendDelay.Text = Format$(SENDDELAY)

    ' 読み込み終了
    ' コンフィグレーション
    If flag = True Then
        frmSetup.Show 1
    End If
End Sub

