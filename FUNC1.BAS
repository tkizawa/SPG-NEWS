'汎用関数
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Function CharFilter (s$) As String
    Dim i%, l%, kflag%, k%
    Dim s1$

    If Len(s$) = 3 Then
	k% = Asc(Left$(s$, 1))
	If Iskanji(k%) = True Then
	    CharFilter = Chr$(13) + Chr$(10)
	    Exit Function
	Else
	    CharFilter = s$
	    Exit Function
	End If
    End If
	    
    If Right$(s$, 1) <> Chr$(&HA) Then
	CharFilter = s$
	Exit Function
    End If
    
    kflag% = False
    s1$ = s$
    For i% = 1 To Len(s$)
	k% = Asc(Mid$(s1$, i%, 1))
	If k% = 13 Then
	    Exit For
	End If
	If Iskanji(k%) = True Then
	    ' 既に漢字フラグが立っている場合は漢字の2バイト目と識別する
	    If kflag% = True Then
		kflag% = False
	    Else
		kflag% = True
	    End If
	Else
	    kflag% = False
	End If
    Next i%
    
    If kflag% = True Then
	CharFilter = Left$(s$, Len(s$) - 3) + Right$(s$, 2)
    Else
	CharFilter = s$
    End If
End Function

Sub ClearPanel ()
    frmPanel!pnlToField.Caption = ""
    frmPanel!pnlAtbbs.Caption = ""
    frmPanel!pnlCall.Caption = ""
    frmPanel!pnlTitle.Caption = ""
    frmPanel!pnlRcv.Caption = ""
End Sub

Function DayCount (s As String) As Long
    Dim i, n, isleapyear  As Integer
    Dim days As Long

    Trace "DayCount"

    s = Left$(s + "000000", 6)

    For i = 1 To Len(s)
	If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then
	    DayCount = -1
	    Exit Function
	End If
    Next i
    
    For i = 1 To Len(s) Step 2
	n = Val(Mid$(s, i, 2))
	Select Case i
	Case 1
	    isleapyear = (n <> 0) And (n Mod 4 = 0)
	    If n < 90 Then
		n = n + 100
	    End If
	    days = (n - 90) * 365
	    days = days + ((n - 89) / 4)
	Case 3
	    If isleapyear = False Then
		days2 = 29
	    Else
		days2 = 28
	    End If

	    Select Case n
	    Case 12             ' 11月までの日数を加算
		days = days + 31 + days2 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30
	    Case 11
		days = days + 31 + days2 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31
	    Case 10
		days = days + 31 + days2 + 31 + 30 + 31 + 30 + 31 + 31 + 30
	    Case 9
		days = days + 31 + days2 + 31 + 30 + 31 + 30 + 31 + 31
	    Case 8
		days = days + 31 + days2 + 31 + 30 + 31 + 30 + 31
	    Case 7
		days = days + 31 + days2 + 31 + 30 + 31 + 30
	    Case 6
		days = days + 31 + days2 + 31 + 30 + 31
	    Case 5
		days = days + 31 + days2 + 31 + 30
	    Case 4
		days = days + 31 + days2 + 31
	    Case 3
		days = days + 31 + days2
	    Case 2
		days = days + 31
	    End Select
	Case 5
	    If n > 31 Then
		DayCount = -1
		Exit Function
	    End If
	    days = days + n
	End Select
    Next i
    DayCount = days
End Function

Function ReplyCount (s$) As Integer
    Dim i, a As Integer
    Dim C$

    i = 1
    Do While i <= Len(s$)
	a = Asc(Mid$(s$, i, 1))
	If a >= &H30 And a <= &H39 Then
	    C$ = C$ + Chr$(a)
	    i = i + 1
	End If
	If Chr$(a) = ":" And i <> 1 Then
	    If Val(C$) <> 0 Then
		ReplyCount = Val(C$)
	    Else
		ReplyCount = -1
	    End If
	    Exit Function
	End If
	If Not (a >= &H30 And a <= &H39) And Chr$(a) <> ":" Then
	    ReplyCount = -1
	    Exit Function
	End If
    Loop
    ReplyCount = -1
End Function

Function ReplyTitleMake (s$) As String
    Dim s1$, t$, t1$
    Dim i, C, c1, c2, tmp As Integer

    s1$ = UCase$(s$)

    i = 1
    C = 1
    If Mid$(s1$, 1, 2) <> "RE" Then
	ReplyTitleMake = "RE:" + s$
	Exit Function
    End If
    Do While i <= Len(s$)
	If Mid$(s1$, i, 1) = " " Then
	    i = i + 1
	End If
	If Mid$(s1$, i, 2) = "RE" Then
	    i = i + 2
	    Select Case Asc(Mid$(s1$, i, 1))
	    Case Asc(":")
		C = C + 1
		c2 = i
	    Case &H30 To &H39
		c1 = ReplyCount(Mid$(s1$, i))
		If c1 = -1 Then
		    t$ = Mid$(s$, c2)
		    If C <= 1 Then
			t1$ = ""
		    Else
			t1$ = Format$(C)
		    End If
'                    ReplyTitleMake = "RE" + t1$ + ":" + t$
		    ReplyTitleMake = "RE:" + t$
		    Exit Function
		End If
		C = C + c1
		i = InStr(i, s1$, ":")
		c2 = i
	    Case Asc("*"), Asc("^")
		If Asc(Mid$(s1$, i + 1, 1)) >= &H30 And Asc(Mid$(s1$, i + 1, 1)) <= &H39 Then
		    c1 = ReplyCount(Mid$(s1$, i + 1))
		    If c1 = -1 Then
			t$ = Mid$(s$, c2)
			If C <= 1 Then
			    t1$ = ""
			Else
			    t1$ = Format$(C)
			End If
'                        ReplyTitleMake = "RE" + t1$ + ":" + t$
			ReplyTitleMake = "RE:" + t$
			Exit Function
		    End If
		    C = C + c1
		    i = InStr(i + 1, s1$, ":")
		    c2 = i
		End If
	    Case Else
		i = i - 2
		t$ = Mid$(s$, i)
		If C <= 1 Then
		    t1$ = ""
		Else
		    t1$ = Format$(C)
		End If
'                ReplyTitleMake = "RE" + t1$ + ":" + t$
		ReplyTitleMake = "RE:" + t$
		Exit Function
	    End Select
	End If
	i = i + 1
    Loop
    t$ = Mid$(s$, c2 + 1)
    If C <= 1 Then
	t1$ = ""
    Else
	t1$ = Format$(C)
    End If
'    ReplyTitleMake = "RE" + t1$ + ":" + t$
    ReplyTitleMake = "RE:" + t$
End Function

'
' 文字列の右側からCHR$(0)以外の文字が見つかるまでCHR$(0)を
' 削除する。
'
Function RNullCut (s As String) As String
    For i% = Len(s) To 1 Step -1
	If Mid$(s, i%, 1) <> Chr$(0) Then
	    Exit For
	End If
    Next i%
    RNullCut = Left$(s, i%)
End Function

Sub SetStatBar1 (s As String)
    Trace "SetStatBar1"
    frmMain!pnl3DStatBar1.Caption = s
End Sub

' 文字列からスペースを取る
'
Function SpaceCut (s As String) As String
    Dim i As Integer
    Dim s1 As String
    Dim ss As String

    Trace "SpaceCut"
    
    s1 = ""
    ss = ""

    For i = 1 To Len(s)
	s1 = Mid$(s, i, 1)
	If s1 = Chr$(0) Then
	    SpaceCut = ss
	    Exit Function
	End If
	If s1 <> " " Then
	    ss = ss + s1
	End If
    Next i
    SpaceCut = ss
End Function

' ルートヘッダーからコールサインを抽出する１
'
Function HeadCut (s As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim ss As String
    Dim s1 As String
    
    Trace "HeadCut"
    
    ss = s + Chr$(0)
    i = 1
    Do While (Mid$(ss, i, 1) <> "@")
	i = i + 1
    Loop
' RLI方式のヘッダー
    If Mid$(ss, i + 1, 1) = ":" Then
	i = i + 2
    Else
' MBL方式のヘッダー
	i = i + 1
    End If
    j = i

    Do
	s1 = Mid$(ss, i, 1)
	If s1 = "." Then
	    Exit Do
	End If
	If s1 = " " Then
	    Exit Do
	End If
	If s1 = Chr$(0) Then
	    Exit Do
	End If
	i = i + 1
    Loop

    HeadCut = Mid(ss, j, i - j)
End Function

' ルートヘッダーからコールサインを抽出する２
'
Function HeadCut2 (s As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim ss As String
    Dim s1 As String
    
    Trace "HeadCut2"
    
    ss = s + Chr$(0)
    i = 1
    Do While (Mid$(ss, i, 1) <> "@")
	i = i + 1
    Loop
' MBL方式のヘッダー
    If Mid$(ss, i + 1, 1) = ":" Then
	i = i + 2
    Else
	i = i + 1
    End If
    j = i

    Do
	s1 = Mid$(ss, i, 1)
	If s1 = " " Then
	    Exit Do
	End If
	If s1 = Chr$(0) Then
	    Exit Do
	End If
	i = i + 1
    Loop

    HeadCut2 = Mid(ss, j, i - j)

End Function

' ルートヘッダーから日付を抽出する
'
Function HeadDate (s As String) As String
    Dim ss As String
    Dim i As Integer

    Trace "HeadDate"
    
    If Mid$(s, 1, 1) <> "R" Then
	HeadDate = Null
	Exit Function
    End If
    i = 1
    Do While (Mid$(s, i, 1) <> "/")
	i = i + 1
    Loop
    HeadDate = Mid$(s, 3, i - 3)
End Function

' ルートヘッダーからメッセージ番号を抽出する
'
Function HeadMsgNum (s As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim ss As String
    Dim num As String

    Trace "HeadMsgNum"
    
    ss = s + Chr$(0)
    i = 1
    j = 1
    Do While (Mid$(ss, i, 1) <> "@")
	i = i + 1
    Loop
' RLI方式のヘッダー
    If Mid$(ss, i + 1, 1) = ":" Then
	Do While (Mid$(ss, j, 1) <> "#")
	    j = j + 1
	Loop
	j = j + 2
	Do While (Mid$(ss, j, 1) <> " ")
	    num = num + Mid$(ss, j, 1)
	    j = j + 1
	Loop
    Else
' MBL方式のヘッダー
	Do While (Mid$(ss, j, 1) <> " ")
	    j = j + 1
	Loop
	j = j + 1
	Do While (Mid$(ss, j, 1) <> "@")
	    num = num + Mid$(ss, j, 1)
	    j = j + 1
	Loop
    End If
    HeadMsgNum = num
End Function

' ルートヘッダーから時間を抽出する
'
Function HeadTime (s As String) As String
    Dim ss As String
    Dim i As Integer
    Dim j As Integer

    Trace "HeadTime"
    
    If Mid$(s, 1, 1) <> "R" Then
	HeadTime = Null
	Exit Function
    End If
    i = 1
    Do While (Mid$(s, i, 1) <> "/")
	i = i + 1
    Loop
    j = i
    Do While (Mid$(s, j, 1) <> " ")
	j = j + 1
    Loop
    
    HeadTime = Mid$(s, i + 1, j - i - 1)
End Function

Sub Idol_off ()
    frmPanel!pnlInfo.Caption = ""
    If L_MODE = 0 Then
	frmMain!pnl3DStatBar1.Caption = "通信中"
    Else
	frmMain!pnl3DStatBar1.Caption = "Now Communication"
    End If
    frmPanel!cmdComStop.Enabled = True
End Sub

Sub Idol_on ()
    If L_MODE = 0 Then
	frmPanel!pnlInfo.Caption = "非通信中"
	frmMain!pnl3DStatBar1.Caption = "非通信中"
    Else
	frmPanel!pnlInfo.Caption = ""
	frmMain!pnl3DStatBar1.Caption = ""
    End If
    frmPanel!pnlToField.Caption = ""
    frmPanel!pnlAtbbs.Caption = ""
    frmPanel!pnlCall.Caption = ""
    frmPanel!pnlTitle.Caption = ""
    frmPanel!pnlRcv.Caption = ""
    frmPanel!cmdComStop.Enabled = False
End Sub

Function Iskanji (C As Integer) As Integer
    Select Case C
	Case &H81 To &H9F, &HE0 To &HFC
	    Iskanji = True
	Case Else
	    Iskanji = False
    End Select
End Function

'
' 文字列の左側からCHR$(0)以外の文字が見つかるまでCHR$(0)を
' 削除する。
'
Function LNullCut (s As String) As String
    For i% = 1 To Len(s)
	If Mid$(s, i%, 1) = Chr$(0) Then
	    Exit For
	End If
    Next i%
    LNullCut = Left$(s, i% - 1)
End Function

Sub MessageCut (s$, msg As MSGD2)
    Dim i As Integer
    Dim j As Integer
    ReDim s1(20) As String
    Dim l As Integer

    Trace "MessageCut"
    
    i = InStr(s$, Chr$(9))
    msg.fname = Mid$(s$, 1, i - 1)
    
    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.mediatype = Mid$(s$, j, i - j)

    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.savemode = Mid$(s$, j, i - j)
	
    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.stat = Mid$(s$, j, i - j)
	
    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.size = Mid$(s$, j, i - j)
	
    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.field = Mid$(s$, j, i - j)
	
    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.atbbs = Mid$(s$, j, i - j)
	
    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.Call = Mid$(s$, j, i - j)
	
    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.date = Mid$(s$, j, i - j)
	
    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.time = Mid$(s$, j, i - j)
	
    j = i + 1
    i = InStr(j, s$, Chr$(9))
    msg.bid = Mid$(s$, j, i - j)
	
    msg.Title = Mid$(s$, i + 1)
End Sub

Sub MouseNormal ()
    Trace "MouseNormal"
    screen.MousePointer = 0
End Sub

Sub MouseWait ()
    Trace "MouseWait"
    screen.MousePointer = 11
End Sub

Sub PrtInfo (s$)
    Dim i%, tmlen%, strlen%, crpos%, k%, kflag%
    Dim tmDisp$, tmDisp1$, tmp$
    Static sttmDisp$

'    Exit Sub
    On Error GoTo PrtInfo_Err

    s$ = sttmDisp$ + s$
    tmlen% = Len(frmMonitor!txtMonitor.Text)
    strlen% = Len(s$)
    If tmlen% + strlen% >= 30000 Then
	frmMonitor!txtMonitor.SelStart = 0
	frmMonitor!txtMonitor.SelLength = 6000  'B175 3000 -> 6000
	frmMonitor!txtMonitor.SelText = ""
    End If
    For i% = 1 To Len(s$)
	tmp$ = Mid$(s$, i%, 1)
	If tmp$ <> Chr$(10) Then
	    tmDisp$ = tmDisp$ + tmp$
	End If
	If tmp$ = Chr$(13) Then
	    tmDisp$ = tmDisp$ + Chr$(10)
	End If
    Next i%
    
    kflag% = False
    tmDisp1$ = tmDisp$
    For i% = 1 To Len(tmDisp$)
	k% = Asc(Mid$(tmDisp1$, i%, 1))
	If Iskanji(k%) = True Then
	    ' 既に漢字フラグが立っている場合は漢字の2バイト目と識別する
	    If kflag% = True Then
		kflag% = False
	    Else
		kflag% = True
	    End If
	Else
	    kflag% = False
	End If
    Next i%
    ' 漢字フラグが立っている場合は最終文字が漢字コードの１バイト目なのでバッファに保存する
    If kflag% = True Then
	sttmDisp$ = Right$(tmDisp$, 1)
	If Len(tmDisp$) <= 1 Then
	    tmDisp$ = ""
	Else
	    tmDisp$ = Left$(tmDisp$, Len(tmDisp$) - 1)
	End If
    Else
	sttmDisp$ = ""
    End If

    tmDisp$ = CharFilter(tmDisp$)

    frmMonitor!txtMonitor.SelStart = Len(frmMonitor!txtMonitor.Text)
5000    frmMonitor!txtMonitor.SelText = tmDisp$
    Exit Sub

PrtInfo_Err:
    If Erl = 5000 Then
	frmMonitor!txtMonitor.Text = ""
	tmDisp$ = ""
	Resume Next
    Else
	Debug.Print tmDisp$
	ErrorProc "PrtInfo", Err, Erl, Error
	FileAdd USERDIR + "LOG.TXT", "PrtInfo>Error " + Format$(Err) + ":" + Format$(Erl) + ": Str" + tmDisp$
	Resume Next
    End If
End Sub

'
' 文字列検索(ワイルドカード対応)
' s1からs2を検索する
'
' 戻り値:一致 0,不一致 -1
'
Function StrCmp (s1 As String, s2 As String)
'    Trace "StrCmp"
    j% = 1
    StrCmp = 0
    ' s2が空文字列のときは検索不一致とする
    If Len(s2) = 0 Then
	StrCmp = -1
	Exit Function
    End If
    ' s2よりs1のサイズが大きくてS2にワイルドカード文字列"*"を含んでいない場合は検索不一致とする
    If Len(s1) > Len(s2) And InStr(s2, "*") = 0 Then
	StrCmp = -1
	Exit Function
    End If

    For i% = 1 To Len(s2)
	If j% > Len(s1) Then
	    If Mid$(s2, j%, 1) = "*" Then
		StrCmp = 0
		Exit For
	    Else
		StrCmp = -1
		Exit For
	    End If
	End If
	s11$ = Mid$(s1, j%, 1)
	s21$ = Mid$(s2, i%, 1)
	If s21$ = "*" Then
	    StrCmp = 0
	    Exit For
	End If
	If (s11$ <> s21$) And s21$ <> "?" Then
	    StrCmp = -1
	    Exit For
	End If
	j% = j% + 1
    Next i%
End Function

Function StrCnv (s As String) As String
    Dim a%
    Dim s1 As String * 10000

    a% = SetNkfOption("s" + Chr$(0) + Chr$(0))
    NkfConvert s1, s + Chr$(0) + Chr$(0)
    StrCnv = s1
End Function

Sub ToolBarDisable ()
    Trace "ToolBarDisable"
    frmMain!cmdBull.Enabled = False
    frmMain!cmdMail.Enabled = False
    frmMain!cmdFollowup.Enabled = False
    frmMain!cmdReply.Enabled = False
End Sub

Sub ToolBarEnable ()
    Trace "ToolBarEnable"
    frmMain!cmdBull.Enabled = True
    frmMain!cmdMail.Enabled = True
    frmMain!cmdFollowup.Enabled = True
    frmMain!cmdReply.Enabled = True
End Sub

' ＲＦＣ－８２２規格準拠の時刻を設定
'
Function Getrfctime (d As String, t As String)
    Dim i As Integer
    Dim tz$, z$

    Trace "Getrfctime"
    
    ' 時刻設定
    t1$ = Left$(t, 2) + ":" + Mid$(t, 3, 2)
    z$ = UCase$(Right$(t, 1))
    If Len(t1$) < 5 Then
	tz$ = ""
    Else
	Select Case Asc(z$)
	Case Asc("Z"), Asc("J")
	    tz$ = ""
	Case Asc("A") To Asc("I")
	    i = Asc(z$) - &H40
	    tz$ = "+" + Right$("00" + Format$(i), 2) + "00"
	Case Asc("K")
	    tz$ = "+1000"
	Case Asc("L")
	    tz$ = "+1100"
	Case Asc("M")
	    tz$ = "+1200"
	Case Asc("N") To Asc("Y")
	    i = Asc(z$) - Asc("M")
	    tz$ = "-" + Right$("00" + Format$(i), 2) + "00"
	Case Else
	    tz$ = ""
	End Select
    End If
    t1$ = t1$ + " " + tz$
    ' 日付設定
    d1$ = Left$(d, 2) + "/" + Mid$(d, 3, 2) + "/" + Mid$(d, 5, 2)
    d1$ = Format$(d1$, "ddd, dd mmm yy")
    Getrfctime = d1$ + " " + t1$
End Function

