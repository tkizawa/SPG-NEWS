' �e�v�c�]���V�X�e��
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA
Option Explicit

'
' �r�R�}���h���
'
Sub CmdAna (s As String, msg As MSGD2)
    Dim i As Integer
    Dim j As Integer
    ReDim s1(20) As String
    Dim l As Integer

    Trace "CmdAna"

    ' �R�}���h���C������
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
    ' �X�e�[�^�X�i�a�^�o�j
    ' To Field�ݒ�
    msg.bid = ""
    msg.stat = UCase$(Mid$(s1(0), 2, 1) + "F")
    msg.field = UCase$(s1(1))

    ' @BBS�ݒ�
    For i = 2 To j
        If s1(i) = "@" Then
            msg.atbbs = UCase$(s1(i + 1))
            Exit For
        End If
    Next i

    ' From�ݒ�
    For i = 2 To j
        If s1(i) = "<" Then
            msg.Call = UCase$(s1(i + 1))
            Exit For
        End If
    Next i

    ' BID�ݒ�
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
    
    ' �^�[�~�i���\�t�g�g�p���͓]�����Ȃ�
    If TERMFLAG = True Then
        Exit Sub
    End If

    ' �R�l�N�g
    frmPanel!pnlInfo.Caption = "�e�v�c�]�� �R�l�N�g��"
    Idol_off
10  If Connect() <> 0 Then
        Exit Sub
    End If
    ' �r�h�c��M�҂�
    frmPanel!pnlInfo.Caption = "�e�v�c�]�� �r�h�c��M�҂�"
20  If (WaitComm("[", RCVWAITTIME) = DISC$) Then
        Disconnect
        Idol_on
        Exit Sub
    End If
    ' �v�����v�g��M�҂�
    frmPanel!pnlInfo.Caption = "�e�v�c�]�� �v�����v�g�҂�"
30  If (WaitComm(">", RCVWAITTIME) = DISC$) Then
        Disconnect
        Idol_on
        Exit Sub
    End If
    
    ' �r�h�c���M
    frmPanel!pnlInfo.Caption = "�e�v�c�]�� �r�h�c���M"
40  a$ = SendStr("[SPG-4.3-$]", 0, True)
    If InStr(a$, DISC$) <> 0 Then
        Disconnect
        Idol_on
        Exit Sub
    End If

    ' �v�����v�g��M�҂�
    frmPanel!pnlInfo.Caption = "�e�v�c�]�� �v�����v�g�҂�"
50  If (WaitComm(">", RCVWAITTIME) = DISC$) Then
        Disconnect
        Idol_on
        Exit Sub
    End If
    
    ' ���M
    MaxNum% = UBound(Message)

    For i% = 0 To MaxNum%
        stat$ = Mid$(Message(i%).stat, 2, 1)
        If (stat$ <> "") And (stat$ <> "K") And (stat$ <> "F") And Mid$(Message(i%).stat, 3, 1) <> "K" And Mid$(Message(i%).stat, 3, 1) <> "H" Then
            ' ���M�R�}���h���M
            frmPanel!pnlInfo.Caption = "�e�v�c�]�� ���M�R�}���h���M"
            frmPanel!pnlToField.Caption = Message(i%).field
            frmPanel!pnlAtbbs.Caption = Message(i%).atbbs
            frmPanel!pnlCall.Caption = Message(i%).Call
            frmPanel!pnlTitle.Caption = Message(i%).Title
60          If (FWDSendCmd(i%, True) <> 0) Then
                Disconnect
                Idol_on
                Exit Sub
            End If
            ' OK/NO�҂�
            frmPanel!pnlInfo.Caption = "�e�v�c�]�� OK/NO�҂�"
            r$ = RcvStr(RCVWAITTIME, True)
            If Left$(r$, 1) <> "N" Then
                ' �t�@�C�����M
                frmPanel!pnlInfo.Caption = "�e�v�c�]�� ���b�Z�[�W���M"
70              If (FWDSendFile(i%) <> 0) Then
                    Disconnect
                    Idol_on
                    Exit Sub
                End If
            End If
            ' �v�����v�g��M�҂�
            frmPanel!pnlInfo.Caption = "�e�v�c�]�� �v�����v�g�҂�"
80          If (WaitComm(">", RCVWAITTIME) = DISC$) Then
                Disconnect
                Idol_on
                Exit Sub
            End If
            ' �X�e�[�^�X�X�V
            Mid$(Message(i%).stat, 2) = "F"
        End If
        ClearPanel
    Next i%

    ' ��M
    frmPanel!pnlInfo.Caption = "�e�v�c�]�� ��M��"
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
        MsgBox "FWD�]�����ɃG���[���������܂����B" + Chr$(13) + Chr$(10) + "�ʐM�������I�����܂�", MB_ICONEXCLAMATION, "SPG-NEWS"
    Else
        MsgBox "FWD Error." + Chr$(13) + Chr$(10) + "FWD Abort.", MB_ICONEXCLAMATION, "SPG-NEWS"
    End If
    FileAdd USERDIR + "LOG.TXT", "FWD>Error " + Format$(Err) + ":" + Format$(Erl)
    Disconnect
    Idol_on
    Exit Sub
End Sub

'
' �e�v�c��M
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
    '�^�C�g����M
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

    ' ��M�t�@�C���I�[�v��
    FileNum = FreeFile
    Open "RCV.TXT" For Output As FileNum
    ' ���[���{�̎�M
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

    ' �t�@�C����������
    If (FileWrite2("RCV.TXT", msg.stat, msg.field, msg.atbbs, msg.Call, msg.Title, msg.bid, "***", False) <> 0) Then
        FWDRcv = -1
        Exit Function
    End If
    FWDRcv = 0
End Function

'
' �e�v�c���M�R�}���h���M
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
' �e�v�c�t�@�C�����M
'
Function FWDSendFile (i As Integer) As Integer
    Dim FileNum As Integer
    Dim Filename As String
    Dim a$, send$

    Trace "FWDSendFile"
    
    ' �^�C�g�����M
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
' �R�l�N�g
' �ڑ�����v�����v�g���Ԃ��Ă���܂�
'
Function Connect () As Integer
    Dim ss As String * 100
    Dim sr As String * 100
    Dim i, FileNum As Integer
    Dim a$, b$, rr$, s1$, s2$, a2$
    Dim a1%

    frmPanel!pnlInfo.Caption = "�ʐM����I�[�v��"
    s1$ = "[���M��] "
    s2$ = "[��M�҂�] "
    If (OpenPort() <> 0) Then
        frmPanel!pnlInfo.Caption = "�ʐM����I�[�v�����s"
        Connect = -1
        Exit Function
    End If
    
    i = FlushComm(nCid, 1)
    i = Dtr_on()

    ' �R�l�N�g����
    If LINK = "TNC" Then
        frmPanel!pnlInfo.Caption = "�s�m�b�ݒ�"
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

    frmPanel!pnlInfo.Caption = "�R�l�N�g����"
    Connect = 0
End Function
