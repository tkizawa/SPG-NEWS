VERSION 2.00
Begin Form frmTerminal 
   Caption         =   "�^�[�~�i��"
   ClientHeight    =   7260
   ClientLeft      =   1425
   ClientTop       =   2505
   ClientWidth     =   10290
   Height          =   7665
   Icon            =   TERMINAL.FRX:0000
   Left            =   1365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10290
   Top             =   2160
   Width           =   10410
   Begin PictureBox picTerminal1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      Height          =   660
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   10260
      TabIndex        =   4
      Top             =   6600
      Width           =   10290
      Begin TextBox txtSend 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   60
         Width           =   10155
      End
   End
   Begin PDQComm Comm1 
      AutoScroll      =   4  '�����y�уL�[���͂̐�������
      AutoSize        =   -1  'True
      BackColor       =   1
      Columns         =   80
      Echo            =   -1  'True
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Terminal"
      FontSize        =   13.5
      FontUnderline   =   0   'False
      ForeColor       =   15
      Handshaking     =   2  'RTS/CTS
      Height          =   5400
      InBufferSize    =   32000
      Interval        =   55
      Left            =   9720
      Notification    =   0  '�|�[�����O�@
      OutBufferSize   =   8192
      Rows            =   20
      RThreshold      =   1
      RTSEnable       =   -1  'True
      ScrollRows      =   409
      Top             =   600
      Width           =   9855
      XferCarrierAbort=   0   'False
      XferProtocol    =   1  'XModem CRC
      XferStatusDialog=   1  '���[�h���X�_�C�A���O
   End
   Begin TextBox txtTerminal 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   5415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "SPG-Talk"
      Top             =   420
      Width           =   9615
   End
   Begin PictureBox picTerminal 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   435
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   10260
      TabIndex        =   0
      Top             =   0
      Width           =   10290
      Begin CommandButton Command3 
         Caption         =   "�޳�۰�ޗ\��"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin CommandButton Command2 
         Caption         =   "����f"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin CommandButton Command1 
         Caption         =   "�ڑ�"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
End
Option Explicit

Sub Comm1_OnComm ()
    Dim rcv$
    Dim i%, tmp$, tmDisp$

    Select Case Comm1.CommEvent
    ' ��M�C�x���g�����������烂�j�^�ɕ\������
    Case 2
        rcv$ = Comm1.Input
        PrintTerminal rcv$
    End Select
End Sub

Sub Command1_Click ()
     Dim r%
     
     r% = TerminalConnect()
     If r% = False Then
        PrintTerminal "�R�l�N�g���s" + Chr$(13)  ' + Chr$(10)
     Else
        PrintTerminal "�R�l�N�g����" + Chr$(13)  ' + Chr$(10)
     End If
    txtSend.SetFocus
End Sub

Sub Command2_Click ()
    TerminalDisconnect
    PrintTerminal DISC$ + Chr$(13) ' + Chr$(10)
    txtSend.SetFocus
End Sub

Sub Form_Activate ()
    On Error GoTo com_active_error

    Comm1.PortOpen = True
    TERMFLAG = True
    Exit Sub

com_active_error:
    MsgBox "���ݓ]�����܂��̓I�[�g�p�C���b�g�^�p���ł��B�^�[�~�i���\�t�g�̎g�p�͂ł��܂���B"
    Unload Me
    Exit Sub
End Sub

Sub Form_Deactivate ()
    Comm1.PortOpen = False
    TERMFLAG = False
End Sub

Sub Form_Load ()
    ReadWinPos frmTerminal, "Terminal"
    
    txtTerminal.Text = "SPG-Talk V4.5" + VERSION + Chr$(13) + Chr$(10)
    Comm1.InputLen = 0
    Comm1.CommPort = PORT
    Comm1.Settings = BAUD + "n,8,1"
'    Comm1.PortOpen = True
'    COMFLAG = True
    Exit Sub
End Sub

Sub Form_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= txtTerminal.Top And X <= txtTerminal.Top + txtTerminal.Width And Y >= txtTerminal.Left And Y <= txtTerminal.Left + txtTerminal.Width Then
    '
    End If
End Sub

Sub Form_Resize ()
    On Error Resume Next ' �ŏ����̃v���p�e�B�G���[�����h�~
    
    txtTerminal.Width = scalewidth
    txtTerminal.Height = scaleheight - picTerminal.Height - picTerminal1.Height
    txtSend.Width = picTerminal1.Width - 150
End Sub

Sub Form_Unload (Cancel As Integer)
    On Error Resume Next

    Comm1.PortOpen = False
    TERMFLAG = False
'    COMFLAG = False
    
    If frmTerminal.WindowState <> 1 Then
        WrtWinPos frmTerminal, "Terminal"
    End If
End Sub

Sub PrintTerminal (s$)
    Dim i%, tmlen%, strlen%, crpos%, k%, kflag%
    Dim tmDisp$, tmDisp1$, tmp$
    Static sttmDisp$

    On Error GoTo PrintTerminal_Err

    s$ = sttmDisp$ + s$
    tmlen% = Len(txtTerminal.Text)
    strlen% = Len(s$)
    If tmlen% + strlen% >= 30000 Then
        txtTerminal.SelStart = 0
        txtTerminal.SelLength = 3000
        txtTerminal.SelText = ""
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
            ' ���Ɋ����t���O�������Ă���ꍇ�͊�����2�o�C�g�ڂƎ��ʂ���
            If kflag% = True Then
                kflag% = False
            Else
                kflag% = True
            End If
        Else
            kflag% = False
        End If
    Next i%
    ' �����t���O�������Ă���ꍇ�͍ŏI�����������R�[�h�̂P�o�C�g�ڂȂ̂Ńo�b�t�@�ɕۑ�����
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

'    ' Right$�Ŏ�M������̍Ō�̕����i�S�pOR���p�j���擾����B
'    k% = Asc(Right$(tmDisp$, 1))
'    ' �S�p�����̑O���P�o�C�g�̏ꍇ�o�b�t�@�ɕۑ����A����ȊO�̕������\������
'    If Iskanji(k%) = True Then
'        sttmDisp$ = Chr$(k%)
'        If Len(tmDisp$) <= 1 Then
'            tmDisp$ = ""
'        Else
'            tmDisp$ = Left$(tmDisp$, Len(tmDisp$) - 1)
'        End If
'    Else
'        sttmDisp$ = ""
'    End If
    txtTerminal.SelStart = Len(txtTerminal.Text)
5000    txtTerminal.SelText = tmDisp$
    Exit Sub

PrintTerminal_Err:
    If Erl = 5000 Then
        txtTerminal.Text = ""
        tmDisp$ = ""
        Resume Next
    Else
        Debug.Print tmDisp$
        ErrorProc "PrintTerminal", Err, Erl, Error
        FileAdd USERDIR + "LOG.TXT", "PrintTerminal>Error " + Format$(Err) + ":" + Format$(Erl) + ": Str" + tmDisp$
        Resume Next
    End If
End Sub

Function TerminalConnect () As Integer
    Dim a$, b$, cr$
    Dim i, FileNum As Integer

    cr$ = Chr$(13)
    TerminalConnect = True
    
    ' �R�l�N�g����
    If LINK = "TNC" Then
        Comm1.Output = String$(3, Chr$(3)) + "ECHO OFF" + cr$
        Comm1.Output = String$(3, Chr$(3)) + "AUTOLF OFF" + cr$
        Comm1.Output = String$(3, Chr$(3)) + "c" + cr$
        If TerminalWaitComm("Link", RCVWAITTIME) = False Then
            TerminalDisconnect
            TerminalConnect = False
            Exit Function
        End If
    Else
        Comm1.Output = cr$ + cr$ + cr$
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
                If TerminalWaitComm(b$, RCVWAITTIME) = False Then
                    TerminalDisconnect
                    TerminalConnect = False
                    Close #FileNum
                    Exit Function
                End If
            Else
                Comm1.Output = b$ + cr$
                PrintTerminal b$ + cr$ ' + Chr$(10)
            End If
        Loop
        Close #FileNum
    Else
        TerminalDisconnect
        TerminalConnect = False
        Exit Function
    End If
End Function

Sub TerminalDisconnect ()
    Dim i As Integer
    Dim a%, b$

    If LINK = "TNC" Then
        Comm1.Output = String$(3, Chr$(3)) + "D" + Chr$(13)
    Else
        Comm1.Output = "*** DISCONNECTED" + Chr$(13) + Chr$(10)
    End If
End Sub

'
' �������̕������M
'
' t% = �҂����ԁi���j
'
Function TerminalRcvStr (t%) As String
    Dim a%, buf$, buf1 As String * 1
    
    buf$ = ""
    buf1 = ""
    TIMEOUTFLAG = 0
    TIMECOUNT = 0
    frmTimer!tmrTimer.Interval = 60000   ' �ꕪ��
    frmTimer!tmrTimer.Enabled = True     '�^�C�}�[�X�^�[�g

    ' ��M�҂�
    a% = Comm1.InBufferCount
    Do While (a% = 0)
        If TIMEOUTFLAG <> 0 And TIMECOUNT >= t% Then
            frmTimer!tmrTimer.Enabled = False
            TIMEOUTFLAG = 0
            TIMECOUNT = 0
            TerminalRcvStr = DISC$
            PrintTerminal DISC$
            Exit Function                       ' �^�C���A�E�g�G���[
        End If
        DoEvents
        a% = Comm1.InBufferCount
    Loop
    
    ' ��M�L���[�Ƀf�[�^���������ꍇ��M����������
    ' 1�b�Ԉȏ��M�L���[���󂾂������M�������I��
    TIMEOUTFLAG = 0
    TIMECOUNT = 0
    frmTimer!tmrTimer.Enabled = False
    Do
        a% = Comm1.InBufferCount
        If (a% <> 0) Then
            Comm1.InputLen = 1
            buf1 = Comm1.Input
            If buf1 = Chr$(10) Then
                buf1 = ""
            End If
            buf$ = buf$ + buf1
            If buf1 = Chr$(13) Then         ' ���s�R�[�h�ŏI��
                frmTimer!tmrTimer.Enabled = False
                TIMEOUTFLAG = 0
                TIMECOUNT = 0
                buf$ = buf$ + Chr$(10)
                TerminalRcvStr = buf$
                PrintTerminal buf$
                Comm1.InputLen = 0
                Exit Function
            End If
        Else
            TIMEOUTFLAG = 0
            TIMECOUNT = 0
            frmTimer!tmrTimer.Interval = 1000   ' ��b��
            frmTimer!tmrTimer.Enabled = True     '�^�C�}�[�X�^�[�g
            Do While (a% = 0)
                a% = Comm1.InBufferCount
                If TIMEOUTFLAG <> 0 Then
                    frmTimer!tmrTimer.Enabled = False
                    TIMEOUTFLAG = 0
                    TIMECOUNT = 0
                    TerminalRcvStr = buf$
                    PrtInfo buf$
                    Comm1.InputLen = 0
                    Exit Function
                End If
                DoEvents
            Loop
            frmTimer!tmrTimer.Enabled = False
            TIMEOUTFLAG = 0
            TIMECOUNT = 0
        End If
    Loop
    Comm1.InputLen = 0
End Function

'
' �������M�҂�
'
' s$ = ��M������
' t% = �҂�����
'
Function TerminalWaitComm (s$, t%) As Integer
    Dim c%, rcv$
    
    rcv$ = TerminalRcvStr(t%)
    Do While (InStr(rcv$, s$) = 0)
        rcv$ = TerminalRcvStr(t%)
        If InStr(rcv$, DISC$) <> 0 Then
            TerminalWaitComm = False
            Exit Function
        End If
        c% = c% + 1
        If c% > 10 Then
            TerminalWaitComm = False
            Exit Function
        End If
    Loop
    TerminalWaitComm = True
End Function

Sub txtSend_KeyPress (KeyAscii As Integer)
    Dim send$, i%
    
    send$ = txtSend.Text
    If KeyAscii = 13 Then
'        For i% = 1 To Len(send$)
'            Debug.Print Hex$(Asc(Mid$(send$, i%, 1)))
'        Next i%
        If Left$(send$, 1) = Chr$(13) Then
            If Len(send$) > 2 Then
                send$ = Mid$(send$, 3)
            Else
                send$ = ""
            End If
        End If
        Comm1.Output = send$ + Chr$(13)
        PrintTerminal send$ + Chr$(13)
        txtSend.Text = ""
        Exit Sub
    End If
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If KeyAscii < 32 Then
        Comm1.Output = Chr$(KeyAscii)
      End If
End Sub

Sub txtTerminal_GotFocus ()
'    txtSend.SetFocus
End Sub

Sub txtTerminal_KeyPress (KeyAscii As Integer)
'    Static send$
'
'    Select Case KeyAscii
'    Case 13
'        Comm1.Output = send$ + Chr$(13)
'        send$ = ""
'   Case 8
'        If Len(send$) <> 0 Then
'            send$ = Left$(send$, Len(send$) - 1)
'        End If
'    Case Else
'        send$ = send$ + Chr$(KeyAscii)
'    End Select
End Sub
