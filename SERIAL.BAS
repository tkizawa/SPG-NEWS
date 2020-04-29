'Option Explicit
'----------------------------------------------------------------------------
'    Serial Communications Module for VB
'----------------------------------------------------------------------------
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA
'
' �q�r�|�Q�R�Q�b�\���̒�`
'
Global TPort As Integer
Global TBaud As Integer
Global TWord As Integer
Global TParity As String
Global TStopBit As Integer
Global TConfigChanged As Integer

Type DCB
    id As String * 1
    baudrate As Integer
    bytesize As String * 1
    Parity As String * 1
    stopbits As String * 1
    rlstimeout As Integer
    ctstimeout As Integer
    dsrtimeout As Integer

    bits1 As String * 1 ' The fifteen actual DCB bit-sized data fields
    bits2 As String * 1 ' within these two bytes can be manipulated by
              ' bitwise logical And/Or operations.  Refer to
              ' SDKWIN.HLP for location/meaning of specific bits

    XonChar As String * 1
    XoffChar As String * 1
    XonLim As Integer
    XoffLim As Integer
    PeChar As String * 1
    EofChar As String * 1
    EvtChar As String * 1
    TxDelay As Integer
End Type

Type COMSTAT
    bits As String * 1 ' For specific bit flags and their
             ' meanings, refer to SDKWIN.HLP.
    cbInQue As Integer
    cbOutQue As Integer
End Type

'----------------------------------------------------------------------------
'
'  COMM declarations
'
'----------------------------------------------------------------------------
Const NOPARITY = 0
Const ODDPARITY = 1
Const EVENPARITY = 2
Const MARKPARITY = 3
Const SPACEPARITY = 4

Const ONESTOPBIT = 0
Const ONE5STOPBITS = 1
Const TWOSTOPBITS = 2

Const IGNORE = 0 '  Ignore signal
Const INFINITE = &HFFFF  '  Infinite timeout

'----------------------------------------------------------------------------
'  Error Flags
'----------------------------------------------------------------------------
Const CE_RXOVER = &H1    '  Receive Queue overflow
Const CE_OVERRUN = &H2   '  Receive Overrun Error
Const CE_RXPARITY = &H4  '  Receive Parity Error
Const CE_FRAME = &H8     '  Receive Framing error
Const CE_BREAK = &H10    '  Break Detected
Const CE_CTSTO = &H20    '  CTS Timeout
Const CE_DSRTO = &H40    '  DSR Timeout
Const CE_RLSDTO = &H80   '  RLSD Timeout
Const CE_TXFULL = &H100  '  TX Queue is full
Const CE_PTO = &H200     '  LPTx Timeout
Const CE_IOE = &H400     '  LPTx I/O Error
Const CE_DNS = &H800     '  LPTx Device not selected
Const CE_OOP = &H1000    '  LPTx Out-Of-Paper
Const CE_MODE = &H8000   '  Requested mode unsupported

Const IE_BADID = (-1)    '  Invalid or unsupported id
Const IE_OPEN = (-2)     '  Device Already Open
Const IE_NOPEN = (-3)    '  Device Not Open
Const IE_MEMORY = (-4)   '  Unable to allocate queues
Const IE_DEFAULT = (-5)  '  Error in default parameters
Const IE_HARDWARE = (-10)        '  Hardware Not Present
Const IE_BYTESIZE = (-11)        '  Illegal Byte Size
Const IE_BAUDRATE = (-12)        '  Unsupported BaudRate

'----------------------------------------------------------------------------
'  Events
'----------------------------------------------------------------------------
Const EV_RXCHAR = &H1    '  Any Character received
Const EV_RXFLAG = &H2    '  Received certain character
Const EV_TXEMPTY = &H4   '  Transmitt Queue Empty
Const EV_CTS = &H8       '  CTS changed state
Const EV_DSR = &H10      '  DSR changed state
Const EV_RLSD = &H20     '  RLSD changed state
Const EV_BREAK = &H40    '  BREAK received
Const EV_ERR = &H80      '  Line status error occurred
Const EV_RING = &H100    '  Ring signal detected
Const EV_PERR = &H200    '  Printer error occured

'----------------------------------------------------------------------------
'  Escape Functions
'----------------------------------------------------------------------------
Const SETXOFF = 1        '  Simulate XOFF received
Const SETXON = 2 '  Simulate XON received
Const SETRTS = 3 '  Set RTS high
Const CLRRTS = 4 '  Set RTS low
Const SETDTR = 5 '  Set DTR high
Const CLRDTR = 6 '  Set DTR low
Const RESETDEV = 7       '  Reset device if possible

Const LPTx = &H80        '  Set if ID is for LPT device


'----------------------------------------------------------------------------
'   Function Definitions
'----------------------------------------------------------------------------
Declare Function OpenComm Lib "User" (ByVal lpComName As String, ByVal wInQueue As Integer, ByVal wOutQueue As Integer) As Integer
Declare Function SetCommState Lib "User" (lpdcb As DCB) As Integer
Declare Function GetCommState Lib "User" (ByVal nCid As Integer, lpdcb As DCB) As Integer
Declare Function ReadComm Lib "User" (ByVal nCid As Integer, ByVal lpBuf As String, ByVal nSize As Integer) As Integer
Declare Function UngetCommChar Lib "User" (ByVal nCid As Integer, ByVal cChar As Integer) As Integer
Declare Function WriteComm Lib "User" (ByVal nCid As Integer, ByVal lpBuf As String, ByVal nSize As Integer) As Integer
Declare Function CloseComm Lib "User" (ByVal nCid As Integer) As Integer
Declare Function BuildCommDCB Lib "User" (ByVal lpDef As String, lpdcb As DCB) As Integer
Declare Function TransmitCommChar Lib "User" (ByVal nCid As Integer, ByVal cChar As Integer) As Integer
Declare Function SetCommEventMask Lib "User" (ByVal nCid As Integer, nEvtMask As Integer) As Long
Declare Function GetCommEventMask Lib "User" (ByVal nCid As Integer, ByVal nEvtMask As Integer) As Integer
Declare Function SetCommBreak Lib "User" (ByVal nCid As Integer) As Integer
Declare Function ClearCommBreak Lib "User" (ByVal nCid As Integer) As Integer
Declare Function FlushComm Lib "User" (ByVal nCid As Integer, ByVal nQueue As Integer) As Integer
Declare Function EscapeCommFunction Lib "User" (ByVal nCid As Integer, ByVal nFunc As Integer) As Integer
Declare Function GetCommError Lib "User" (ByVal nCid As Integer, lpStat As COMSTAT) As Integer

'----------------------------------------------------------------------------
'   Bits  for bits1 and bits2
'----------------------------------------------------------------------------
'   Bits1
'----------------------------------------------------------------------------
Const fbinary = &H1
Const frtsdiable = &H2
Const fparity = &H4
Const foutxctsflow = &H8
Const foutxdsrflow = &H10
Const fdtrdisable = &H80

'----------------------------------------------------------------------------
'   Bits2
'----------------------------------------------------------------------------
Const foutx = &H1
Const finx = &H2
Const fpechar = &H4
Const fnull = &H8
Const fchevt = &H10
Const fdtrflow = &H20
Const frtsflow = &H40

'----------------------------------------------------------------------------
'   Definitions of our open port
'----------------------------------------------------------------------------
Global nCid      As Integer
Global PortName    As String

Sub SendDelayWait (j As Long)
    Dim i As Long

    For i = 0 To j
    Next i
End Sub

' �����񑗐M
' send$:���M������
' tnc%:"1" TNC,"0" �V���A�������N
' disp%:True �\������
'       False �\�����Ȃ�
'
' ���M�L���[����ɂȂ�܂ő��M���Ȃ��悤�ɂ���(H7.2.19)
' CTS�^�C���A�E�g�̊Ď�������悤�ɂ����B(H7.3.26) V3.35.5
'
Function SendStr (send$, tnc%, disp%) As String
    Dim st As COMSTAT
    Dim buf1 As String * 1
    
    ' �ʐM�|�[�g�N���[�Y�Ȃ�ΏI������
    If COMFLAG = False Then
        SendStr = DISC$
        FileAdd USERDIR + "LOG.TXT", "�����I��"
        Exit Function
    End If
    
    ' ��M�f�[�^������Ύ�M����
    a% = GetCommError(nCid, st)
    If (st.cbInQue <> 0) Then
        r$ = RcvStr(RCVWAITTIME, True)
        SendStr = r$
    Else
        SendStr = ""
    End If

    ' �^�C�}�[�Z�b�g
    TIMEOUTFLAG = 0
    TIMECOUNT = 0
    frmTimer!tmrTimer.Interval = 60000   ' �ꕪ��

    ' TNC�ւ̃R�}���h���M�̏ꍇ��0x03���O�񑗂�
    If tnc% <> 0 Then
        send$ = String$(3, Chr$(3)) + send$
    End If
    
    ' ���M�L���[����ɂȂ�܂ő҂�
    frmTimer!tmrTimer.Enabled = True     '�^�C�}�[�X�^�[�g
    a% = GetCommError(nCid, st)
    ' �O���ԑ��M�L���[����ɂȂ�Ȃ������瑗�M�^�C���A�E�g
    Do While (st.cbOutQue <> 0)
        If (TIMEOUTFLAG <> 0 And TIMECOUNT >= SENDWAITTIME) Or COMFLAG = False Then ' ���M�^�C���A�E�g����
            TIMEOUTFLAG = 0
            TIMECOUNT = 0
            frmTimer!tmrTimer.Enabled = False
            SendStr = DISC$
            PrtInfo "���M�^�C���A�E�g"
            Exit Function
        End If
        a% = GetCommError(nCid, st)
        DoEvents
    Loop
    ' �����񑗐M
    send$ = send$ + Chr$(13)
    For i% = 1 To Len(send$)
        s$ = Mid$(send$, i%, 1)
        ' �^�C�}�[�Z�b�g
        TIMEOUTFLAG = 0
        TIMECOUNT = 0
        frmTimer!tmrTimer.Interval = 60000  ' �ꕪ��
        ' ���M�^�C���A�E�g���Ԉȏ�CTS��ON�ɂȂ�Ȃ������瑗�M�^�C���A�E�g
        Do
            If (TIMEOUTFLAG <> 0 And TIMECOUNT >= SENDWAITTIME) Or COMFLAG = False Then  ' B202
                TIMEOUTFLAG = 0
                TIMECOUNT = 0
                frmTimer!tmrTimer.Enabled = False
                SendStr = DISC$
                PrtInfo "���M�^�C���A�E�g(CTS-OFF)" + Chr$(13) + Chr$(10)
                Exit Function
            End If
            DoEvents
            a% = WriteComm(nCid, s$ + Chr$(0) + Chr$(0), 1)
            b% = GetCommError(nCid, st)
'            PrtInfo "Status=" + Format$(b%) + " Buffer=" + Format$(st.cbOutQue) + Chr$(13)
        Loop Until b% = 0 ' b%=32 ��CTS�^�C���A�E�g�G���[
        SendDelayWait (SENDDELAY)
    Next i%
    
    If disp% = True Then
        PrtInfo send$
        FileAdd USERDIR + "LOG.TXT", "SEND>" + send$
    End If
End Function

'
'�@�ʐM�|�[�g�N���[�Y
'
'�@�߂�l�F����I���͂O
'�@�@�@�@�@�ُ�I���͕��̒l
'
Function SerialClose () As Integer
    x% = CloseComm(nCid)
    COMFLAG = False
    If (x% < 0) Then
        SerialClose = x%
    Else
        SerialClose = 0
    End If
End Function

'
'�@�ʐM�|�[�g�ݒ�
'�@�@�@�I�[�v����ɏ������邱��
'
'�@�������F�{�[���[�g�C�f�[�^�T�C�Y(bit)�C�p���e�B,�X�g�b�v�r�b�g
'�@�i��j�@�@(9600)�@�@�@�@�@(8)�@�@�@�@�@�@(n)        (1)�@�@�@
'
Function SerialConfig (BAUD%, bits%, Parity$, stopbits%) As Integer
    Dim lpdcb As DCB
    Dim ConfigString As String

    ConfigString = PortName + ":"
    ConfigString = ConfigString + Format$(BAUD%) + ","
    ConfigString = ConfigString + Left$(UCase$(Parity$), 1) + ","
    ConfigString = ConfigString + Format$(bits%, "#") + ",1"
    i% = BuildCommDCB(ConfigString, lpdcb)

    lpdcb.id = Chr$(nCid)                       ' �@�����f�o�C�X���ʎq
    lpdcb.stopbits = Chr$(stopbits%)
    lpdcb.ctstimeout = 1 ' CTS�^�C���A�E�g���� V3.35.5
    If XFLOWFLAG = 0 Then
        lpdcb.bits1 = Chr$(Asc(lpdcb.bits1) Or fbinary Or foutctsflow) ' ���MCTS�Ď�
        lpdcb.bits2 = Chr$(Asc(lpdcb.bits2) Or frtsflow)    ' ��M��RTS���� �@
        lpdcb.XonChar = Chr$(0)
        lpdcb.XoffChar = Chr$(0)
    Else
        lpdcb.bits1 = Chr$(Asc(lpdcb.bits1) Or fbinary)
        lpdcb.bits2 = Chr$(Asc(lpdcb.bits2) Or foutx Or finx)    ' ��M��X-FLOW���� �@
        lpdcb.XonChar = Chr$(17)
        lpdcb.XoffChar = Chr$(19)
    End If
    lpdcb.XonLim = 256
    lpdcb.XoffLim = 256

    SerialConfig = SetCommState(lpdcb)

End Function

'
'�@�ʐM�|�[�g�I�[�v��
'
Function SerialOpen (ComPort As Integer) As Integer
    PortName = "COM" + Format$(ComPort, "#")
    nCid = OpenComm(PortName, 32767, 20480)
    If (nCid <> 0) Then
        SerialOpen = nCid
    Else
        SerialOpen = 0
    End If
End Function

'
'�@�ʐM�|�[�g��M
'
'�@�������Fbuf$=��M�o�b�t�@
'�@�@�@�@�@max%=�ǂݎ��o�C�g��
'�@�߂�l�F�ǂݎ��o�C�g��
'
Function SerialRead (buf$, ByVal max%) As Integer
    Dim st As COMSTAT
    Dim s As String * 256

    buf$ = Space$(max%)
    i% = ReadComm(nCid, s, max%)

    If (i% > 0) Then
        SerialRead = i%
        buf$ = Left$(s, i%)
    Else
        SerialRead = Abs(i%)
        i% = GetCommError(nCid, st)
    End If
End Function

'
'�@�ʐM�|�[�g���M
'
'�@�������Ft$=���M������
'�@�߂�l�F���M�����o�C�g��
'�@�@�@�@�F���̒l���G���[
'
Function SerialWrite (t$) As Integer

    Dim st As COMSTAT

    status% = GetCommError(nCid, st)
    status% = WriteComm(nCid, t$, Len(t$))
    If status% < 0 Then status% = GetCommError(nCid, st)
    SerialWrite = status%

End Function

'
' �������M�҂�
'
' s$ = ��M������
' t% = �҂�����
'
Function WaitComm (s$, t%) As String
    a$ = RcvStr(t%, True)
    Do While (InStr(a$, s$) = 0)
        a$ = RcvStr(t%, True)
        If InStr(a$, DISC$) <> 0 Then
            WaitComm = DISC$
            Exit Function
        End If
        c% = c% + 1
        If c% > 50 Then
            WaitComm = DISC$
            Exit Function
        End If
    Loop
    WaitComm = a$
End Function

Function Dtr_off ()
    Dim i As Integer
    i = EscapeCommFunction(nCid, CLRDTR)
End Function

Function Dtr_on ()
    Dim i As Integer
    i = EscapeCommFunction(nCid, SETDTR)
End Function

Function OpenPort () As Integer
    a% = SerialOpen(PORT)
    ' �ʐM�|�[�g�I�[�v���G���[����
    If (a% < 0) Then
        Select Case a%
        Case IE_BADID
            MsgBox "Device name invalidity. COM:" + Format$(PORT), 16, "SPG-NEWS"
        Case IE_BAUDRATE
            MsgBox "No support speed. COM:" + Format$(PORT), 16, "SPG-NEWS"
        Case IE_BYTESIZE
            MsgBox "No support data-size. COM:" + Format$(PORT), 16, "SPG-NEWS"
        Case IE_DEFAULT
            MsgBox "Default parameter error COM:" + Format$(PORT), 16, "SPG-NEWS"
        Case IE_HARDWARE
            MsgBox "No port COM:" + Format$(PORT), 16, "SPG-NEWS"
        Case IE_MEMORY
            MsgBox "Communication buffer can't be secured. COM:" + Format$(PORT), 16, "SPG-NEWS"
        Case IE_NOPEN
            MsgBox "Can't open port. COM:" + Format$(PORT), 16, "SPG-NEWS"
        Case IE_OPEN
            MsgBox "Port all ready open. COM:" + Format$(PORT), 16, "SPG-NEWS"
        Case Else
            MsgBox "Undefined error COM:" + Format$(PORT) + "Error Code" + Format$(a%), 16, "SPG-NEWS"
        End Select
        a% = SerialClose()
        COMFLAG = False
        OpenPort = -1
        Exit Function
    End If
    
    b% = Val(BAUD)
    a% = SerialConfig(b%, 8, "N", 0)
    If (a% <> 0) Then
        MsgBox "Can't initialization COM:" + Format$(PORT), 16, "SPG-NEWS"
        a% = SerialClose()
        OpenPort = -1
        Exit Function
    End If
    
    ' �|�[�g�C�l�[�u��
    COMFLAG = True
    OpenPort = 0
End Function

'
' �������̕������M
'
' t% = �҂����ԁi���j
' disp% = True ���j�^�ɕ\��
'
Function RcvStr (t%, disp%) As String
    Dim st As COMSTAT
    Dim buf1 As String * 1
    Dim buf2 As String * 10
    
    On Error GoTo RcvStr_Err

    ' �ʐM�|�[�g�N���[�Y�Ȃ�ΏI������
    If COMFLAG = False Then
        RcvStr = DISC$
        FileAdd USERDIR + "LOG.TXT", "�����I��"
        Exit Function
    End If

    ' ��M�e�ʃI�[�o�[�Ȃ�I������
'    If RCVCOUNT >= RCVLIMIT And Not (RCVLIMIT = -1) Then
'        RcvStr = DISC$
'        FileAdd SYSDIR + "LOG.TXT", "�����I��"
'        Exit Function
'    End If
    
    buf$ = ""
    buf1$ = ""
    TIMEOUTFLAG = 0
    TIMECOUNT = 0
    frmTimer!tmrTimer.Interval = 60000   ' �ꕪ��
    frmTimer!tmrTimer.Enabled = True     '�^�C�}�[�X�^�[�g

    ' ��M�҂�
    a% = GetCommError(nCid, st)
    Do While (st.cbInQue = 0)
        If (TIMEOUTFLAG <> 0 And TIMECOUNT >= t%) Or COMFLAG = False Then
            frmTimer!tmrTimer.Enabled = False
            TIMEOUTFLAG = 0
            TIMECOUNT = 0
            RcvStr = DISC$
2000            PrtInfo DISC$
            Exit Function                       ' �^�C���A�E�g�G���[
        End If
        DoEvents
        a% = GetCommError(nCid, st)
    Loop
    
    ' ��M�L���[�Ƀf�[�^���������ꍇ��M����������
    ' 1�b�Ԉȏ��M�L���[���󂾂������M�������I��
    TIMEOUTFLAG = 0
    TIMECOUNT = 0
    frmTimer!tmrTimer.Enabled = False
    Do
        a% = GetCommError(nCid, st)
        If COMFLAG = False Then
            RcvStr = DISC$
            Exit Function
        End If
        If (st.cbInQue <> 0) Then
'        Debug.Print st.cbInQue
2001            i% = ReadComm(nCid, buf2, 1)
            buf1 = Left$(buf2, 1)
            If buf1 = Chr$(10) Then
                buf1 = ""
            End If
            buf$ = buf$ + buf1
            If buf1 = Chr$(13) Then         ' ���s�R�[�h�ŏI��
                frmTimer!tmrTimer.Enabled = False
                TIMEOUTFLAG = 0
                TIMECOUNT = 0
                buf$ = buf$ + Chr$(10)
                RcvStr = buf$
                If disp% = True Then
2003                    PrtInfo buf$
                    FileAdd USERDIR + "LOG.TXT", "Recive>" + buf$
                End If
                RCVCOUNT = RCVCOUNT + Len(buf$)
                Exit Function
            End If
        Else
            TIMEOUTFLAG = 0
            TIMECOUNT = 0
            frmTimer!tmrTimer.Interval = 1000   ' ��b��
            frmTimer!tmrTimer.Enabled = True     '�^�C�}�[�X�^�[�g
            Do While (st.cbInQue = 0)
                If COMFLAG = False Then
                    frmTimer!tmrTimer.Enabled = False
                    TIMEOUTFLAG = 0
                    TIMECOUNT = 0
                    RcvStr = DISC$
                    Exit Function
                End If
                a% = GetCommError(nCid, st)
                If TIMEOUTFLAG <> 0 Then
                    frmTimer!tmrTimer.Enabled = False
                    TIMEOUTFLAG = 0
                    TIMECOUNT = 0
                    RcvStr = buf$
                    If disp% = True Then
2004                        PrtInfo buf$
                    End If
                    RCVCOUNT = RCVCOUNT + Len(buf$)
                    Exit Function
                End If
                DoEvents
            Loop
            frmTimer!tmrTimer.Enabled = False
            TIMEOUTFLAG = 0
            TIMECOUNT = 0
        End If
    Loop
    Exit Function

RcvStr_Err:
    ErrorProc "RcvStr", Err, Erl, Error
    FileAdd USERDIR + "LOG.TXT", "RcvStr>Error " + Format$(Err) + ":" + Format$(Erl)
    Exit Function
End Function
