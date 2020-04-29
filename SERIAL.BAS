'Option Explicit
'----------------------------------------------------------------------------
'    Serial Communications Module for VB
'----------------------------------------------------------------------------
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA
'
' ＲＳ－２３２Ｃ構造体定義
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

' 文字列送信
' send$:送信文字列
' tnc%:"1" TNC,"0" シリアルリンク
' disp%:True 表示する
'       False 表示しない
'
' 送信キューが空になるまで送信しないようにした(H7.2.19)
' CTSタイムアウトの監視をするようにした。(H7.3.26) V3.35.5
'
Function SendStr (send$, tnc%, disp%) As String
    Dim st As COMSTAT
    Dim buf1 As String * 1
    
    ' 通信ポートクローズならば終了する
    If COMFLAG = False Then
        SendStr = DISC$
        FileAdd USERDIR + "LOG.TXT", "強制終了"
        Exit Function
    End If
    
    ' 受信データがあれば受信する
    a% = GetCommError(nCid, st)
    If (st.cbInQue <> 0) Then
        r$ = RcvStr(RCVWAITTIME, True)
        SendStr = r$
    Else
        SendStr = ""
    End If

    ' タイマーセット
    TIMEOUTFLAG = 0
    TIMECOUNT = 0
    frmTimer!tmrTimer.Interval = 60000   ' 一分間

    ' TNCへのコマンド送信の場合は0x03を三回送る
    If tnc% <> 0 Then
        send$ = String$(3, Chr$(3)) + send$
    End If
    
    ' 送信キューが空になるまで待つ
    frmTimer!tmrTimer.Enabled = True     'タイマースタート
    a% = GetCommError(nCid, st)
    ' 三分間送信キューが空にならなかったら送信タイムアウト
    Do While (st.cbOutQue <> 0)
        If (TIMEOUTFLAG <> 0 And TIMECOUNT >= SENDWAITTIME) Or COMFLAG = False Then ' 送信タイムアウト時間
            TIMEOUTFLAG = 0
            TIMECOUNT = 0
            frmTimer!tmrTimer.Enabled = False
            SendStr = DISC$
            PrtInfo "送信タイムアウト"
            Exit Function
        End If
        a% = GetCommError(nCid, st)
        DoEvents
    Loop
    ' 文字列送信
    send$ = send$ + Chr$(13)
    For i% = 1 To Len(send$)
        s$ = Mid$(send$, i%, 1)
        ' タイマーセット
        TIMEOUTFLAG = 0
        TIMECOUNT = 0
        frmTimer!tmrTimer.Interval = 60000  ' 一分間
        ' 送信タイムアウト時間以上CTSがONにならなかったら送信タイムアウト
        Do
            If (TIMEOUTFLAG <> 0 And TIMECOUNT >= SENDWAITTIME) Or COMFLAG = False Then  ' B202
                TIMEOUTFLAG = 0
                TIMECOUNT = 0
                frmTimer!tmrTimer.Enabled = False
                SendStr = DISC$
                PrtInfo "送信タイムアウト(CTS-OFF)" + Chr$(13) + Chr$(10)
                Exit Function
            End If
            DoEvents
            a% = WriteComm(nCid, s$ + Chr$(0) + Chr$(0), 1)
            b% = GetCommError(nCid, st)
'            PrtInfo "Status=" + Format$(b%) + " Buffer=" + Format$(st.cbOutQue) + Chr$(13)
        Loop Until b% = 0 ' b%=32 でCTSタイムアウトエラー
        SendDelayWait (SENDDELAY)
    Next i%
    
    If disp% = True Then
        PrtInfo send$
        FileAdd USERDIR + "LOG.TXT", "SEND>" + send$
    End If
End Function

'
'　通信ポートクローズ
'
'　戻り値：正常終了は０
'　　　　　異常終了は負の値
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
'　通信ポート設定
'　　　オープン後に処理すること
'
'　引き数：ボーレート，データサイズ(bit)，パリティ,ストップビット
'　（例）　　(9600)　　　　　(8)　　　　　　(n)        (1)　　　
'
Function SerialConfig (BAUD%, bits%, Parity$, stopbits%) As Integer
    Dim lpdcb As DCB
    Dim ConfigString As String

    ConfigString = PortName + ":"
    ConfigString = ConfigString + Format$(BAUD%) + ","
    ConfigString = ConfigString + Left$(UCase$(Parity$), 1) + ","
    ConfigString = ConfigString + Format$(bits%, "#") + ",1"
    i% = BuildCommDCB(ConfigString, lpdcb)

    lpdcb.id = Chr$(nCid)                       ' 　内部デバイス識別子
    lpdcb.stopbits = Chr$(stopbits%)
    lpdcb.ctstimeout = 1 ' CTSタイムアウト時間 V3.35.5
    If XFLOWFLAG = 0 Then
        lpdcb.bits1 = Chr$(Asc(lpdcb.bits1) Or fbinary Or foutctsflow) ' 送信CTS監視
        lpdcb.bits2 = Chr$(Asc(lpdcb.bits2) Or frtsflow)    ' 受信時RTS制御 　
        lpdcb.XonChar = Chr$(0)
        lpdcb.XoffChar = Chr$(0)
    Else
        lpdcb.bits1 = Chr$(Asc(lpdcb.bits1) Or fbinary)
        lpdcb.bits2 = Chr$(Asc(lpdcb.bits2) Or foutx Or finx)    ' 受信時X-FLOW制御 　
        lpdcb.XonChar = Chr$(17)
        lpdcb.XoffChar = Chr$(19)
    End If
    lpdcb.XonLim = 256
    lpdcb.XoffLim = 256

    SerialConfig = SetCommState(lpdcb)

End Function

'
'　通信ポートオープン
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
'　通信ポート受信
'
'　引き数：buf$=受信バッファ
'　　　　　max%=読み取りバイト数
'　戻り値：読み取りバイト数
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
'　通信ポート送信
'
'　引き数：t$=送信文字列
'　戻り値：送信したバイト数
'　　　　：負の値＝エラー
'
Function SerialWrite (t$) As Integer

    Dim st As COMSTAT

    status% = GetCommError(nCid, st)
    status% = WriteComm(nCid, t$, Len(t$))
    If status% < 0 Then status% = GetCommError(nCid, st)
    SerialWrite = status%

End Function

'
' 文字列受信待ち
'
' s$ = 受信文字列
' t% = 待ち時間
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
    ' 通信ポートオープンエラー処理
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
    
    ' ポートイネーブル
    COMFLAG = True
    OpenPort = 0
End Function

'
' 回線からの文字列受信
'
' t% = 待ち時間（分）
' disp% = True モニタに表示
'
Function RcvStr (t%, disp%) As String
    Dim st As COMSTAT
    Dim buf1 As String * 1
    Dim buf2 As String * 10
    
    On Error GoTo RcvStr_Err

    ' 通信ポートクローズならば終了する
    If COMFLAG = False Then
        RcvStr = DISC$
        FileAdd USERDIR + "LOG.TXT", "強制終了"
        Exit Function
    End If

    ' 受信容量オーバーなら終了する
'    If RCVCOUNT >= RCVLIMIT And Not (RCVLIMIT = -1) Then
'        RcvStr = DISC$
'        FileAdd SYSDIR + "LOG.TXT", "強制終了"
'        Exit Function
'    End If
    
    buf$ = ""
    buf1$ = ""
    TIMEOUTFLAG = 0
    TIMECOUNT = 0
    frmTimer!tmrTimer.Interval = 60000   ' 一分間
    frmTimer!tmrTimer.Enabled = True     'タイマースタート

    ' 受信待ち
    a% = GetCommError(nCid, st)
    Do While (st.cbInQue = 0)
        If (TIMEOUTFLAG <> 0 And TIMECOUNT >= t%) Or COMFLAG = False Then
            frmTimer!tmrTimer.Enabled = False
            TIMEOUTFLAG = 0
            TIMECOUNT = 0
            RcvStr = DISC$
2000            PrtInfo DISC$
            Exit Function                       ' タイムアウトエラー
        End If
        DoEvents
        a% = GetCommError(nCid, st)
    Loop
    
    ' 受信キューにデータがあった場合受信処理をする
    ' 1秒間以上受信キューが空だったら受信処理を終了
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
            If buf1 = Chr$(13) Then         ' 改行コードで終了
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
            frmTimer!tmrTimer.Interval = 1000   ' 一秒間
            frmTimer!tmrTimer.Enabled = True     'タイマースタート
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

