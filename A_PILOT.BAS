' �I�[�g�p�C���b�g�V�X�e���ʐM���W���[��
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA
Option Explicit

'
' �I�[�g�p�C���b�g���C�����[�`��
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

    ' �^�[�~�i���\�t�g�g�p���̓I�[�g�p���b�g�����s���Ȃ�
    If TERMFLAG = True Then
        Exit Sub
    End If
    
    ' �����񌟍��ʒu�ݒ�
10    SetSerchPos

    ' �R�l�N�g
    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g �R�l�N�g��"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>�R�l�N�g�J�n"
    Idol_off
20  If Connect() <> 0 Then
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>�R�l�N�g���s"
        Disconnect
        Idol_on
        Exit Sub
    End If
    ' �r�h�c��M�҂�
    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g �r�h�c��M�҂�"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>�r�h�c��M�҂�"

30  If WaitComm("[", RCVWAITTIME) = DISC$ Then
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>SID��M���s"
        Disconnect
        Idol_on
        Exit Sub
    End If

    ' �v�����v�g�擾
    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g �v�����v�g�҂�"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>�v�����v�g�҂�"

    ' �v�����v�g�̃`�F�b�N�Ɛ���
    p$ = PromptMake()
    If p$ = DISC$ Then
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>�v�����v�g�擾���s"
        Disconnect
        Idol_on
        Exit Sub
    End If
    PROMPT1 = p$

    ' �^�C���A�E�g�ݒ�
    frmTimer!tmrAuto.Interval = 60000   ' �ꕪ��
    frmTimer!tmrAuto.Enabled = True     '�^�C�}�[�X�^�[�g
    
    ' �r�x�r�n�o�R�}���h���s
    If AutoSendSysopCmd() = False Then
        Exit Sub
    End If

    ' ���M�\�胁�b�Z�[�W�̑��M
    If AutoSendMessage() = False Then
        Exit Sub
    End If

    ' �c�����X�g������Ύ擾����
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
        ' ���X�g���擾���ă_�E�����[�h���郁�b�Z�[�W��READNEXT.TXT�ɋL�^����
        If AutoGetList(ln1$, False) = False Then
            Exit Sub
        End If
    End If
    
    ' �\��σ��b�Z�[�W�ǂݎ��
    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g �\��σ��b�Z�[�W�ǂݎ�蒆"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>�\��σ��b�Z�[�W�ǂݎ��"
50  If AutoReadNextMail() = False Then
        Disconnect
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>�\��σ��b�Z�[�W�ǂݎ�莸�s"
        Idol_on
        Exit Sub
    End If
    
' -----------------------------------------------------------------------------
    
    ' ���X�g�R�}���h���M
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

    ' ���X�g���擾���ă_�E�����[�h���郁�b�Z�[�W��READNEXT.TXT�ɋL�^����
    If AutoGetList(ln1$, True) = False Then
        Exit Sub
    End If

    ' ���b�Z�[�W�ǂݍ���
    ' READNEXT.TXT�ɏ]�����b�Z�[�W���_�E�����[�h
    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g ���b�Z�[�W��M��"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>���b�Z�[�W��M"
    If AutoReadNextMail() = False Then
        Disconnect
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>���b�Z�[�W�ǂݎ�莸�s"
        Idol_on
        Exit Sub
    End If
    
    ' �I�[�g�p�C���b�g����I��
    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g ����f"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>�I�[�g�p�C���b�g����I��"
    r$ = SendStr("B", 0, True)
    Disconnect
    Idol_on
    Exit Sub

Auto_Error:
    ErrorProc "AutoPilot", Err, Erl, Error
    PrtInfo "�s���ȃ��X�g���擾���܂����B" + Chr$(13) + Chr$(10) + "���X�g���������������Ɛ�������܂��B" + Chr$(13) + Chr$(10) + "�ʐM�������I�����܂�"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>Error " + Format$(Err) + ":" + Format$(Erl)
    Disconnect
    Idol_on
    Exit Sub
End Sub

'
' ���X�g���擾���ă_�E�����[�h���郁�b�Z�[�W��READNEXT.TXT�ɋL�^����
' ���X�g�擾�Ɏ��s�����ꍇ���X�g���Ƃ��Ă��Ȃ����X�g�ԍ���LISTNEXT.TXT�ɋL�^����
' ln1$ : ���n�߂郊�X�g�̍ŏ��l
' flag : True=����I����ŐV���X�g�ԍ����L�^����
'�@�@�@�@False=�L�^���Ȃ�
'
Function AutoGetList (lnst$, flag As Integer)
    Dim a%
    Dim ln As String * 20
    Dim r$, rb$, rn$, s$
    Dim sel%

    ' �O��R�l�N�g���̍ŐV���X�g�ԍ����擾
'    a% = GetPrivateProfileString("AutoPilot", "ListNum", "1", ln, 10, INIFILE)
'    ln1$ = RNullCut(ln)
    ' ���X�g���t�@�C���ɋL�^����
    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g ���X�g��M�҂�"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>���X�g��M�҂�"
    r$ = RcvStrLine(RCVWAITTIME, True)
    If r$ = DISC$ Then
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>���X�g��M�ҋ@���ɉ���f1" + Chr$(34) + r$ + Chr$(34)
        Disconnect
        Idol_on
        AutoGetList = False
        Exit Function
    End If
    ' ���X�g�{�̂���M����܂ő҂�
    Do While (Val(r$) = 0)
        r$ = RcvStrLine(RCVWAITTIME, True)
        If r$ Like PROMPT1 = True Then
            ' ���X�g�Ȃ������_�E�����[�h���Ȃ�
            frmPanel!pnlRcv.Caption = ""
            AutoGetList = True
            Exit Function
        End If
        If r$ = DISC$ Then
            FileAdd USERDIR + "LOG.TXT", "AutoPilot>���X�g��M�ҋ@���ɉ���f2" + Chr$(34) + r$ + Chr$(34)
            Disconnect
            Idol_on
            AutoGetList = False
            Exit Function
        End If
    Loop

    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g ���X�g��M��"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>���X�g��M��"
    ' ���X�g�̍ŐV�ԍ����L�^
    rn$ = r$
    rb$ = r$
    ' ���X�g�擾
    ' �v�����v�g������܂Ń��X�g���擾����
    Do While (r$ Like PROMPT1 = False)
        If InStr(r$, DISC$) <> 0 Then
            ' ���X�g��M�ҋ@���ɉ���f
            FileAdd USERDIR + "LOG.TXT", "AutoPilot>���X�g��M�ҋ@���ɉ���f3" + Chr$(34) + r$ + Chr$(34)
            Disconnect
            Idol_on
            AutoGetList = False
            FileAdd USERDIR + "LISTNEXT.TXT", lnst$ + "," + Format$(Val(rb$) - 1)
            s$ = Format$(Val(rn$))
           a% = WritePrivateProfileString("AutoPilot", "ListNum", UCase$(s$), INIFILE)
            Exit Function
        End If
        rb$ = r$
        ' �_�E�����[�h����
60      sel% = MessageSelect(r$)
        ' �����Ώۂ�������ǂ�
        If sel% = True Then
            frmPanel!pnlRcv.Caption = "����"
            FileAdd USERDIR + "READNEXT.TXT", r$
        Else
            frmPanel!pnlRcv.Caption = "���Ȃ�"
        End If
        ' �L�^
        If Val(r$) <> 0 Then
            r$ = Left$(r$, InStr(r$, Chr$(13)) - 1)
            FileAdd USERDIR + "LIST.TXT", r$
        End If
        ' ��s��M����܂ő҂�
        r$ = ""
        r$ = RcvStrLine(RCVWAITTIME, True)
        DoEvents
    Loop
    frmPanel!pnlRcv.Caption = ""
    AutoGetList = True
    ' ���X�g����擾�I���ŏI���X�g�ԍ��L�^
    If flag = True Then
        s$ = Format$(Val(rn$))
        a% = WritePrivateProfileString("AutoPilot", "ListNum", UCase$(s$), INIFILE)
    End If
End Function

'
' �I�[�g�p�C���b�g�p�A�C�e�����o
'
Sub AutoListCut (r$, msg As MSGD2)
    Dim i As Integer
    Dim at$, f$, t$, tt$, r1$
    
    Trace "AutoListCut"
    
    r1$ = r$
    If Val(r1$) < 0 Then
        r1$ = Mid$(r1$, 2)
    End If
    ' ���X�g��蕶���񒊏o
    't$ = mid$(r$, AUTOPOSTOFIELD, instr(AUTOPOSTOFIELD, r$, " ") - AUTOPOSTOFIELD)
    ' To Field�̋�؂蕶���̓X�y�[�X��@�̏ꍇ������̂ŗ�����؂蕶���Ƃ��Č�������
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
    ' MSGD2�ݒ�
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
' ���[���ǂݏo��
'
' n$ : <num>
'      ������擪�� - ������ƍ폜����
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
    ' �v�����v�g�擾
    ' Read�R�}���h����
6002    cmd$ = READCMD + " " + Format$(Abs(num))
    ' Read�R�}���h���M
6003    r$ = SendStr(cmd$, 0, True)
    ' �t�@�C���I�[�v��
6004    FileNum = FreeFile
6005    Open USERDIR + "READ.TXT" For Output As FileNum
    ' ���[���ǂݏo��
6006    r$ = RcvStrLine(RCVWAITTIME, True)
    ' �v�����v�g�����o����܂œǂݏo��
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
    ' �������̃��[����������폜����
6019    a = GetPrivateProfileString("AutoPilot", "MailKill", "Y", mk, 5, INIFILE)
6020    mkf$ = Left$(mk, 1)
6021    If Val(n$) < 0 And mkf$ = "Y" Then     ' B185
        ' Kill�R�}���h����
6022        cmd$ = "K " + Format$(Abs(num))
        ' Read�R�}���h���M
6023        r$ = SendStr(cmd$, 0, True)
        ' �v�����v�g��M�҂�
6024        c% = 0
6025        r$ = RcvStrLine(RCVWAITTIME, True)
6026        Do While (r$ Like PROMPT1 = False)
6027            c% = c% + 1
6028            If c% = 10 Then
6029                FileAdd USERDIR + "LOG.TXT", "AutoPilot>���������[����M������f"
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
' READNEXT.TXT�ɂ��ƂÂ��ă��b�Z�[�W��ǂݏo��
' ���b�Z�[�W�ǂݎ�蒆�G���[������������READNEXT.TXT���Đ�����
'
Function AutoReadNextMail ()
    ReDim ReadNum(0) As String
    Dim msg As MSGD2
    Dim FileNum As Integer
    Dim i%, j%, rc%, st%, r$

    On Error GoTo AutoReadNextMail_Err
    
    ' �ǂݍ��ݗ\��̃��b�Z�[�W��ǂݍ���
5000    rc% = 0
5001    FileNum = FreeFile
5002    st% = True
    ' READNEXT.TXT������Γǂݍ��ݗ\�胁�b�Z�[�W��ǂݍ���
5003    If (FileExist(USERDIR + "READNEXT.TXT") = True) Then
5004        Open USERDIR + "READNEXT.TXT" For Input As FileNum
5005        ' ���X�g�ǂݍ���
5006        Do While (EOF(FileNum) = False)
5007            ReDim Preserve ReadNum(rc%) As String
5008            Line Input #FileNum, r$
5009            ReadNum(rc%) = r$
5010            rc% = rc% + 1
5011        Loop
5012        Close FileNum
        ' ���b�Z�[�W�ǂݍ���
5013        For i% = rc% - 1 To 0 Step -1
5014            If Val(ReadNum(i%)) <> 0 Then  ' R 0 �����M�����o�O�̎b��΍�(��Ŗ{�΍􂷂�)
5015                AutoListCut ReadNum(i%) + Chr$(13), msg
5016                frmPanel!pnlToField.Caption = msg.field
5017                frmPanel!pnlAtbbs.Caption = msg.atbbs
5018                frmPanel!pnlCall.Caption = msg.Call
5019                frmPanel!pnlTitle.Caption = msg.Title
                    ' �������̃��[����������s����"-"��ǉ����� 'B186
                    If msg.field = MYCALL Then
                        ReadNum(i%) = "-" + ReadNum(i%)
                    End If
5020                st% = AutoReadMail(ReadNum(i%))
                ' �ǂݍ��ݎ��s�̏ꍇ
5021                If st% = False Then
5022                    Exit For
5023                End If
                ' �t�@�C����������
                ' @BBS���w�肳��Ă��Ȃ�������HomeRBBS���w�肷��
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
        ' ���b�Z�[�W�𐳏�Ƀ_�E�����[�h�ł��Ȃ������ꍇ
        ' ����_�E�����[�h�\��t�@�C�����쐬����
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
' ���b�Z�[�W�̑��M
' �߂�l: -1 DISCONNECT
'         -2 S�R�}���h���͋���
'
Function AutoSendFile (i As Integer) As Integer
    Dim FileNum As Integer
    Dim Filename As String
    Dim a$, r$, send$

    Trace "AutoSendFile"
    
    r$ = RcvStr(RCVWAITTIME, True)
    ' TITLEINPUT���ݒ肳��Ă��Ȃ��ꍇ�̓v�����v�g�̃`�F�b�N�����Ȃ�
    If TITLEINPUT <> "" Then
        If r$ <> TITLEINPUT + Chr$(13) + Chr$(10) Then
            ' �^�C�g�����̓v�����v�g�łȂ�
            AutoSendFile = -2
            Exit Function
        End If
    End If
    If InStr(r$, DISC$) <> 0 Then
        AutoSendFile = -1
        Exit Function
    End If

    ' �^�C�g�����M
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
' ���X�g�R�}���h���M
'
Function AutoSendList (ln1$, lc1$, listcommand As Integer)
    Dim a%
'    Dim ln As String * 20
'    Dim lc As String * 20
'    Dim listcommand As Integer
    Dim sp$, s$, r$

    ' ���X�g�R�}���h���M
    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g ���X�g�R�}���h���M"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>���X�g�R�}���h���M"
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
        FileAdd USERDIR + "LOG.TXT", "AutoPilot>���X�g�R�}���h���M���s"
        Disconnect
        Idol_on
        AutoSendList = False
        Exit Function
    End If
    AutoSendList = True
End Function

'
' ���M�\��̃��b�Z�[�W�𑗐M����
'
Function AutoSendMessage ()
    Dim c%, i%, MaxNum%, st%
    Dim r$, stat$

    ' ���b�Z�[�W��������
    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g ���b�Z�[�W���M��"
    FileAdd USERDIR + "LOG.TXT", "AutoPilot>���b�Z�[�W���M"
    MaxNum% = UBound(MESSAGE)

    For i% = 0 To MaxNum%
        stat$ = Mid$(MESSAGE(i%).stat, 2, 1)
        If (stat$ <> "") And (stat$ <> "K") And (stat$ <> "F") And Mid$(MESSAGE(i%).stat, 3, 1) <> "K" And Mid$(MESSAGE(i%).stat, 3, 1) <> "H" Then
            frmPanel!pnlToField.Caption = MESSAGE(i%).field
            frmPanel!pnlAtbbs.Caption = MESSAGE(i%).atbbs
            frmPanel!pnlCall.Caption = MESSAGE(i%).Call
            frmPanel!pnlTitle.Caption = MESSAGE(i%).Title
            ' ���M�R�}���h���M
100         If (FWDSendCmd(i%, False) <> 0) Then
                FileAdd USERDIR + "LOG.TXT", "AutoPilot>���M�R�}���h���M���ɉ���f" + Chr$(34) + r$ + Chr$(34)
                Disconnect
                Idol_on
                AutoSendMessage = False
                Exit Function
            End If
            ' �t�@�C�����M
            st% = AutoSendFile(i%)
110         If st% = -1 Then
                ' ����f
                FileAdd USERDIR + "LOG.TXT", "AutoPilot>�t�@�C�����M���ɉ���f" + Chr$(34) + r$ + Chr$(34)
                Disconnect
                Idol_on
                AutoSendMessage = False
                Exit Function
            End If
            If st% = -2 Then
                ' �r�R�}���h�����ۂ��ꂽ F�t���O�𗧂Ăđ��M���Ȃ��悤�ɂ���
                Mid$(MESSAGE(i%).stat, 2) = "F"
                FILECHANGEFLAG = True
            End If
            ' �v�����v�g��M�҂�
            c% = 0
            r$ = RcvStrLine(RCVWAITTIME, True)
            Do While (r$ Like PROMPT1 = False)
                c% = c% + 1
                If c% = 10 Then
                    FileAdd USERDIR + "LOG.TXT", "AutoPilot>�t�@�C�����M�I����v�����v�g�҂��ŉ���f"
                    Disconnect
                    Idol_on
                    AutoSendMessage = False
                    Exit Function
                End If
                r$ = RcvStrLine(RCVWAITTIME, True)
            Loop
            ' �X�e�[�^�X�X�V
            Mid$(MESSAGE(i%).stat, 2) = "F"
            FILECHANGEFLAG = True
        End If
    Next i%
    AutoSendMessage = True
End Function

'
' �r�x�r�n�o�R�}���h���s
'
Function AutoSendSysopCmd ()
    Dim a%
    Dim sm As String * 5
    Dim smf$, r$, p$

    a% = GetPrivateProfileString("AutoPilot", "SysopMode", "N", sm, 3, INIFILE)
    smf$ = Left$(sm, 1)
    If smf$ = "Y" Then
        r$ = SendStr("@", 0, True)
        ' �v�����v�g�̃`�F�b�N�ƍĐ���
        p$ = PromptMake()
        If p$ = DISC$ Then
            FileAdd USERDIR + "LOG.TXT", "AutoPilot>�r�x�r�n�o���[�h�ڍs���s"
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
'    ' ���b�Z�[�W�ǂݍ���
'    frmPanel!pnlInfo.Caption = "�I�[�g�p�C���b�g ���b�Z�[�W��M��"
'    FileAdd SYSDIR + "LOG.TXT", "AutoPilot>���b�Z�[�W��M"
'    st% = True
'    For i% = rc% - 1 To 0 Step -1
'70      AutoListCut ReadNum(i%), msg
'        frmPanel!pnlToField.Caption = msg.field
'        frmPanel!pnlAtbbs.Caption = msg.atbbs
'        frmPanel!pnlCall.Caption = msg.Call
'        frmPanel!pnlTitle.Caption = msg.Title
'        ' �������̃��[����������s����"-"��ǉ�����
'        If msg.field = MYCALL Then
'            ReadNum(i%) = "-" + ReadNum(i%)
'        End If
'80      st% = AutoReadMail(ReadNum(i%))
'        If st% = False Then
'            Exit For
'        End If
'        ' �t�@�C����������
'        ' @BBS���w�肳��Ă��Ȃ�������HomeRBBS���w�肷��
'        If msg.atbbs = "" Then
'            msg.atbbs = HOMERBBS
'        End If
'        If (FileWrite2("READ.TXT", msg.stat, msg.field, msg.atbbs, msg.Call, msg.Title, msg.bid, "***") <> 0) Then
'            FileAdd SYSDIR + "LOG.TXT", "AutoPilot>�t�@�C���������ݎ��s" + Chr$(34) + r$ + Chr$(34)
'            Disconnect
'            Idol_on
'            st% = False
'            Exit For
'        End If
'    Next i%
'    ' ���b�Z�[�W�𐳏�Ƀ_�E�����[�h�ł��Ȃ������ꍇ
'    ' ����_�E�����[�h�\��t�@�C�����쐬����
'    If st% = False Then
'        For j% = i% To 0 Step -1
'90          AutoListCut ReadNum(j%), msg
'            If msg.field = MYCALL And Val(ReadNum(j%)) >= 0 Then
'                ReadNum(j%) = "-" + ReadNum(j%)
'            End If
'            FileAdd SYSDIR + "READNEXT.TXT", ReadNum(j%)
'        Next j%
'        FileAdd SYSDIR + "LOG.TXT", "AutoPilot>���b�Z�[�W�_�E�����[�h���ɉ���f"
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
' �I�[�g�p�C���b�g�p�����f�[�^�쐬
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
'  ���X�g�̃_�E�����[�h�I��
'
Function MessageSelect (s$) As Integer
    Dim t$, tt$, titl$, at$, f$
    Dim i%, imax%
    Dim FileNum As Integer
'    Dim yn$, tfl$, atbbs$, f1$, Title$
    ReDim yn$(100), tfl$(100), atbbs$(100), f1$(100), Title$(100)
    
    MessageSelect = False

    ' ���X�g�ȊO�̕�����??
    If Val(s$) = 0 Then
        MessageSelect = False
        Exit Function
    End If

    ' ���X�g��蕶���񒊏o
    ' To Field�̋�؂蕶���̓X�y�[�X��@�̏ꍇ������̂ŗ�����؂蕶���Ƃ��Č�������
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

    ' �������̃��[���͓ǂݍ���
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
' ���o����v�����v�g�̐���
' �v�����v�g�̃`�F�b�N�Ɛ������s��
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
            MsgBox "�q�a�a�r�̃v�����v�g�� > �݂̂ł��B���b�Z�[�W���̈��p���Ȃǂ��v�����v�g�Ƃ��Č�F������\��������܂��B�I�[�g�p�C���b�g��~��蓮�łq�a�a�r�ɐڑ����v�����v�g��ύX���邱�Ƃ��������߂��܂��B"
        End If
        If InStr(r$, ">" + Chr$(13) + Chr$(10)) <> 0 Then
        ' �v�����v�g����
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

' ��s��M
' ��s��M����܂őҋ@����
'
' t:�^�C���A�E�g����
' dsp_flag:���j�^�\���t���O(True=�\������)
'
Function RcvStrLine (t As Integer, dsp_flag As Integer) As String
    Dim s$, r$

'    Trace "RcvStrLine"
    
    ' ��s��M����܂ő҂�
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
' �_�E�����[�h�J�����ݒ�
'
Sub SetSerchPos ()
    Trace "SetSerchPos"
    AUTOPOSTOFIELD = GetPrivateProfileInt("AutoPilot", "ToFieldColum", 17, INIFILE)
    AUTOPOSCALL = GetPrivateProfileInt("AutoPilot", "CallColum", 31, INIFILE)
    AUTOPOSTITLE = GetPrivateProfileInt("AutoPilot", "TitleColum", 48, INIFILE)
    AUTOPOSATBBS = GetPrivateProfileInt("AutoPilot", "AtbbsColum", 24, INIFILE)
End Sub
