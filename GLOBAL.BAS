' �O���[�o���֐���`
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Global MSGFILENAME As String        '�@���b�Z�[�W�t�@�C���l�[��
Global MYCALL As String             '�@�R�[���T�C��
Global HOMERBBS As String           '  �e�v�c����̃R�[���T�C��
Global LINK As String               '  �����N���@
Global BAUD As String               '  �{�[���[�g
Global MSGDIR As String             '  ���b�Z�[�W�i�[�f�B���N�g��
Global EDITOR As String             '  �N���G�f�B�^��
Global PORT As Integer              '�@�ʐM�|�[�g�i�P����w��j
Global MODE As String               ' ���샂�[�h(ex.FWD,AUTO)
Global MAILHEADER As String         ' �d�q���[���w�b�_�[�̕t��
Global COM_MONITOR As String        ' �ʐM���j�^�̕\��
Global MYNAME As String             ' �g�p�Җ�
Global USERDIR As String            ' ���[�U�[�ݒ�t�@�C���i�[�f�B���N�g��
Global XFLOWFLAG As Integer         ' �\�t�g�E�F�A�t���[����t���O
'
Global SYSDIR As String              '  �J�����g�f�B���N�g��
Global VERSION As String             '  �o�[�W����
Global WRITEMODE As String           '  �������݃��[�h�@�o�@�����@�a
Global WRITE_TFL As String
Global WRITE_ATBBS As String
'
Global TOFIELDCOL As Integer
Global ATBBSCOL As Integer
Global CALLCOL As Integer
Global TITLECOL As Integer
'
Global MESSAGERECORDNUM As Integer
Global DEBUGFLAG As Integer
'
Global MESSAGE() As MSGD2
'
Global FONTNAMELISTMENU As String * 40
Global FONTSIZELISTMENU As Single
Global FONTNAMETEXTVIEW As String * 40
Global FONTSIZETEXTVIEW As Single
'
Global DOWNSETMODE As String
Global SERVERSETMODE As String
'
Global TIMEZONE As String

' ���[�h�R�}���h
Global READCMD As String

'�^�C�}�[�A�E�g�t���O
Global TIMEOUTFLAG As Integer
Global TIMECOUNT As Integer

' �Œ蕶����
Global Const DISC$ = "*** DISCONNECTED"

' �t�@�C���Ǘ���񃍁[�h�t���O
Global FILELOADFLAG As Integer      ' FILE.IDX��������荞�ݍς݃t���O
Global FILECHANGEFLAG As Integer    ' FILE.IDX���e�������t���O

' �t�@�C���_�E�����[�h���̃_�C�A���O����
Global FILEEX As Integer

' �I�[�g�p�C���b�g�ݒ�
Global AUTOTOFIELD() As String
Global AUTOATBBS() As String
Global AUTOCALL() As String

'�I�[�g�p�C���b�g�����񌟍��ʒu
Global AUTOPOSTOFIELD As Integer
Global AUTOPOSATBBS As Integer
Global AUTOPOSCALL As Integer
Global AUTOPOSTITLE As Integer

' �ʐM�N�����Ԃ̐ݒ�
Global TIMERENABLE As String
Global TIMESTARTHOUR As String * 2
Global TIMEENDHOUR As String * 2
Global TIMEMINUTES As String * 2
Global TIMERCOM As Integer
Global TIMERSET() As String * 5
Global TIMERMODE As Integer
Global TIMERINTERVAL As Integer

' ���ݒ�t�@�C���t�@�C����
Global INIFILE As String

' �T�[�o�[�ݒ�
Global SERVERENABLE As String           '�T�[�o�[��ON/OFF
Global SERVERCYCLE As String * 2        ' �T�[�o�[�N���T�C�N��
Global SERVERFILE() As String * 40      ' �T�[�o�[�t�@�C�����z��
Global SERVERTOFIELD() As String * 10   ' �T�[�o�[ToField
Global SERVERATBBS() As String * 10     ' �T�[�o�[@BBS
Global SERVERCALL() As String * 10      ' �T�[�o�[From
Global SERVERMSGNUM As Integer          ' �T�[�o�[�L�^���b�Z�[�W�ԍ�
Global SVNAME$(), SVTFL$(), SVATBBS$(), SVCALL$()
                                        ' �T�[�o�[�ݒ荀��

' �^�C���A�E�g����
Global RCVWAITTIME As Integer
Global SENDWAITTIME As Integer

' ���M�f�B���C
Global SENDDELAY As Long

' �f�[�^�x�[�X�t�@�C���ԍ�
Global DATABASEFILENUM As Integer

' Windows��
Global R_MODE As Integer
Global WINDOW_MODE As Integer

' ���ꃂ�[�h
Global L_MODE As Integer

' �ʐM�|�[�g�I�[�v���t���O
Global COMFLAG As Integer

' �h�l�d���[�h
Global IMEVALUE As Integer

' DJL-MB Multi DIr
Global DJLMAILDIR As String
Global DJLMULTIDIR As Integer
Global DJLSPOOLDIR As String

' �_�E�����[�h�t���O
Global DOWNLOAD As Integer

' �v�����v�g
Global PROMPT1 As String

' �^�[�~�i���\�t�g�g�p�t���O
Global TERMFLAG As Integer

' �^�C�g�����̓v�����v�g
Global TITLEINPUT As String

' �X�v�[���f�B���N�g��
Global SPOOLDIR As String
Global SPOOLUP As Integer

' �ʐM���j�^�̕\��
Global MONITORON As Integer

' ���p��
Global RE_MARK As String

' ���X�g�\���J�X�^���ݒ�
Global MENUSET$(100, 3)
Global MENUSETINDEX As Integer

' �I�[�g�p�C���b�g�����l
Global TIMELIMIT As Long
Global RCVLIMIT As Long
Global RCVCOUNT As Long

' ����N�����̃t���O
Global FIRST As Integer

' LaPuta���[�h�t���O
Global LAPUTAMODE As Integer
Global NEWSGROUP As String

' �O���e�L�X�g�G�f�B�^�쐬�t�@�C����
Global EXTFILENAME As String

' HTML�G�f�B�^
Global HTMLEDITOR As String

' �E�F�b�u�u���E�U
Global WEBBROWSER As String

' HTML�t���O
Global HTMLFLAG As Integer

' Window�̏��
Global WINDOW_STATE As Integer
