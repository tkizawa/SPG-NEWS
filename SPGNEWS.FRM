VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SPG-NEWS for Windows ListMenu"
   ClientHeight    =   7935
   ClientLeft      =   1305
   ClientTop       =   2310
   ClientWidth     =   12795
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "FixedSys"
   FontSize        =   13.5
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   8625
   Icon            =   SPGNEWS.FRX:0000
   Left            =   1245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   12795
   Top             =   1680
   Width           =   12915
   Begin PictureBox picSelect2 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   12765
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   975
      Width           =   12795
      Begin CommandButton cmdMenuSet 
         Caption         =   "���J�X�^���ݒ�"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   5160
         TabIndex        =   26
         Top             =   0
         Width           =   2055
      End
      Begin ComboBox cboList 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   390
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   60
         Width           =   2715
      End
      Begin Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "���X�g�\�����e"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   23
         Top             =   120
         Width           =   1815
      End
   End
   Begin PictureBox picStatBar 
      Align           =   2  'Align Bottom
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   12795
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7500
      Width           =   12795
      Begin SSPanel pnl3DStatBar1 
         BevelOuter      =   1  'Inset
         BorderWidth     =   2
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   225
         TabIndex        =   21
         Top             =   45
         Width           =   4335
      End
      Begin SSPanel pnl3DTime 
         BevelOuter      =   1  'Inset
         BorderWidth     =   2
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4665
         TabIndex        =   20
         Top             =   60
         Width           =   1770
      End
   End
   Begin PictureBox picToolBox 
      Align           =   1  'Align Top
      BackColor       =   &H8000000F&
      Height          =   435
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   12765
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   12795
      Begin CommandButton cmdTerminal 
         Cancel          =   -1  'True
         Caption         =   "�^�[�~�i��"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "�l�r �S�V�b�N"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   7860
         TabIndex        =   25
         Top             =   0
         Width           =   975
      End
      Begin SSCommand cmdAuto 
         AutoSize        =   2  'Adjust Button Size To Picture
         Caption         =   "����߲ۯ�"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "�l�r �S�V�b�N"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   6900
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin SSCommand cmdFWD 
         AutoSize        =   2  'Adjust Button Size To Picture
         Caption         =   "FWD�]��"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "�l�r �S�V�b�N"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   5940
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin SSCommand cmdRead 
         AutoSize        =   2  'Adjust Button Size To Picture
         Caption         =   "�ǂ�"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "�l�r �S�V�b�N"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   4800
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin SSCommand cmdDelete 
         AutoSize        =   2  'Adjust Button Size To Picture
         Caption         =   "�폜"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "�l�r �S�V�b�N"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   3840
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin SSCommand cmdReply 
         AutoSize        =   2  'Adjust Button Size To Picture
         Caption         =   "ҰٕԐM"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "�l�r �S�V�b�N"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   2880
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin SSCommand cmdFollowup 
         AutoSize        =   2  'Adjust Button Size To Picture
         Caption         =   "�f���ԐM"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "�l�r �S�V�b�N"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   1920
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin SSCommand cmdMail 
         AutoSize        =   2  'Adjust Button Size To Picture
         Caption         =   "Ұُ���"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "�l�r �S�V�b�N"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   960
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
      Begin SSCommand cmdBull 
         AutoSize        =   2  'Adjust Button Size To Picture
         Caption         =   "�f������"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "�l�r �S�V�b�N"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   975
      End
      Begin Image imgBull_up 
         Height          =   360
         Left            =   5280
         Picture         =   SPGNEWS.FRX:0302
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin CommonDialog CMDlgFont 
      Left            =   9420
      Top             =   2340
   End
   Begin ListBox lstMenu 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   2190
      Left            =   2040
      MultiSelect     =   2  'Extended
      TabIndex        =   13
      Top             =   3120
      Width           =   5895
   End
   Begin PictureBox picListMenu 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   12795
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   435
      Width           =   12795
      Begin CommandButton cmdScan 
         Caption         =   "�ŐV���X�g"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   6420
         TabIndex        =   11
         Top             =   0
         Width           =   1335
      End
      Begin ComboBox cboCall 
         BackColor       =   &H00FFFFFF&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   390
         Left            =   4920
         Sorted          =   -1  'True
         TabIndex        =   10
         Text            =   "cboCall"
         Top             =   60
         Width           =   1335
      End
      Begin ComboBox cboAtbbs 
         BackColor       =   &H00FFFFFF&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   390
         Left            =   2820
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "cboAtbbs"
         Top             =   60
         Width           =   1335
      End
      Begin ComboBox cboToField 
         BackColor       =   &H00FFFFFF&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   660
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "cboToField"
         Top             =   60
         Width           =   1335
      End
      Begin Label lblCall 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Call:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   120
         Width           =   615
      End
      Begin Label lblAtbbs 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "@BBS"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   270
         Left            =   2160
         TabIndex        =   16
         Top             =   120
         Width           =   480
      End
      Begin Label lblToField 
         BackColor       =   &H00C0C0C0&
         Caption         =   "To:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   270
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   330
         WordWrap        =   -1  'True
      End
   End
   Begin Label Label3 
      BackColor       =   &H00008000&
      Caption         =   "Label3"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   1680
      TabIndex        =   24
      Top             =   2220
      Width           =   7155
   End
   Begin Menu mnuFile 
      Caption         =   "�t�@�C��(&F)"
      Begin Menu mnuBulletin 
         Caption         =   "�f����������"
      End
      Begin Menu mnuMail 
         Caption         =   "���[����������"
      End
      Begin Menu mnuReplyBulletin 
         Caption         =   "�f���ŕԐM"
      End
      Begin Menu mnuReplyMail 
         Caption         =   "���[���ԐM"
      End
      Begin Menu mnuNo0 
         Caption         =   "-"
      End
      Begin Menu mnuKill 
         Caption         =   "�폜"
         Begin Menu mnuKillMark 
            Caption         =   "�}�[�N�������b�Z�[�W���폜"
         End
         Begin Menu mnuTimeKill 
            Caption         =   "n���ȑO�̃��b�Z�[�W���폜"
         End
      End
      Begin Menu mnuRevival 
         Caption         =   "����"
      End
      Begin Menu mnuReWrite 
         Caption         =   "�ĕҏW"
      End
      Begin Menu mnuSpoolOut 
         Caption         =   "Spool�̓��e�������o��"
      End
      Begin Menu mnuNo1 
         Caption         =   "-"
      End
      Begin Menu mnuReduce 
         Caption         =   "���ݔ�����ɂ���"
      End
      Begin Menu mnuNo2 
         Caption         =   "-"
      End
      Begin Menu mnuExit 
         Caption         =   "�I��"
      End
   End
   Begin Menu mnuRead 
      Caption         =   "�ǂ�(&R)"
   End
   Begin Menu mnuCom 
      Caption         =   "�ʐM(&C)"
      Begin Menu mnuComSR 
         Caption         =   "�e�v�c�]��"
      End
      Begin Menu mnuAutoPilot 
         Caption         =   "�����^�]"
      End
      Begin Menu mnuTerm 
         Caption         =   "�^�[�~�i��"
      End
   End
   Begin Menu mnuSetUp 
      Caption         =   "�ݒ�(&S)"
      Begin Menu mnuConfig 
         Caption         =   "���ݒ�..."
      End
      Begin Menu mnuFWDConfig 
         Caption         =   "�ڑ��ݒ�..."
      End
      Begin Menu mnuAutoPilotConfig 
         Caption         =   "�I�[�g�p�C���b�g�ݒ�..."
      End
      Begin Menu mnuDownLoad 
         Caption         =   "�_�E�����[�h�ݒ�..."
      End
      Begin Menu mnuDownset 
         Caption         =   "�_�E�����[�h�\��..."
      End
      Begin Menu mnuTimer 
         Caption         =   "�^�C�}�[�ݒ�..."
      End
      Begin Menu mnuSignaturSet 
         Caption         =   "�����ݒ�..."
      End
      Begin Menu mnuServerSet 
         Caption         =   "�A�[�J�C�u�ݒ�..."
      End
      Begin Menu mnuSetFont 
         Caption         =   "�t�H���g"
         Begin Menu mnuFontText 
            Caption         =   "�e�L�X�g�r���A�[..."
         End
         Begin Menu mnuFontList 
            Caption         =   "���X�g..."
         End
      End
      Begin Menu mnuOpt0 
         Caption         =   "�I�v�V����"
         Begin Menu mnuDJLSet 
            Caption         =   "DJL-MB�ިڸ�ؐݒ�..."
         End
         Begin Menu mnuSpoolDir 
            Caption         =   "�����A�b�v���[�h�f�B���N�g���ݒ�..."
         End
         Begin Menu mnuHTMLEditor 
            Caption         =   "�g�s�l�k�G�f�B�^�̐ݒ�..."
         End
         Begin Menu mnuWEBBrowser 
            Caption         =   "�v�d�a�u���E�U�̐ݒ�..."
         End
         Begin Menu mnuUserDir 
            Caption         =   "�^�p�f�[�^�f�B���N�g���̐ݒ�..."
         End
         Begin Menu mnuMonitor 
            Caption         =   "�ʐM���j�^�̕\��"
            Checked         =   -1  'True
         End
      End
   End
   Begin Menu mnuWin 
      Caption         =   "�E�B���h�D(&W)"
      Begin Menu mnuListOn 
         Caption         =   "���X�g"
         Shortcut        =   ^L
      End
      Begin Menu mnuTextOn 
         Caption         =   "�e�L�X�g�r���A�["
         Shortcut        =   ^T
      End
      Begin Menu mnuEditOn 
         Caption         =   "�e�L�X�g�G�f�B�^"
         Shortcut        =   ^E
      End
   End
   Begin Menu mnuHelp0 
      Caption         =   "�w���v(&H)"
      Begin Menu mnuHelp 
         Caption         =   "�ڎ�"
         Shortcut        =   {F1}
      End
      Begin Menu mnuAbout 
         Caption         =   "�o�[�W�������"
         Shortcut        =   ^A
      End
   End
End
' ���X�g���j���[
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Sub cboAtbbs_Click ()
    cboChange
    FileListMake "!K"
End Sub

Sub cboAtbbs_KeyPress (KeyAscii As Integer)
    If KeyAscii = 13 Then
	cboChange
	FileListMake "!K"
    End If
End Sub

Sub cboCall_Click ()
    cboChange
    FileListMake "!K"
End Sub

Sub cboCall_KeyPress (KeyAscii As Integer)
    If KeyAscii = 13 Then
	cboChange
	FileListMake "!K"
    End If
End Sub

Sub cboChange ()
    If cboToField.Text <> "*" Or cboAtbbs.Text <> "*" Or cboCall.Text <> "*" Then
	frmMain!cboList.ListIndex = 0
    End If
End Sub

Sub cboList_Click ()
    Dim i As Integer
    Dim s$
    
    s$ = cboList.Text
    If s$ = "���[�U�[�w��" Then
	FileListMake "!K"
    End If
    If s$ = "�S�ĕ\��" Then
	frmMain!cboToField.Text = "*"
	frmMain!cboAtbbs.Text = "*"
	frmMain!cboCall.Text = "*"
	FileListMake "!K"
    End If
    If s$ = MYCALL + "�����[��" Then
	frmMain!cboToField.Text = MYCALL
	frmMain!cboAtbbs.Text = "*"
	frmMain!cboCall.Text = "*"
	FileListMake "!K"
    End If
    If s$ = "���M�ۗ�" Then
	frmMain!cboToField.Text = "*"
	frmMain!cboAtbbs.Text = "*"
	frmMain!cboCall.Text = MYCALL
	FileListMake "H"
    End If
    If s$ = "���ݔ�" Then
	frmMain!cboToField.Text = "*"
	frmMain!cboAtbbs.Text = "*"
	frmMain!cboCall.Text = "*"
	FileListMake "K"
    End If

    If FileExist(USERDIR + "LISTMENU.TXT") = True Then
	i = 0
	Do
	    If s$ = MENUSET$(i, 0) Then
		frmMain!cboToField.Text = MENUSET$(i, 1)
		frmMain!cboAtbbs.Text = MENUSET$(i, 2)
		frmMain!cboCall.Text = MENUSET$(i, 3)
		FileListMake "!K"
		Exit Sub
	    End If
	    i = i + 1
	    If MENUSET$(i, 0) = "" Then
		Exit Sub
	    End If
	Loop
    End If
End Sub

Sub cboToField_Click ()
    cboChange
    FileListMake "!K"
End Sub

Sub cboToField_KeyPress (KeyAscii As Integer)
    If KeyAscii = 13 Then
	cboChange
	FileListMake "!K"
    End If
End Sub

Sub cmdAuto_Click ()
    Call mnuAutoPilot_Click
End Sub

Sub cmdBull_Click ()
    Call mnuBulletin_Click
End Sub

Sub cmdDelete_Click ()
    Call mnuKillMark_Click
End Sub

Sub cmdFollowup_Click ()
    Call mnuReplyBulletin_Click
End Sub

Sub cmdFWD_Click ()
    Call mnuComSR_Click
End Sub

Sub cmdMail_Click ()
    Call mnuMail_Click
End Sub

Sub cmdMenuSet_Click ()
    frmMenuSet.Show 1
End Sub

Sub cmdRead_Click ()
    Call mnuRead_Click
End Sub

Sub cmdReply_Click ()
    Call mnuReplyMail_Click
End Sub

Sub cmdScan_Click ()
    Select Case MODE
    Case "AUTO", "FWD"
	FileListMake "!K"
    Case "DJLMB"
	FileDJLMsgLoad (False)
	FileListMake "!K"
    End Select
End Sub

Sub cmdTerminal_Click ()
    frmTerminal.Show
End Sub

Sub Form_Activate ()
    WINDOW_MODE = 0
End Sub

Sub Form_Load ()
    Dim FileNum As Integer
    Dim i, j As Integer
    Dim fl As Integer
    Dim MaxNum As Integer
    Dim s As String
    Dim n$, t$, a$, f$, w%
    Dim r, LimitWidth

    On Error GoTo ERRORPROC
    
    frmMain.Width = frmMain.ScaleWidth
    w% = GetPrivateProfileInt("Window", "WIndowState", 0, INIFILE)
    frmMain.WindowState = w%

    LimitWidth = 1024
    r = SendMessage(lstMenu.hWnd, &H415, LimitWidth, 0)
    
    If L_MODE = 1 Then
	MENUSET1
    End If

    If MODE = "FWD" Then
	mnuComSR.Enabled = True
	cmdFWD.Enabled = True
	cmdAuto.Enabled = False
	mnuAutoPilot.Enabled = False
	mnuAutoPilotConfig.Enabled = False
    ElseIf MODE = "AUTO" Then
	mnuComSR.Enabled = False
	cmdFWD.Enabled = False
	cmdAuto.Enabled = True
	mnuAutoPilot.Enabled = True
	mnuAutoPilotConfig.Enabled = True
    ElseIf MODE = "DJLMB" Then
	mnuComSR.Enabled = False
	cmdFWD.Enabled = False
	cmdAuto.Enabled = False
	mnuAutoPilot.Enabled = False
	mnuAutoPilotConfig.Enabled = False
'        mnuKill.Enabled = False
'        mnuRevival.Enabled = False
	mnuReWrite.Enabled = False
'        mnuReduce.Enabled = False
'        cmdDelete.Enabled = False
	mnuServerSet.Enabled = False
    End If

    frmMain!cboList.AddItem "���[�U�[�w��"
    frmMain!cboList.AddItem "�S�ĕ\��"
    frmMain!cboList.AddItem MYCALL + "�����[��"
    If FileExist(USERDIR + "LISTMENU.TXT") = True Then
	i = 0
	FileNum = FreeFile
	Open USERDIR + "LISTMENU.TXT" For Input As #FileNum
	Do While EOF(FileNum) = False
	    Input #FileNum, n$, t$, a$, f$
	    MENUSET$(i, 0) = n$
	    MENUSET$(i, 1) = t$
	    MENUSET$(i, 2) = a$
	    MENUSET$(i, 3) = f$
	    i = i + 1
	    frmMain!cboList.AddItem n$
	Loop
	Close #FileNum
    End If
    frmMain!cboList.AddItem "���M�ۗ�"
    frmMain!cboList.AddItem "���ݔ�"

    frmMain!cboToField.Clear
    frmMain!cboAtbbs.Clear
    frmMain!cboCall.Clear
    frmMain!cboToField.AddItem "*"
    frmMain!cboAtbbs.AddItem "*"
    frmMain!cboCall.AddItem "*"

    FileNum = FreeFile
1000 Open USERDIR + "TOFIELD.DAT" For Input As FileNum
    Do While Not EOF(FileNum)
	Line Input #FileNum, s
	If s <> "" Then
	    frmMain!cboToField.AddItem s
	End If
    Loop
    Close FileNum

    FileNum = FreeFile
1010 Open USERDIR + "ATBBS.DAT" For Input As FileNum
    Do While Not EOF(FileNum)
	Line Input #FileNum, s
	If s <> "" Then
	    frmMain!cboAtbbs.AddItem s
	End If
    Loop
    Close FileNum
    
    FileNum = FreeFile
1020 Open USERDIR + "CALL.DAT" For Input As FileNum
    Do While Not EOF(FileNum)
	Line Input #FileNum, s
	If s <> "" Then
	    frmMain!cboCall.AddItem s
	End If
    Loop
    Close FileNum
    
    frmMain!cboToField.ListIndex = 0
    frmMain!cboAtbbs.ListIndex = 0
    frmMain!cboCall.ListIndex = 0
    frmMain!cboList.ListIndex = 1

    frmMain!lstMenu.FontName = Trim$(FONTNAMELISTMENU)
    frmMain!lstMenu.FontSize = FONTSIZELISTMENU
'    frmMain!lstListTitle.FontName = Trim$(FONTNAMELISTMENU)
'    frmMain!lstListTitle.FontSize = FONTSIZELISTMENU
    frmMain!Label3.FontName = Trim$(FONTNAMELISTMENU)
    frmMain!Label3.FontSize = FONTSIZELISTMENU
    
    ' ���j���[�ʐM���j�^�̃`�F�b�N
    If MONITORON = True Then
	mnuMonitor.Checked = True
	frmMonitor.Show
    Else
	mnuMonitor.Checked = False
    End If
    
    Exit Sub

ERRORPROC:
    Select Case Erl
	Case 1000
	    f$ = USERDIR + "TOFIELD.DAT"
	Case 1010
	    f$ = USERDIR + "ATBBS.DAT"
	Case 1020
	    f$ = USERDIR + "CALL.DAT"
	Case Else
	    Error Err
    End Select
    Open f$ For Output As FileNum
    If Erl = 1020 Then
	Print #FileNum, MYCALL
    End If
    Close FileNum
    Resume

End Sub

Sub Form_Resize ()
    Dim y1, y2
    
    ' ���C���E�B���h�D���A�C�R��������Ă���Ƃ��̓T�C�Y�̕ύX���s��Ȃ�
    If frmMain.WindowState = 1 Then
	WINDOW_STATE = 1
	frmTerminal.WindowState = 1
	frmTextView.WindowState = 1
	frmEdit.WindowState = 1
	frmMonitor.WindowState = 1
	frmPanel.WindowState = 1
	Exit Sub
    End If
    
    If frmMain.WindowState = 0 Then
	WINDOW_STATE = 0
	frmTerminal.WindowState = 0
	frmTextView.WindowState = 0
	frmEdit.WindowState = 0
	frmMonitor.WindowState = 0
	frmPanel.WindowState = 0
    End If

    Label3.Top = picListMenu.Height + picToolBox.Height + picSelect2.Height
    Label3.Left = 0
    Label3.Width = frmMain.ScaleWidth

    lstMenu.Top = Label3.Height + picListMenu.Height + picToolBox.Height + picSelect2.Height
    lstMenu.Left = 0
    lstMenu.Width = frmMain.ScaleWidth
'    lstMenu.Height = frmMain.ScaleHeight - (picListMenu.Height + picToolBox.Height + picSelect2.Height + picStatBar.Height + Label3.Height)
    y1 = picStatBar.Top
    y2 = lstMenu.Top
    lstMenu.Height = y1 - y2
End Sub

Sub Form_Unload (Cancel As Integer)
    Unload frmTextView
    Unload frmEdit
    Unload frmPanel
    Unload frmMonitor
    a% = WritePrivateProfileString("Window", "WindowState", Format$(WINDOW_STATE), INIFILE)
    SPGNEWS_End
End Sub

Sub lstMenu_Click ()
    Dim ListNum&
    Dim i As Integer

    ListNum& = Val(Left$(frmMain!lstMenu.Text, 5))
    ' ���b�Z�[�W�ʂ��ԍ��ݒ�
    i = DataGetNum(ListNum&)
    If Left$(MESSAGE(i).stat, 1) = "P" Then
	cmdFollowup.Enabled = False
	mnuReplyBulletin.Enabled = False
    Else
	cmdFollowup.Enabled = True
	mnuReplyBulletin.Enabled = True
    End If
End Sub

Sub lstMenu_DblClick ()
    Dim ListNum&
    Dim num As Integer

    ListNum = Val(Left$(frmMain!lstMenu.Text, 5))
    ' ���b�Z�[�W�ʂ��ԍ��ݒ�
    num = DataGetNum(ListNum&)
    If MESSAGE(num).Call = MYCALL And Mid$(MESSAGE(num).stat, 2, 1) <> "F" Then
	' ���M�O��������ĕҏW����
	If Left$(MESSAGE(num).stat, 1) = "B" Then
	    MESSAGE(num).stat = "B$H"
	Else
	    MESSAGE(num).stat = "P H"
	End If
	ReWrite
    Else
	ReadMail
    End If
End Sub

Sub lstMenu_KeyPress (KeyAscii As Integer)
    If KeyAscii = 13 Then
	ReadMail
    End If
End Sub

Sub MENUSET1 ()
    mnuFile.Caption = "File(&F)"
    mnuBulletin.Caption = "Bulletin Write"
    mnuMail.Caption = "Mail Write"
    mnuReplyBulletin.Caption = "Follow Up"
    mnuReplyMail.Caption = "Reply"
    mnuKill.Caption = "Delete"
    mnuKillMark.Caption = "Delete(Mark)"
    mnuTimeKill.Caption = "Delete(before day of 'n')"
    mnuRevival.Caption = "Revival"
    mnuReWrite.Caption = "Collection"
    mnuReduce.Caption = "Emptied Trush Box"
    mnuExit.Caption = "Exit"
    '
    mnuRead.Caption = "Read(&R)"
    mnuCom.Caption = "Communication(&C)"
    mnuComSR.Caption = "Forward Start"
    mnuAutoPilot.Caption = "Auto Pilot Start"
    '
    mnuSetup.Caption = "Setup(&C)"
    mnuConfig.Caption = "Environment..."
    mnuFWDConfig.Caption = "Connect Method..."
    mnuAutoPilotConfig.Caption = "Auto Pilot..."
    mnuDownLoad.Caption = "Down Load..."
    mnuDownset.Caption = "Down Load Reserve..."
    mnuTimer.Caption = "Timer..."
    mnuSetFont.Caption = "Font..."
    mnuFontText.Caption = "Text Viewer..."
    mnuFontList.Caption = "List..."
    mnuSignaturSet.Caption = "Signature..."
    mnuServerSet.Caption = "Archive..."
    '
    mnuWin.Caption = "Window"
    mnuListOn.Caption = "Display List"
    mnuTextOn.Caption = "Display Text Viewer"
    mnuEditOn.Caption = "Display Text Editor"
    '
    mnuHelp0.Caption = "Help(&H)"
    mnuHelp.Caption = "Contents"
    mnuAbout.Caption = "About..."
    '
    cmdScan.Caption = "Search"
    Label1.Caption = "Content of List"
End Sub

Sub mnuAbout_Click ()
    OpenAbout
End Sub

Sub mnuAutoPilot_Click ()
    TIMERCOM = True
    cmdAuto.Enabled = False
    mnuAutoPilot.Enabled = False
    AutoPilotStart
    TIMERCOM = False
    cmdAuto.Enabled = True
    mnuAutoPilot.Enabled = True
End Sub

Sub mnuAutoPilotConfig_Click ()
    frmAutoConfig.Show 1
End Sub

' �f����������
'
Sub mnuBulletin_Click ()
    ToolBarDisable
    WriteBulletin
End Sub

Sub mnuComSR_Click ()
    TIMERCOM = True
    cmdFWD.Enabled = False
    mnuComSR.Enabled = False
    Forward
    FILECHANGEFLAG = True
    TIMERCOM = False
    cmdFWD.Enabled = True
    mnuComSR.Enabled = True
End Sub

Sub mnuConfig_Click ()
    SetUpConfig (True)
End Sub

Sub mnuDJLSet_Click ()
    frmDJLSetUP.Show 1
End Sub

Sub mnuDownLoad_Click ()
    frmDownSet.Show 1
End Sub

Sub mnuDownset_Click ()
    frmListNext.Show 1
End Sub

Sub mnuEditOn_Click ()
    frmEdit.Show
End Sub

Sub mnuExit_Click ()
    Unload frmMain
End Sub

Sub mnuFontList_Click ()
    On Error Resume Next
    ' [�t�H���g�̎w��] �_�C�A���O �{�b�N�X�̃t���O
    Const CF_BOTH = &H3&
'    Const CF_EFFECTS = &H100&
    ' �t���O���Z�b�g���܂�.
    CMDlgFont.Flags = CF_BOTH Or CF_EFFECTS
    ' �_�C�A���O �{�b�N�X�̏����l��ݒ肵�܂�.
    CMDlgFont.FontSize = frmMain!lstMenu.FontSize
    CMDlgFont.FontName = frmMain!lstMenu.FontName
    ' [�t�H���g�̎w��] �_�C�A���O �{�b�N�X��\�����܂�.
    CMDlgFont.Action = 4
    ' �I�������I�v�V�����ɏ]����, �e�L�X�g �{�b�N�X��
    ' �e�L�X�g�̃t�H���g��ύX���܂�.
    frmMain!lstMenu.FontName = CMDlgFont.FontName
    frmMain!lstMenu.FontSize = CMDlgFont.FontSize
'    frmMain!lstListTitle.FontName = CMDlgFont.FontName
'    frmMain!lstListTitle.FontSize = CMDlgFont.FontSize
    frmMain!Label3.FontName = CMDlgFont.FontName
    frmMain!Label3.FontSize = CMDlgFont.FontSize
    FONTNAMELISTMENU = CMDlgFont.FontName
    FontNameSizeText = CMDlgFont.FontSize
    FONTNAMELISTMENU = Trim$(FONTNAMELISTMENU)

    Call Form_Resize
'    frmMain!lstMenu.FontBold = CMDlgFont.FontBold
'    frmMain!lstMenu.FontItalic = CMDlgFont.FontItalic
'    frmMain!lstMenu.FontUnderline = CMDlgFont.FontUnderLine
'    frmMain!lstMenu.FontStrikethru = CMDlgFont.FontStrikeThru
'    frmMain!lstMenu.ForeColor = CMDlgFont.Color
    a = WritePrivateProfileString("Font", "ListMenuFontName", frmMain!lstMenu.FontName, INIFILE)
    a = WritePrivateProfileString("Font", "ListMenuFontSize", Format$(frmMain!lstMenu.FontSize), INIFILE)
End Sub

Sub mnuFontText_Click ()
    On Error Resume Next
    ' [�t�H���g�̎w��] �_�C�A���O �{�b�N�X�̃t���O
    Const CF_BOTH = &H3&
'    Const CF_EFFECTS = &H100&
    ' �t���O���Z�b�g���܂�.
    CMDlgFont.Flags = CF_BOTH Or CF_EFFECTS
    ' �_�C�A���O �{�b�N�X�̏����l��ݒ肵�܂�.
    CMDlgFont.FontSize = frmTextView!txtTextView.FontSize
    CMDlgFont.FontName = frmTextView!txtTextView.FontName
    ' [�t�H���g�̎w��] �_�C�A���O �{�b�N�X��\�����܂�.
    CMDlgFont.Action = 4
    ' �I�������I�v�V�����ɏ]����, �e�L�X�g �{�b�N�X��
    ' �e�L�X�g�̃t�H���g��ύX���܂�.
    frmTextView!txtTextView.FontName = CMDlgFont.FontName
    frmTextView!txtTextView.FontSize = CMDlgFont.FontSize
    FONTNAMETEXTVIEW = CMDlgFont.FontName
    FontNameSizeText = CMDlgFont.FontSize
    FONTNAMETEXTVIEW = Trim$(FONTNAMETEXTVIEW)

'    frmTextView!txtTextView.FontBold = CMDlgFont.FontBold
'    frmTextView!txtTextView.FontItalic = CMDlgFont.FontItalic
'    frmTextView!txtTextView.FontUnderline = CMDlgFont.FontUnderLine
'    frmTextView!txtTextView.FontStrikethru = CMDlgFont.FontStrikeThru
'    frmTextView!txtTextView.ForeColor = CMDlgFont.Color
    a = WritePrivateProfileString("Font", "TextViewFontName", frmTextView!txtTextView.FontName, INIFILE)
    a = WritePrivateProfileString("Font", "TextViewFontSize", Format$(frmTextView!txtTextView.FontSize), INIFILE)
End Sub

Sub mnuFWDConfig_Click ()
    frmFWDConfig.Show 1
End Sub

Sub mnuHelp_Click ()
    OpenHelp
End Sub

Sub mnuHTMLEditor_Click ()
    frmHTMLEditorSet.Show 1
End Sub

Sub mnuKillMark_Click ()
    MenuFileKill
    FileListMake "!K"
End Sub

Sub mnuListOn_Click ()
    frmMain.Show
End Sub

' ���[����������
'
Sub mnuMail_Click ()
    ToolBarDisable
    WriteMail
End Sub

Sub mnuMonitor_Click ()
    Dim a%
    
    If mnuMonitor.Checked = True Then
	mnuMonitor.Checked = False
	a% = WritePrivateProfileString("Config", "Monitor", "OFF", INIFILE)
	Unload frmMonitor
    Else
	mnuMonitor.Checked = True
	a% = WritePrivateProfileString("Config", "Monitor", "ON", INIFILE)
	frmMonitor.Show
    End If
End Sub

Sub mnuRead_Click ()
    ReadMail
End Sub

Sub mnuReduce_Click ()
    FileCSVSave
    FileListMake "!K"
    FILECHANGEFLAG = False
End Sub

' �f���ԐM
'
Sub mnuReplyBulletin_Click ()
    ReplyBulletin
End Sub

Sub mnuReplyMail_Click ()
    ReplyMail
End Sub

Sub mnuRevival_Click ()
    MenuFileRevival
End Sub

Sub mnuReWrite_Click ()
    ReWrite
End Sub

Sub mnuServerSet_Click ()
    frmServerSet.Show 1
End Sub

Sub mnuSignaturSet_Click ()
    Dim a As Integer

    a = Shell(EDITOR + " " + USERDIR + "SIGNATUR.TXT", 1)
End Sub

Sub mnuSpoolDir_Click ()
    frmSpoolSetup.Show 1
End Sub

Sub mnuSpoolOut_Click ()
    SpoolAll
End Sub

Sub mnuTerm_Click ()
    frmTerminal.Show
End Sub

Sub mnuTextOn_Click ()
    frmTextView.Show
End Sub

Sub mnuTimeKill_Click ()
    frmTimeKill.Show 1
End Sub

Sub mnuTimer_Click ()
    frmTimerSet.Show 1
End Sub

Sub mnuUserDir_Click ()
    frmUserDir.Show 1
End Sub

Sub mnuWEBBrowser_Click ()
    frmWEBBrowserSet.Show 1
End Sub
