VERSION 2.00
Begin Form frmTextView 
   Caption         =   "SPG-NEWS for Windows TextViewer"
   ClientHeight    =   5040
   ClientLeft      =   750
   ClientTop       =   2985
   ClientWidth     =   9480
   Height          =   5730
   Icon            =   TEXTVIEW.FRX:0000
   Left            =   690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9480
   Top             =   2355
   Width           =   9600
   Begin CommandButton cmdBrowser 
      Caption         =   "WEB�u���E�U�N��"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   0
      Width           =   2055
   End
   Begin CommandButton cmdNextDel 
      Caption         =   "�폜���"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin CommonDialog CMDlgDownLoad 
      Left            =   7140
      Top             =   3660
   End
   Begin CommandButton cmdNext 
      Caption         =   "����ǂ�"
      Default         =   -1  'True
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin CommandButton cmdPre 
      Caption         =   "�O��ǂ�"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin CommandButton cmdDownLoad 
      Caption         =   "�ۑ�"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin CommandButton cmdCopy 
      Caption         =   "�R�s�["
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin PictureBox picTextView 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DragIcon        =   TEXTVIEW.FRX:0302
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00C0C0C0&
      Height          =   1395
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   9480
      TabIndex        =   6
      Top             =   0
      Width           =   9480
      Begin OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�L�̎轸۰�"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   1
         Left            =   6600
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "÷�đI��"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin SSPanel pnlHTML 
         BevelInner      =   2  'Raised
         BevelOuter      =   0  'None
         BevelWidth      =   2
         BorderWidth     =   1
         Caption         =   "HTML"
         Font3D          =   2  'Raised w/heavy shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   510
         Left            =   8280
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin CommandButton cmdSelectAll 
         Caption         =   "�S�ĺ�߰"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   0
         Width           =   1215
      End
      Begin CommandButton cmdMode 
         Enabled         =   0   'False
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   6600
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin CommandButton cmdUpDown 
         Caption         =   "UP"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   330
         Index           =   1
         Left            =   4800
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin CommandButton cmdUpDown 
         Caption         =   "DOWN"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   330
         Index           =   0
         Left            =   4200
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin SSPanel pnlTitle 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00008000&
         BevelOuter      =   1  'Inset
         BorderWidth     =   1
         FloodColor      =   &H00C00000&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   900
         Width           =   7875
      End
      Begin SSPanel pnlFrom 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00008000&
         BevelOuter      =   1  'Inset
         BorderWidth     =   1
         FloodColor      =   &H00C00000&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2460
         TabIndex        =   8
         Top             =   480
         Width           =   1635
      End
      Begin SSPanel pnlToField 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00008000&
         BevelOuter      =   1  'Inset
         BorderWidth     =   1
         FloodColor      =   &H00C00000&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
   End
   Begin TextBox txtTextView 
      DragIcon        =   TEXTVIEW.FRX:0604
      DragMode        =   1  'Automatic
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1680
      Width           =   6735
   End
   Begin Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   7560
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
   End
   Begin Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   7320
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Menu mnuFile 
      Caption         =   "�t�@�C��(&F)"
      Begin Menu mnuSave 
         Caption         =   "�ۑ�"
      End
      Begin Menu mnu0 
         Caption         =   "-"
      End
      Begin Menu mnuClose 
         Caption         =   "����"
      End
   End
   Begin Menu mnuDisplay 
      Caption         =   "�\��(&D)"
      Begin Menu mnuPreRead 
         Caption         =   "�O��ǂ�"
         Shortcut        =   ^P
      End
      Begin Menu mnuNextRead 
         Caption         =   "����ǂ�"
         Shortcut        =   ^N
      End
      Begin Menu mnuNextReadDelete 
         Caption         =   "�����Ď���ǂ�"
         Shortcut        =   ^D
      End
   End
   Begin Menu mnuReply 
      Caption         =   "�ԐM(&R)"
      Begin Menu mnuReplyBulletin 
         Caption         =   "�f���ԐM"
         Shortcut        =   ^F
      End
      Begin Menu mnuReplyMail 
         Caption         =   "���[���ԐM"
         Shortcut        =   ^R
      End
   End
End
' �e�L�X�g�r���A�[
' Copyright (c) 1993-1997 by JO1SPG/T.KIZAWA
Option Explicit

Sub cmdBrowser_Click ()
    Dim a, FileNum, FileNum1, s$
    Dim html_flag

    On Error GoTo cmdBrowser_Click_err

    FileNum = FreeFile
    html_flag = False
    Open USERDIR + "readhtml.txt" For Input As FileNum
    FileNum1 = FreeFile
    Open USERDIR + "read.htm" For Output As FileNum1
    Print #FileNum1, "<HTML>"
    While Not EOF(FileNum)
        Line Input #FileNum, s$
        If UCase$(s$) = "<HTML>" Then
            html_flag = True
        End If
        If UCase$(s$) = "</HTML>" Then
            html_flag = False
        End If
        If html_flag = False Then
            s$ = s$ + "<BR>"
        End If
        Print #FileNum1, s$
    Wend
    Print #FileNum1, "</HTML>"
    Close #FileNum
    Close #FileNum1
    html_flag = False

    a = Shell(WEBBROWSER + " " + USERDIR + "READ.HTM", 1)
    Exit Sub

cmdBrowser_Click_err:
    MsgBox "�w�肵���v�d�a�u���E�U [ " + HTMLEDITOR + " ] ���N���ł��܂���B", 48, "�v�d�a�u���E�U�̋N��"
    Exit Sub
End Sub

Sub cmdCopy_Click ()
    TextCopy
End Sub

Sub cmdDownLoad_Click ()
    TextSave
    ChDrive Left$(SYSDIR, 2)
    ChDir Left$(SYSDIR, Len(SYSDIR) - 1)
End Sub

Sub cmdMode_Click ()
    If cmdMode.Caption = "�e�L�X�g�I�����[�h" Then
        cmdMode.Caption = "�L�̎轸۰�Ӱ��"
        txtTextView.DragMode = 0
    Else
        cmdMode.Caption = "�e�L�X�g�I�����[�h"
        txtTextView.DragMode = 1
    End If
End Sub

Sub cmdNext_Click ()
    NextRead
    If HTMLFLAG = True Then
        pnlHTML.Visible = True
    Else
        pnlHTML.Visible = False
    End If
    If Label2.Caption <> "**" Then
        Label2.Caption = ""
    End If
    Call Form_Activate
End Sub

Sub cmdNext_KeyPress (KeyAscii As Integer)
    cmdNext_Click
End Sub

Sub cmdNextDel_Click ()
    NextReadDelete
    If HTMLFLAG = True Then
        pnlHTML.Visible = True
    Else
        pnlHTML.Visible = False
    End If
    Label2.Caption = ""
    Call Form_Activate
End Sub

Sub cmdPre_Click ()
    PreRead
    If HTMLFLAG = True Then
        pnlHTML.Visible = True
    Else
        pnlHTML.Visible = False
    End If
    Label2.Caption = ""
    Call Form_Activate
End Sub

Sub cmdSelectAll_Click ()
    txtTextView.SelStart = 0
    txtTextView.SelLength = Len(txtTextView.Text)
    TextCopy
End Sub

Sub cmdUpDown_Click (Index As Integer)
    Dim i%
    Dim LParam, hscroll, vscroll As Long

    If Index = 0 Then
        vscroll = -1
    Else
        vscroll = 1
    End If
    If vscroll = -1 Then
        vscroll = &HFFFF
    End If
    LParam = hscroll * 65536 + vscroll
    i% = SendMESSAGE(txtTextView.hWnd, EM_LINESCROLL, 0, LParam)
End Sub

Sub Command1_Click ()
    Dim s As String
    Dim s1 As String * 10000
    Dim a%

    s1 = String$(10000, " ")
    s = txtTextView.Text
    a% = SetNkfOption("s" + Chr$(0) + Chr$(0))
    NkfConvert s1, s + Chr$(0) + Chr$(0)
    txtTextView.Text = s1
End Sub

Sub Form_Activate ()
    Dim s$
    Dim i%
    Dim LParam, hscroll, vscroll As Long

    WINDOW_MODE = 1
    If HTMLFLAG = True Then
        pnlHTML.Visible = True
    Else
        pnlHTML.Visible = False
    End If
    
    s$ = pnlTitle.Caption
    If InStr(s$, "[JIS]") <> 0 Or InStr(s$, "(JIS)") <> 0 Or InStr(s$, "<JIS>") <> 0 Or InStr(s$, "[EUC]") <> 0 Or InStr(s$, "<EUC>") <> 0 Or InStr(s$, "(EUC)") <> 0 Then
        s$ = txtTextView.Text
        txtTextView.Text = StrCnv(s$)
        s$ = pnlTitle.Caption
        pnlTitle.Caption = StrCnv(s$)
    End If

    ' �e�L�X�g�X�N���[��
    If Label2.Caption = "" Then
        vscroll = Val(Label1.Caption) - 1                   ' �X�N���[���s��
        LParam = hscroll * 65536 + vscroll
        i% = SendMESSAGE(txtTextView.hWnd, EM_LINESCROLL, 0, LParam)
        Label2.Caption = "*"
    End If
End Sub

Sub Form_Load ()
    Dim a%

    ReadWinPos frmTextView, "TextView"
    frmTextView!txtTextView.FontName = Trim$(FONTNAMETEXTVIEW)
    frmTextView!txtTextView.FontSize = FONTSIZETEXTVIEW

    a% = GetPrivateProfileInt("Config", "TextViewMode", 0, INIFILE)
    
    If a% = 0 Then
        Option1(0).Value = True
        txtTextView.DragMode = 0
    Else
        Option1(1).Value = True
        txtTextView.DragMode = 1
    End If
End Sub

Sub Form_Resize ()
    On Error Resume Next ' �ŏ����̃v���p�e�B�G���[�����h�~

    txtTextView.Top = picTextView.Height
    txtTextView.Left = 0
    txtTextView.Height = ScaleHeight - picTextView.Height
    txtTextView.Width = ScaleWidth
End Sub

Sub Form_Unload (Cancel As Integer)
    Dim a%, i%
    
'    If cmdMode.Caption = "�e�L�X�g�I�����[�h" Then
    If Option1(0).Value = True Then
        i% = 0
    Else
        i% = -1
    End If
    a% = WritePrivateProfileString("Config", "TextViewMode", Format$(i%) + Chr$(0), INIFILE)
    
    If frmTextView.WindowState <> 1 Then
        WrtWinPos frmTextView, "TextView"
    End If
End Sub

Sub mnuClose_Click ()
    Unload frmTextView
End Sub

Sub mnuNextRead_Click ()
    Call cmdNext_Click
'    NextRead
'    If HTMLFLAG = True Then
'        pnlHTML.Visible = True
'    Else
'        pnlHTML.Visible = False
'    End If
End Sub

Sub mnuNextReadDelete_Click ()
    Call cmdNextDel_Click
'    NextReadDelete
'    If HTMLFLAG = True Then
'        pnlHTML.Visible = True
'    Else
'        pnlHTML.Visible = False
'    End If
End Sub

Sub mnuPreRead_Click ()
    Call cmdPre_Click
'    PreRead
'    If HTMLFLAG = True Then
'        pnlHTML.Visible = True
'    Else
'        pnlHTML.Visible = False
'    End If
End Sub

Sub mnuReplyBulletin_Click ()
    ReplyBulletin
End Sub

Sub mnuReplyMail_Click ()
    ReplyMail
End Sub

Sub mnuSave_Click ()
    TextSave
    ChDrive Left$(SYSDIR, 2)
    ChDir Left$(SYSDIR, Len(SYSDIR) - 1)
End Sub

Sub NextRead ()
    Dim i As Integer
    Dim idx As Integer
    Dim ListNum&
    Dim lst$

    idx = frmMain!lstMenu.ListIndex
    If idx = 0 Then
        Beep
        MsgBox "����ȍ~�̃��b�Z�[�W�͂���܂���", MB_OK Or MB_ICONEXCLAMATION, "SPG-NEWS"
        Label2.Caption = "**"
        Exit Sub
    End If
    
    ' ���̃��b�Z�[�W�ɐi�߂�
    frmMain!lstMenu.ListIndex = idx - 1
    frmMain!lstMenu.Selected(idx) = False
    frmMain!lstMenu.Selected(idx - 1) = True
    ListNum& = Val(Left$(frmMain!lstMenu.Text, 5))
    FileTextMake ListNum&
End Sub

Sub NextReadDelete ()
    Dim i As Integer
    Dim idx As Integer
    Dim ListNum&
    Dim lst$

    ' �t�@�C������
    MenuFileKill
    
    idx = frmMain!lstMenu.ListIndex
    If idx = 0 Then
        Beep
        MsgBox "����ȍ~�̃��b�Z�[�W�͂���܂���", MB_OK Or MB_ICONEXCLAMATION, "SPG-NEWS"
        Label2.Caption = "**"
        Exit Sub
    End If
    
    ' ���̃��b�Z�[�W�ɐi�߂�
    frmMain!lstMenu.ListIndex = idx - 1
    frmMain!lstMenu.Selected(idx) = False
    frmMain!lstMenu.Selected(idx - 1) = True
    ListNum& = Val(Left$(frmMain!lstMenu.Text, 5))
    FileTextMake ListNum&
End Sub

Sub Option1_Click (Index As Integer)
    If Index = 0 Then
        txtTextView.DragMode = 0
    Else
        txtTextView.DragMode = 1
    End If
End Sub

Sub PreRead ()
    Dim i As Integer
    Dim idx As Integer
    Dim ListNum&
    Dim lst$

    idx = frmMain!lstMenu.ListIndex
    If idx >= frmMain!lstMenu.ListCount - 1 Then
        If L_MODE = 0 Then
            MsgBox "����ȑO�̃��b�Z�[�W�͂���܂���", MB_OK Or MB_ICONEXCLAMATION, "SPG-NEWS"
        Else
            MsgBox "No message before this.", MB_OK Or MB_ICONEXCLAMATION, "SPG-NEWS"
        End If
        Exit Sub
    End If
    frmMain!lstMenu.ListIndex = idx + 1
    frmMain!lstMenu.Selected(idx) = False
    frmMain!lstMenu.Selected(idx + 1) = True
    ListNum& = Val(Left$(frmMain!lstMenu.Text, 5))
    FileTextMake ListNum&
End Sub

Sub TextCopy ()
    '�e�L�X�g����
    ClipBoard.SetText frmTextView!txtTextView.SelText
End Sub

Sub TextSave ()
    Dim FileNum As Integer
    Dim a%, s$, t$

    On Error Resume Next
    
    CMDlgDownLoad.DefaultExt = "*"
    CMDlgDownLoad.DialogTitle = "�t�@�C���̕ۑ�"
    CMDlgDownLoad.Filter = "All File (*.*)|*.*|"
    ' �ŏ��Ɏg���t�B���^���gAll Files�h�ɐݒ肵�܂�.
    CMDlgDownLoad.FilterIndex = 1
    ' [�㏑���֎~] �`�F�b�N �{�b�N�X���`�F�b�N���܂�.
    CMDlgDownLoad.Flags = OFN_READONLY
    CMDlgDownLoad.CancelError = -1
    ' [�t�@�C�����J��] �_�C�A���O �{�b�N�X��\�����܂�.
    CMDlgDownLoad.Action = 1
    ' ���[�U�[�� <OK> �{�^����I�������Ƃ��̏���
    If Err = 0 Then
        s$ = CMDlgDownLoad.Filename
        a% = FILEEXist(s$)
        FileNum = FreeFile
        ' �t�@�C���Ȃ�
        If a% = False Then
            Open s$ For Output As FileNum
            t$ = frmTextView!txtTextView.Text
            Print #FileNum, t$
            Close FileNum
        End If
        ' �t�@�C������
        If a% = True Then
            frmFILEEXist.Show 1
            Select Case FILEEX
            Case 1
                Kill s$
                Open s$ For Output As FileNum
            Case 2
                Open s$ For Append As FileNum
            Case Else
                Exit Sub
            End Select
            t$ = frmTextView!txtTextView.Text
            Print #FileNum, t$
            Close FileNum
        End If
        Exit Sub
    ElseIf Err = 32755 Then
        Exit Sub
    End If
End Sub

Sub txtTextView_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    Dim d As Integer
    Static yy, oldyy As Single
    
    d = Y - yy
'    Debug.Print Y; ":"; yy; ":"; d
    yy = Y
'    Exit Sub
    
    If d > 30 Then
        Call cmdUpDown_Click(0)
    ElseIf d < -30 Then
        Call cmdUpDown_Click(1)
        yy = Y
    End If
End Sub
