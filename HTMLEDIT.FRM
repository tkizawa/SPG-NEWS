VERSION 2.00
Begin Form frmHTMLEditorSet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�g�s�l�k�G�f�B�^�ݒ�"
   ClientHeight    =   3105
   ClientLeft      =   1335
   ClientTop       =   6810
   ClientWidth     =   6780
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "System"
   FontSize        =   13.5
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   3510
   Left            =   1275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6780
   Top             =   6465
   Width           =   6900
   Begin CommonDialog CMDialog1 
      Left            =   3360
      Top             =   1920
   End
   Begin CommandButton Command3 
      Caption         =   "�Q��..."
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin TextBox txtHTMLEditor 
      BackColor       =   &H00FFFFFF&
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Text            =   "C:\MSOFFICE\WINWORD\WINWORD.EXE"
      Top             =   840
      Width           =   4815
   End
   Begin CommandButton Command2 
      Caption         =   "��ݾ�"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin CommandButton Command1 
      Caption         =   "OK"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�g�s�l�k�G�f�B�^"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
End

Sub Command1_Click ()
    Dim a%
    
    a% = WritePrivateProfileString("Option", "HTMLEditor", UCase$(txtHTMLEditor.Text), INIFILE)
    HTMLEDITOR = UCase$(txtHTMLEditor.Text)
    Unload Me
    ChDrive Left$(SYSDIR, 2)
    ChDir Left$(SYSDIR, Len(SYSDIR) - 1)
End Sub

Sub Command2_Click ()
    Unload Me
    ChDrive Left$(SYSDIR, 2)
    ChDir Left$(SYSDIR, Len(SYSDIR) - 1)
End Sub

Sub Command3_Click ()
    Dim FileNum As Integer
    Dim a%, s$, t$

    On Error Resume Next
    
    CMDialog1.DefaultExt = "*"
    CMDialog1.DialogTitle = "�g�s�l�k�G�f�B�^�̐ݒ�"
    CMDialog1.Filter = "EXE File (*.EXE)|*.EXE|"
    ' �ŏ��Ɏg���t�B���^���gAll Files�h�ɐݒ肵�܂�.
    CMDialog1.FilterIndex = 1
    ' [�㏑���֎~] �`�F�b�N �{�b�N�X���`�F�b�N���܂�.
    CMDialog1.Flags = OFN_READONLY
    CMDialog1.CancelError = -1
    ' [�t�@�C�����J��] �_�C�A���O �{�b�N�X��\�����܂�.
    CMDialog1.Action = 1
    ' ���[�U�[�� <OK> �{�^����I�������Ƃ��̏���
    If Err = 0 Then
        s$ = CMDialog1.Filename
        txtHTMLEditor.Text = s$
    ElseIf Err = 32755 Then
        Exit Sub
    End If
End Sub

Sub Form_Load ()
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2

    txtHTMLEditor.Text = HTMLEDITOR
End Sub
