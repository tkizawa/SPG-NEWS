VERSION 2.00
Begin Form frmAutoOptionSet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "�I�[�g�p�C���b�g�ݒ�i�I�v�V�����j"
   ClientHeight    =   3120
   ClientLeft      =   2910
   ClientTop       =   2820
   ClientWidth     =   6825
   Height          =   3525
   Left            =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6825
   Top             =   2475
   Width           =   6945
   Begin TextBox txtTimeLimit 
      Alignment       =   1  'Right Justify
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin SSCheck chkSysopMode 
      Caption         =   "�r�x�r�n�o���[�h���g��"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   555
      Left            =   300
      TabIndex        =   3
      Top             =   900
      Width           =   3435
   End
   Begin SSCheck chkMailKill 
      Caption         =   "���[���ǂݎ���폜����"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   300
      TabIndex        =   2
      Top             =   240
      Width           =   3315
   End
   Begin CommandButton cmdCancel 
      Caption         =   "��ݾ�"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5460
      TabIndex        =   1
      Top             =   780
      Width           =   1215
   End
   Begin CommandButton cmdOK 
      Caption         =   "OK"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5460
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
   Begin Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "�^�C���A�E�g�ݒ肵�Ȃ��ꍇ��-1����͂��ĉ�����"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "�l�r �S�V�b�N"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   2400
      Width           =   4500
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "��"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   240
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "�^�C���A�E�g�ݒ�"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1920
   End
End
Option Explicit

Sub cmdCancel_Click ()
    Unload Me
End Sub

Sub cmdOK_Click ()
    Dim a%
    
    If chkMailKill.Value = True Then
        a% = WritePrivateProfileString("AutoPilot", "MailKill", "YES", INIFILE)
    Else
        a% = WritePrivateProfileString("AutoPilot", "MailKill", "NO", INIFILE)
    End If
    
    If chkSysopMode.Value = True Then
        a% = WritePrivateProfileString("AutoPilot", "SysopMode", "YES", INIFILE)
    Else
        a% = WritePrivateProfileString("AutoPilot", "SysopMode", "NO", INIFILE)
    End If
    a% = WritePrivateProfileString("AutoPilot", "TimeLimit", txtTimeLimit.Text, INIFILE)
    Unload Me
End Sub

Sub Form_Load ()
    Dim a%
    Dim mk As String * 10
    Dim mkf$

    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2

    ' ���[���ǂݎ��チ�b�Z�[�W���폜���邩�H
    a% = GetPrivateProfileString("AutoPilot", "MailKill", "Y", mk, 5, INIFILE)
    mkf$ = Left$(mk, 1)
    If mkf$ = "Y" Then
        chkMailKill.Value = True
    Else
        chkMailKill.Value = False
    End If

    ' �r�x�r�n�o���[�h
    a% = GetPrivateProfileString("AutoPilot", "SysopMode", "N", mk, 5, INIFILE)
    mkf$ = Left$(mk, 1)
    If mkf$ = "Y" Then
        chkSysopMode.Value = True
    Else
        chkSysopMode.Value = False
    End If

    ' �^�C���A�E�g����
    txtTimeLimit.Text = Format$(TIMELIMIT)
End Sub
