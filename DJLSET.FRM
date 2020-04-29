VERSION 2.00
Begin Form frmDJLSetUp 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "DJL-MB√ﬁ®⁄∏ƒÿê›íË"
   ClientHeight    =   3090
   ClientLeft      =   4215
   ClientTop       =   7260
   ClientWidth     =   6810
   Height          =   3495
   Left            =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6810
   Top             =   6915
   Width           =   6930
   Begin TextBox txtDJLSPOOL 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Text            =   "C:\DJLSPOOL"
      Top             =   960
      Width           =   2115
   End
   Begin TextBox txtDJLDir 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Text            =   "C:\BBS"
      Top             =   360
      Width           =   2115
   End
   Begin CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5460
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin CommandButton cmdCancel 
      Caption         =   "∑¨›æŸ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5460
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "é©ìÆ±ØÃﬂ€∞ƒﬁ√ﬁ®⁄∏ƒÿ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   5
      Top             =   1020
      Width           =   2775
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MAIL.MBÇÃÇ†ÇÈ√ﬁ®⁄∏ƒÿ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
End
Option Explicit

Sub cmdCancel_Click ()
    Unload Me
End Sub

Sub cmdOK_Click ()
    Dim a%
    
    a% = WritePrivateProfileString("Option", "djldir", UCase$(txtDJLDir.Text), INIFILE)
    a% = WritePrivateProfileString("Option", "djlspool", UCase$(txtDJLSPOOL.Text), INIFILE)
    
    DJLMAILDIR = UCase$(txtDJLDir.Text)
    DJLSPOOLDIR = UCase$(txtDJLSPOOL.Text)
    
    Unload Me
End Sub

Sub Form_Load ()
    Dim a%
    Dim mk As String * 10
    Dim mkf$

    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2

    txtDJLDir.Text = DJLMAILDIR
    txtDJLSPOOL.Text = DJLSPOOLDIR
'    a% = GetPrivateProfileString("AutoPilot", "MailKill", "Y", mk, 3, INIFILE)
End Sub

