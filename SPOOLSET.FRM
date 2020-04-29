VERSION 2.00
Begin Form frmSpoolSetup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "自動アップロードディレクトリ設定"
   ClientHeight    =   3120
   ClientLeft      =   2985
   ClientTop       =   5835
   ClientWidth     =   6810
   FontBold        =   0   'False
   FontItalic      =   0   'False
   FontName        =   "System"
   FontSize        =   13.5
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   3525
   Left            =   2925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6810
   Top             =   5490
   Width           =   6930
   Begin SSCheck chkSpool 
      Caption         =   "自分自身の書き込みとして送信する"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4815
   End
   Begin TextBox txtSpoolDir 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   555
      Left            =   2220
      TabIndex        =   3
      Top             =   420
      Width           =   2835
   End
   Begin CommandButton Command2 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5460
      TabIndex        =   1
      Top             =   900
      Width           =   1215
   End
   Begin CommandButton Command1 
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
      Top             =   300
      Width           =   1215
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "自動アップロードディレクトリ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   1935
   End
End
Option Explicit

Sub Command1_Click ()
    Dim a%, sd$
    
    sd$ = UCase$(txtSpoolDir.Text)
    
    If sd$ + "\" = SYSDIR Or sd$ = MSGDIR Or sd$ = USERDIR Then
        MsgBox "自動アップロードディレクトリがSPG-NEWSのカレントディレクトリまたはメッセージ格納ディレクトリと重複しています。"
        Exit Sub
    End If
    
    a% = WritePrivateProfileString("Option", "spooldir", UCase$(txtSpoolDir.Text), INIFILE)
    SPOOLDIR = UCase$(txtSpoolDir.Text)
    a% = WritePrivateProfileString("Option", "spoolup", Format$(chkSpool.Value), INIFILE)
    SPOOLUP = chkSpool.Value
    Unload Me
End Sub

Sub Command2_Click ()
    Unload Me
End Sub

Sub Form_Load ()
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2
    txtSpoolDir.Text = SPOOLDIR
    If SPOOLUP = True Then
        chkSpool.Value = True
    Else
        chkSpool.Value = False
    End If
End Sub

