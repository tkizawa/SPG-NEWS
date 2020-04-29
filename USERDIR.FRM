VERSION 2.00
Begin Form frmUserDir 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ユーザー設定ファイルディレクトリの設定"
   ClientHeight    =   3105
   ClientLeft      =   3315
   ClientTop       =   4635
   ClientWidth     =   6765
   Height          =   3510
   Left            =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6765
   Top             =   4290
   Width           =   6885
   Begin CommandButton Command2 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   3
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
      Left            =   5280
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin TextBox txtUserDir 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ディレクトリは予め作成しておいて下さい。"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4935
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ディレクトリ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End

Sub Command1_Click ()
    a% = WritePrivateProfileString("Config", "UserDir", UCase$(txtUserDir.Text), INIFILE)
    MsgBox "この設定を有効にするにはＳＰＧ－ＮＥＷＳを再起動して下さい。", MB_ICONINFORMATION
    Unload Me
End Sub

Sub Command2_Click ()
    Unload Me
End Sub

Sub Form_Load ()
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2
    txtUserDir.Text = USERDIR
End Sub

