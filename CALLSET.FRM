VERSION 2.00
Begin Form frmCallSet 
   BackColor       =   &H00C0C0C0&
   Caption         =   "å¬êlèÓïÒê›íË"
   Height          =   4065
   Left            =   1020
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   8085
   Top             =   1290
   Width           =   8205
   Begin CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   0
      Top             =   2820
      Width           =   1215
   End
End
Option Explicit

Sub cmdCancel_Click (Index As Integer)
    Unload frmCallSet
End Sub

Sub cmdOK_Click (Index As Integer)
    Dim a As Integer

    a = WritePrivateProfileString("Config", "MyCall", MYCALL + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "HomeRBBS", HOMERBBS + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "MyName", MYNAME + Chr$(0), INIFILE)
End Sub

Sub Form_Load ()
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2
End Sub

