VERSION 2.00
Begin Form frmInfo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'å≈íË(2èdê¸)
   Caption         =   "SPG-NEWS"
   ControlBox      =   0   'False
   Height          =   1485
   Left            =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   6180
   Top             =   4605
   Width           =   6300
   Begin Label lblInfo 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0C0C0&
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   5835
   End
End
Option Explicit

Sub Form_Load ()
    top = (screen.Height - Height) / 2
    left = (screen.Width - Width) / 2
End Sub

