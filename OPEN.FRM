VERSION 2.00
Begin Form frmOpen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4185
   ClientLeft      =   2115
   ClientTop       =   2490
   ClientWidth     =   7860
   ControlBox      =   0   'False
   Enabled         =   0   'False
   ForeColor       =   &H00FF0000&
   Height          =   4590
   Left            =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7860
   Top             =   2145
   Visible         =   0   'False
   Width           =   7980
   Begin PictureBox picCount 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      ScaleHeight     =   225
      ScaleWidth      =   6945
      TabIndex        =   0
      Top             =   3120
      Width           =   6975
   End
   Begin PictureBox picCount1 
      BackColor       =   &H00FF0000&
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      ScaleHeight     =   225
      ScaleWidth      =   6945
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   6975
   End
   Begin SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   2  'Raised
      BevelOuter      =   0  'None
      BevelWidth      =   5
      BorderWidth     =   0
      ForeColor       =   &H00FFFFFF&
      Height          =   3795
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7740
      Begin PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   1320
         Picture         =   OPEN.FRX:0000
         ScaleHeight     =   1455
         ScaleWidth      =   5175
         TabIndex        =   3
         Top             =   240
         Width           =   5175
      End
      Begin Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Power to Make our Dreams come true."
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "Times New Roman"
         FontSize        =   18
         FontStrikethru  =   0   'False
         FontUnderline   =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   900
         TabIndex        =   5
         Top             =   2160
         Width           =   6075
      End
      Begin Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copyright (c) 1991-1997 by JO1SPG/T.KIZAWA Allrights Reserved."
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   7515
      End
      Begin Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Version 4.5"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   24
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   555
         Left            =   2640
         TabIndex        =   4
         Top             =   1620
         Width           =   2595
      End
   End
End
' ÇrÇoÇfÅ|ÇmÇdÇvÇrãÜã…ÇÃÉIÅ[ÉvÉjÉìÉOâÊñ 
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Sub Form_Load ()
    Height = Panel3D1.Height
    Width = Panel3D1.Width
    top = (screen.Height - Height) / 2
    left = (screen.Width - Width) / 2
End Sub

