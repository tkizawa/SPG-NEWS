VERSION 2.00
Begin Form frmCount 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SPG-NEWS メールデータ読み出し中"
   ClientHeight    =   600
   ClientLeft      =   2490
   ClientTop       =   2955
   ClientWidth     =   5280
   ControlBox      =   0   'False
   Height          =   1005
   Left            =   2430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   5280
   Top             =   2610
   Width           =   5400
   Begin PictureBox picCount 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   240
      ScaleHeight     =   285
      ScaleWidth      =   4725
      TabIndex        =   0
      Top             =   120
      Width           =   4755
   End
End
' メールデータカウンタダイアログ
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA
Option Explicit

Sub Form_Load ()
    top = (screen.Height - Height) / 2
    left = (screen.Width - Width) / 2
End Sub

