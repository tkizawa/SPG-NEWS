VERSION 2.00
Begin Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "ÇrÇoÇfÅ|ÇmÇdÇvÇrÇ…Ç¬Ç¢Çƒ"
   ClientHeight    =   2865
   ClientLeft      =   3120
   ClientTop       =   5535
   ClientWidth     =   7230
   ControlBox      =   0   'False
   Height          =   3270
   Icon            =   ABOUT.FRX:0000
   Left            =   3060
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2961.777
   ScaleMode       =   0  'User
   ScaleWidth      =   6572.729
   Top             =   5190
   Visible         =   0   'False
   Width           =   7350
   Begin CommandButton cmdAbout 
      Caption         =   "OK"
      Default         =   -1  'True
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   3180
      TabIndex        =   0
      Top             =   2160
      Width           =   915
   End
   Begin Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TACK.VBX Copyright(c) 1994 by T.Kawada Allrights Reserved (ComputerFan 1995.2) "
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   6915
   End
   Begin Label lblAbout11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SPG-NEWS for Windows V4.0"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   15.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   795
      Left            =   1380
      TabIndex        =   2
      Top             =   240
      Width           =   4275
   End
   Begin Label lblAbout2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copyright (c) 1991-1997 by JO1SPG / T.KIZAWA / SPG SYSTEMS Allrights Reserved"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   1320
      Width           =   6975
   End
End
' SPG-NEWS íòçÏå†ï\é¶ÉtÉHÅ[ÉÄ
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Sub cmdAbout_Click ()
    Unload frmAbout
End Sub

Sub Form_Load ()
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2
    
    If L_MODE = 1 Then
        Caption = "About SPG-NEWS"
    End If
    lblAbout11.Caption = "SPG-NEWS for Windows V4.5" + Chr$(13) + Chr$(10) + VERSION
End Sub

