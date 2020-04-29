VERSION 2.00
Begin Form frmPanel 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPG-NEWS動作パネル"
   ClientHeight    =   4590
   ClientLeft      =   1380
   ClientTop       =   4035
   ClientWidth     =   9480
   ControlBox      =   0   'False
   Height          =   4995
   Icon            =   PANEL.FRX:0000
   Left            =   1320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9480
   Top             =   3690
   Width           =   9600
   Begin SSPanel pnlRcv 
      Alignment       =   1  'Left Justify - MIDDLE
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1560
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin SSPanel pnlTitle 
      Alignment       =   1  'Left Justify - MIDDLE
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1560
      TabIndex        =   11
      Top             =   2700
      Width           =   6075
   End
   Begin SSPanel pnlCall 
      Alignment       =   1  'Left Justify - MIDDLE
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   3855
   End
   Begin SSPanel pnlAtbbs 
      Alignment       =   1  'Left Justify - MIDDLE
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1560
      TabIndex        =   9
      Top             =   1620
      Width           =   3855
   End
   Begin SSPanel pnlToField 
      Alignment       =   1  'Left Justify - MIDDLE
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1560
      TabIndex        =   8
      Top             =   1080
      Width           =   3855
   End
   Begin SSPanel Panel3D6 
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "受信"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin SSPanel Panel3D5 
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "Title"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   2700
      Width           =   1335
   End
   Begin SSPanel Panel3D4 
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "From"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin SSPanel Panel3D3 
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "@BBS"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   1620
      Width           =   1335
   End
   Begin SSPanel Panel3D2 
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "To Field"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin SSCommand cmdComStop 
      AutoSize        =   2  'Adjust Button Size To Picture
      Caption         =   "通信中断"
      Enabled         =   0   'False
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   3900
      Width           =   2415
   End
   Begin SSPanel pnlInfo 
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   7815
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0C0C0&
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "動作状態"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Option Explicit

Sub cmdComStop_Click ()
    Dim X%
    
    If MODE = "FWD" Then
        frmMain!cmdFWD.Enabled = True
        frmMain!mnuComSR.Enabled = True
    Else
        frmMain!cmdAuto.Enabled = True
        frmMain!mnuAutoPilot.Enabled = True
    End If
    Idol_on
    Disconnect
    X% = SerialClose()
End Sub

Sub Form_Load ()
    Dim i As Integer

    i = GetPrivateProfileInt("Window", "Panel_Mode", 0, INIFILE)
    frmPanel.WindowState = i
    If i = 0 Then
        ReadWinPos frmPanel, "Panel"
        width = 9600
        height = 4995
    End If

    If L_MODE = 1 Then
        Caption = "SPG-NEWS Operation Panel"
        Panel3D1.Caption = "Status"
        Panel3D6.Caption = "Receive"
        cmdComStop.Caption = "Abort"
    End If
End Sub

Sub Form_Unload (Cancel As Integer)
    Dim i, m As Integer
    m = frmPanel.WindowState
    i = WritePrivateProfileString("Window", "Panel_Mode", Format$(m), INIFILE)
    '
    WrtWinPos frmPanel, "Panel"
End Sub

