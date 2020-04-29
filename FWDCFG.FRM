VERSION 2.00
Begin Form frmFWDConfig 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "接続設定"
   ClientHeight    =   5340
   ClientLeft      =   2895
   ClientTop       =   1935
   ClientWidth     =   8475
   FontBold        =   0   'False
   FontItalic      =   0   'False
   FontName        =   "MS Sans Serif"
   FontSize        =   8.25
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   5745
   Left            =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8475
   Top             =   1590
   Width           =   8595
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Index           =   7
      Left            =   240
      TabIndex        =   26
      Top             =   4620
      Width           =   6915
      Begin TextBox txtConnect1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Index           =   7
         Left            =   1920
         TabIndex        =   23
         Top             =   180
         Width           =   4875
      End
      Begin SSOption Option3D1 
         Caption         =   "Send"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   21
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
      Begin SSOption Option3D1 
         Caption         =   "Rcv"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   15
         Left            =   1140
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
   End
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Index           =   6
      Left            =   240
      TabIndex        =   27
      Top             =   3960
      Width           =   6915
      Begin SSOption Option3D1 
         Caption         =   "Rcv"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   13
         Left            =   1140
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
      Begin SSOption Option3D1 
         Caption         =   "Send"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
      Begin TextBox txtConnect1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Index           =   6
         Left            =   1920
         TabIndex        =   20
         Top             =   180
         Width           =   4875
      End
   End
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Index           =   5
      Left            =   240
      TabIndex        =   28
      Top             =   3300
      Width           =   6915
      Begin TextBox txtConnect1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Index           =   5
         Left            =   1920
         TabIndex        =   17
         Top             =   180
         Width           =   4875
      End
      Begin SSOption Option3D1 
         Caption         =   "Send"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   15
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
      Begin SSOption Option3D1 
         Caption         =   "Rcv"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   11
         Left            =   1140
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
   End
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Index           =   4
      Left            =   240
      TabIndex        =   29
      Top             =   2640
      Width           =   6915
      Begin SSOption Option3D1 
         Caption         =   "Rcv"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   9
         Left            =   1140
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
      Begin SSOption Option3D1 
         Caption         =   "Send"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
      Begin TextBox txtConnect1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Index           =   4
         Left            =   1920
         TabIndex        =   14
         Top             =   180
         Width           =   4875
      End
   End
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   30
      Top             =   1980
      Width           =   6915
      Begin TextBox txtConnect1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Index           =   3
         Left            =   1920
         TabIndex        =   11
         Top             =   180
         Width           =   4875
      End
      Begin SSOption Option3D1 
         Caption         =   "Send"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
      Begin SSOption Option3D1 
         Caption         =   "Rcv"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   7
         Left            =   1140
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
   End
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   31
      Top             =   1320
      Width           =   6915
      Begin TextBox txtConnect1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Index           =   2
         Left            =   1920
         TabIndex        =   8
         Top             =   180
         Width           =   4875
      End
      Begin SSOption Option3D1 
         Caption         =   "Send"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
      Begin SSOption Option3D1 
         Caption         =   "Rcv"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   5
         Left            =   1140
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
   End
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   32
      Top             =   660
      Width           =   6915
      Begin SSOption Option3D1 
         Caption         =   "Rcv"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   3
         Left            =   1140
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
      Begin SSOption Option3D1 
         Caption         =   "Send"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
      Begin TextBox txtConnect1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Index           =   1
         Left            =   1920
         TabIndex        =   5
         Top             =   180
         Width           =   4875
      End
   End
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   33
      Top             =   0
      Width           =   6915
      Begin TextBox txtConnect1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   180
         Width           =   4875
      End
      Begin SSOption Option3D1 
         Caption         =   "Send"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
      Begin SSOption Option3D1 
         Caption         =   "Rcv"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1140
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
   End
   Begin CommandButton cmdFWDCAN 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   25
      Top             =   720
      Width           =   975
   End
   Begin CommandButton cmdFWDOK 
      Caption         =   "OK"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   24
      Top             =   120
      Width           =   975
   End
End
' 接続設定ダイアログ
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Sub cmdFWDCAN_Click ()
    Unload frmFWDConfig
End Sub

Sub cmdFWDOK_Click ()
    Dim i, FileNum As Integer
    Dim cmd$

    If FILEEXist(USERDIR + "CONNECT.SPG") = True Then
        Kill USERDIR + "CONNECT.SPG"
    End If
    FileNum = FreeFile
    Open USERDIR + "CONNECT.SPG" For Output As #FileNum
    For i = 0 To 7
        If txtConnect1(i) = "" Then Exit For
        If Option3D1(i * 2).Value = True Then
            cmd$ = "S"
        Else
            cmd$ = "R"
        End If
        Print #FileNum, cmd$
        Print #FileNum, txtConnect1(i)
    Next i
    Close #FileNum
    Unload frmFWDConfig
End Sub

Sub Form_Load ()
    Dim i, FileNum As Integer
    Dim a$, b$

    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2
    
    If FILEEXist(USERDIR + "CONNECT.SPG") = True Then
        FileNum = FreeFile
        Open USERDIR + "CONNECT.SPG" For Input As #FileNum
        i = 0
        Do While (EOF(FileNum) = False)
            Line Input #FileNum, a$
            Line Input #FileNum, b$
            If a$ = "R" Then
                Option3D1(i * 2 + 1).Value = True
            Else
                Option3D1(i * 2).Value = True
            End If
            txtConnect1(i) = b$
            i = i + 1
        Loop
        Close #FileNum
    End If

    If L_MODE = 1 Then
        frmFWDConfig.Caption = "Connect Setting"
        cmdFWDCAN.Caption = "Cancel"
    End If
End Sub

