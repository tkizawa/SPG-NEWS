VERSION 2.00
Begin Form frmReMark 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ˆø—p•„‚ÌŽw’è"
   ClientHeight    =   2385
   ClientLeft      =   5805
   ClientTop       =   5520
   ClientWidth     =   4695
   Height          =   2790
   Left            =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4695
   Top             =   5175
   Width           =   4815
   Begin OptionButton optReMark 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ˆø—p•„‚È‚µ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
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
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin OptionButton optReMark 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option1"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2235
   End
   Begin OptionButton optReMark 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option1"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   780
      Width           =   2235
   End
   Begin OptionButton optReMark 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option1"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2235
   End
End
Option Explicit

Sub Command1_Click ()
    Dim i%

    For i% = 0 To 3
        If optReMark(i%).Value = True Then
            RE_MARK = optReMark(i%).Caption
            If i% = 3 Then
                RE_MARK = ""
            End If
            Exit For
        End If
    Next i%
    Unload Me
End Sub

Sub Form_Load ()
    top = (screen.Height - Height) / 2
    left = (screen.Width - Width) / 2
End Sub

