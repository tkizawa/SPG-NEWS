VERSION 2.00
Begin Form frmMonitor 
   Caption         =   "通信モニタ"
   ClientHeight    =   5940
   ClientLeft      =   2235
   ClientTop       =   5145
   ClientWidth     =   9600
   Height          =   6345
   Left            =   2175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9600
   Top             =   4800
   Width           =   9720
   Begin TextBox txtMonitor 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   5415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   9615
   End
End
Option Explicit

Sub Form_Load ()
    ReadWinPos frmMonitor, "Monitor"
    frmMain!mnuMonitor.Checked = True
End Sub

Sub Form_Resize ()
    On Error Resume Next ' 最小時のプロパティエラー発生防止

    txtMonitor.Top = 0
    txtMonitor.Left = 0
    txtMonitor.Height = ScaleHeight
    txtMonitor.Width = ScaleWidth
End Sub

Sub Form_Unload (Cancel As Integer)
    If frmMonitor.WindowState <> 1 Then
        WrtWinPos frmMonitor, "Monitor"
    End If
    frmMain!mnuMonitor.Checked = False
End Sub

