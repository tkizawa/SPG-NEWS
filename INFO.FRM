VERSION 2.00
Begin Form frmMonitor 
   Caption         =   "通信モニタ"
   ClientHeight    =   4815
   ClientLeft      =   1365
   ClientTop       =   3270
   ClientWidth     =   9450
   FontBold        =   0   'False
   FontItalic      =   0   'False
   FontName        =   "標準明朝"
   FontSize        =   12
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   5550
   Icon            =   INFO.FRX:0000
   Left            =   1305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   9450
   Top             =   2595
   Width           =   9570
   Begin TextBox txtMonitor 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "標準明朝"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   4815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   9435
   End
   Begin Menu mnuEnd 
      Caption         =   "閉じる"
   End
End
' 通信モニタウィンドゥ
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Sub Form_Activate ()
    Exit Sub

    Select Case WINDOW_MODE
    Case 0
        frmMain.Show
    Case 1
        frmTextView.Show
    Case 2
        frmEdit.Show
    End Select
End Sub

Sub Form_Load ()
    ReadWinPos frmMonitor, "Monitor"
    frmTextView!txtTextView.FontName = Trim$(FontNameTextView)
    frmTextView!txtTextView.FontSize = FontSizeTextView
End Sub

Sub Form_Resize ()
    On Error Resume Next

    txtMonitor.Width = frmMonitor.ScaleWidth
    txtMonitor.Height = frmMonitor.ScaleHeight
End Sub

Sub Form_Unload (Cancel As Integer)
    If frmMonitor.WindowState <> 1 Then
        WrtWinPos frmMonitor, "Monitor"
    End If
    Disconnect
    x% = SerialClose()
    Close
End Sub

Sub mnuEnd_Click ()
    Unload frmMonitor
End Sub

