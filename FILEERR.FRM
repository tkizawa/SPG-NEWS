VERSION 2.00
Begin Form frmFileExist 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SPG-NEWS"
   ClientHeight    =   1365
   ClientLeft      =   2715
   ClientTop       =   2850
   ClientWidth     =   4320
   ControlBox      =   0   'False
   Height          =   1770
   Left            =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4320
   Top             =   2505
   Width           =   4440
   Begin CommandButton cmdCancel 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   2820
      TabIndex        =   2
      Top             =   780
      Width           =   1035
   End
   Begin CommandButton cmdAdd 
      Caption         =   "追加"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   780
      Width           =   1035
   End
   Begin CommandButton cmdOver 
      Caption         =   "上書き"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   420
      TabIndex        =   0
      Top             =   780
      Width           =   1035
   End
   Begin Label lblFileExist 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "指定したファイルは存在します"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4155
   End
End
' ファイルエラー対応ダイアログ
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA
Option Explicit

Sub cmdAdd_Click ()
    FILEEX = 2
    Unload frmFILEEXist
End Sub

Sub cmdCancel_Click ()
    FILEEX = 0
    Unload frmFILEEXist
End Sub

Sub cmdOver_Click ()
    FILEEX = 1
    Unload frmFILEEXist
End Sub

Sub Form_Load ()
    top = (screen.Height - Height) / 2
    left = (screen.Width - Width) / 2
    If L_MODE = 1 Then
        lblFILEEXist.Caption = "File Exist."
        cmdOver.Caption = "Over Write"
        cmdAdd.Caption = "Add"
        cmdCancel.Caption = "Cancel"
    End If
End Sub

