VERSION 2.00
Begin Form frmTimeKill 
   BackColor       =   &H00C0C0C0&
   Caption         =   "メッセージ削除"
   ClientHeight    =   1455
   ClientLeft      =   2895
   ClientTop       =   1755
   ClientWidth     =   4485
   ControlBox      =   0   'False
   Height          =   1860
   Left            =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4485
   Top             =   1410
   Width           =   4605
   Begin TextBox txtTimeKill 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Text            =   "7"
      Top             =   180
      Width           =   495
   End
   Begin CommandButton cmdTimeKillCancel 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   3180
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin CommandButton cmdTimeKillOk 
      Caption         =   "OK"
      Default         =   -1  'True
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   1860
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "日以前のメッセージを削除"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   660
      TabIndex        =   3
      Top             =   240
      Width           =   3015
   End
End

Sub cmdTimeKillCancel_Click ()
    Unload frmTimeKill
End Sub

Sub cmdTimeKillOk_Click ()
    Dim d As Integer

    On Error GoTo cmdTimeKillOk_Click_Err

    d = Format(txtTimekill.Text)
    If d <> 0 Then
        MouseWait
        FileTimeKill "*", "*", "*", d
        a = WritePrivateProfileString("Config", "DeleteDay", txtTimekill.Text, INIFILE)
        FILECHANGEFLAG = True
        MouseNormal
    End If
    Unload frmTimeKill
    FileListMake "!K"
    Exit Sub

cmdTimeKillOk_Click_Err:
    txtTimekill.Text = ""
    Exit Sub
End Sub

Sub Form_Load ()
    Dim d As Integer
    
    Top = (frmMain.Height - Height) / 2
    Left = (frmMain.Width - Width) / 2
    
    d = GetPrivateProfileInt("Config", "DeleteDay", 7, INIFILE)
    txtTimekill.Text = Format(d)
    
    If L_MODE = 1 Then
        Caption = "Delete Messase"
        Label1.Caption = " Before day of message"
        cmdTimeKillCancel.Caption = "Cancel"
    End If
End Sub

