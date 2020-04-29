VERSION 2.00
Begin Form frmWEBBrowserSet 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ＷＥＢブラウザ設定"
   ClientHeight    =   3105
   ClientLeft      =   1275
   ClientTop       =   3075
   ClientWidth     =   6780
   Height          =   3510
   Left            =   1215
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   6780
   Top             =   2730
   Width           =   6900
   Begin CommonDialog CMDialog1 
      Left            =   3360
      Top             =   1920
   End
   Begin CommandButton Command3 
      Caption         =   "参照..."
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin TextBox txtWEBBrowser 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Text            =   "C:\PROGRA~1\MICROS~3\IEXPLORE.EXE"
      Top             =   960
      Width           =   4815
   End
   Begin CommandButton Command2 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   960
      Width           =   1215
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
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ＷＥＢブラウザ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
End

Sub Command1_Click ()
    Dim a%
    
    a% = WritePrivateProfileString("Option", "WEBBrowser", UCase$(txtWEBBrowser.Text), INIFILE)
    WEBBROWSER = UCase$(txtWEBBrowser.Text)
    ChDrive Left$(SYSDIR, 2)
    ChDir Left$(SYSDIR, Len(SYSDIR) - 1)
    Unload Me
End Sub

Sub Command2_Click ()
    Unload Me
    ChDrive Left$(SYSDIR, 2)
    ChDir Left$(SYSDIR, Len(SYSDIR) - 1)
End Sub

Sub Command3_Click ()
    Dim FileNum As Integer
    Dim a%, s$, t$

    On Error Resume Next
    
    CMDialog1.DefaultExt = "*"
    CMDialog1.DialogTitle = "ＷＥＢブラウザの設定"
    CMDialog1.Filter = "EXE File (*.EXE)|*.EXE|"
    ' 最初に使うフィルタを“All Files”に設定します.
    CMDialog1.FilterIndex = 1
    ' [上書き禁止] チェック ボックスをチェックします.
    CMDialog1.Flags = OFN_READONLY
    CMDialog1.CancelError = -1
    ' [ファイルを開く] ダイアログ ボックスを表示します.
    CMDialog1.Action = 1
    ' ユーザーが <OK> ボタンを選択したときの処理
    If Err = 0 Then
        s$ = CMDialog1.Filename
        txtWEBBrowser.Text = s$
    ElseIf Err = 32755 Then
        Exit Sub
    End If
End Sub

Sub Form_Load ()
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2

    txtWEBBrowser.Text = WEBBROWSER

End Sub

