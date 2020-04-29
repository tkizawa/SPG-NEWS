VERSION 2.00
Begin Form frmListNext 
   BackColor       =   &H00C0C0C0&
   Caption         =   "次回読み出し候補選択"
   ClientHeight    =   5520
   ClientLeft      =   2970
   ClientTop       =   4260
   ClientWidth     =   9105
   Height          =   5925
   Left            =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9105
   Top             =   3915
   Width           =   9225
   Begin ListBox lstNextRead0 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1110
      Left            =   300
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1620
      Visible         =   0   'False
      Width           =   6975
   End
   Begin CommandButton cmdReadNextKill 
      Caption         =   "未読メッセージファイル削除"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   3375
   End
   Begin CommandButton cmdCancel 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   7380
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   6060
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin SSFrame fraNextRead2 
      Caption         =   "次回読み込み予定メッセージ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   8475
      Begin ListBox lstNextRead2 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1110
         Left            =   180
         TabIndex        =   3
         Top             =   420
         Width           =   7935
      End
   End
   Begin SSFrame fraNextRead1 
      Caption         =   "未読メッセージ一覧"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8475
      Begin ListBox lstNextRead1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1110
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Width           =   7935
      End
   End
End
' ダウンロード予約設定ダイアログ
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Sub cmdCancel_Click ()
    Unload frmListNext
End Sub

Sub cmdOK_Click ()
    On Error Resume Next
    Kill USERDIR + "READNEXT.TXT"
    For i% = 0 To lstNextRead2.ListCount - 1
        a$ = lstNextRead2.List(i%)
        FileAdd USERDIR + "READNEXT.TXT", a$
    Next i%
    Unload frmListNext
End Sub

Sub cmdReadNextKill_Click ()
    On Error Resume Next
    Kill USERDIR + "LIST.TXT"
'    Unload frmListNext             ' B240
End Sub

Sub Form_Load ()
    Dim Filenum As Integer

    On Error Resume Next
    
    ReadWinPos frmListNext, "ListNext"

    ' 未読一覧設定
    Screen.MousePointer = 11
'    lstNextRead1.Clear
    lstNextRead0.Clear
    Filenum = FreeFile
    Open USERDIR + "LIST.TXT" For Input As Filenum
    If Err <> 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    Do While (EOF(Filenum) = False)
        Line Input #Filenum, a$
'        lstNextRead1.AddItem a$
        lstNextRead0.AddItem a$
    Loop
    Close Filenum

    ' 未読一覧作成
    MakeList

    ' ダウンロード予約一覧設定
    lstNextRead2.Clear
    Filenum = FreeFile
    Open USERDIR + "READNEXT.TXT" For Input As Filenum
    If Err = 0 Then
        Do While (EOF(Filenum) = False)
            Line Input #Filenum, a$
            lstNextRead2.AddItem a$
        Loop
        Close Filenum
    End If
    Screen.MousePointer = 0
End Sub

Sub Form_Resize ()
    fraNextRead1.Width = (frmListNext.ScaleWidth - 120) - 360
    lstNextRead1.Width = (fraNextRead1.Width - 60) - 360
    fraNextRead2.Width = (frmListNext.ScaleWidth - 120) - 360
    lstNextRead2.Width = (fraNextRead2.Width - 60) - 360
    cmdOk.Left = frmListNext.ScaleWidth - (cmdOk.Width + 120 + cmdCancel.Width + 540)
    cmdCancel.Left = cmdOk.Left + cmdOk.Width + 120
    '
    a = (frmListNext.ScaleHeight - 60 - 120 - 180 - cmdReadNextKill.Height - 120) / 2
    fraNextRead1.Height = a
    fraNextRead2.Height = a
    fraNextRead2.Top = fraNextRead1.Top + fraNextRead1.Height + 120
    cmdReadNextKill.Top = fraNextRead2.Top + fraNextRead2.Height + 180
    cmdOk.Top = cmdReadNextKill.Top
    cmdCancel.Top = cmdReadNextKill.Top
    lstNextRead1.Height = fraNextRead1.Height - 420 - 120
    lstNextRead2.Height = fraNextRead2.Height - 420 - 120
End Sub

Sub Form_Unload (Cancel As Integer)
    If frmListNext.WindowState <> 1 Then
        WrtWinPos frmListNext, "ListNext"
    End If
End Sub

Sub lstNextRead1_DblClick ()
    a$ = lstNextRead1.Text
    For i% = 0 To lstNextRead2.ListCount - 1
        If lstNextRead2.List(i%) = a$ Then
            Exit Sub
        End If
    Next i%
    lstNextRead2.AddItem a$
End Sub

Sub lstNextRead2_DblClick ()
    idx = lstNextRead2.ListIndex
    lstNextRead2.RemoveItem idx
End Sub

Sub MakeList ()
    Dim c%, i%, a$, a1$

    c% = lstNextRead0.ListCount
    For i% = c% - 1 To 0 Step -1
        a$ = lstNextRead0.List(i%)
        If a$ <> a1$ Then
            lstNextRead1.AddItem a$
            a1$ = a$
        End If
    Next i%
End Sub

