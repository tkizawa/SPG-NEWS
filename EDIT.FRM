VERSION 2.00
Begin Form frmEdit 
   Caption         =   "SPG-NEWS for Windows テキストエディタ"
   ClientHeight    =   5145
   ClientLeft      =   1425
   ClientTop       =   2565
   ClientWidth     =   10020
   ControlBox      =   0   'False
   Height          =   5835
   Left            =   1365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   10020
   Top             =   1935
   Width           =   10140
   Begin PictureBox picEdit 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   9990
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   10020
      Begin CommandButton cmdHTMLEdit 
         Caption         =   "HTML Editor"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   8400
         TabIndex        =   10
         Top             =   60
         Width           =   975
      End
      Begin ComboBox cboNewsgroups 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   390
         Left            =   6960
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   600
         Width           =   2295
      End
      Begin CommandButton cmdEditor 
         Caption         =   "ｴﾃﾞｨﾀ起動"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   6960
         TabIndex        =   4
         Top             =   60
         Width           =   1395
      End
      Begin TextBox txtTitle 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   1020
         TabIndex        =   2
         Top             =   600
         Width           =   5895
      End
      Begin ComboBox cboAtbbs 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   390
         Left            =   5280
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin ComboBox cboToField 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   390
         Left            =   1020
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   3495
      End
      Begin Label lblAtBBS 
         BackColor       =   &H00C0C0C0&
         Caption         =   "@BBS"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   180
         Width           =   675
      End
      Begin Label lblTitle 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ﾀｲﾄﾙ"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   660
         Width           =   675
      End
      Begin Label lblToField 
         BackColor       =   &H00C0C0C0&
         Caption         =   "宛先"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   675
      End
   End
   Begin TextBox txtEdit 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   3195
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
      Top             =   1200
      Width           =   9600
   End
   Begin Menu mnuFile 
      Caption         =   "ファイル"
      Begin Menu mnuEditNew 
         Caption         =   "新規作成"
      End
      Begin Menu mnuNo12 
         Caption         =   "-"
      End
      Begin Menu mnuEditCancel 
         Caption         =   "取り消し"
      End
      Begin Menu mnuReserve 
         Caption         =   "送信保留"
      End
      Begin Menu mnuEditExit 
         Caption         =   "送信予約"
      End
   End
   Begin Menu mnuEdit 
      Caption         =   "編集(&E)"
      Begin Menu mnuCut 
         Caption         =   "切り取り"
         Shortcut        =   ^X
      End
      Begin Menu mnuCopy 
         Caption         =   "コピー"
         Shortcut        =   ^C
      End
      Begin Menu mnuPaste 
         Caption         =   "貼り付け"
         Shortcut        =   ^V
      End
   End
   Begin Menu mnuHelp0 
      Caption         =   "ヘルプ(&H)"
      Begin Menu mnuHelp 
         Caption         =   "目次"
         Shortcut        =   {F1}
      End
      Begin Menu mnuAbout 
         Caption         =   "バージョン情報"
      End
   End
End
' テキストエディタ ｢キザワード｣
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA
Option Explicit

Sub cboAtbbs_GotFocus ()
'    IMEVALUE = IMEStatus
'    cboAtbbs.ImeMode = 2
End Sub

Sub cboAtbbs_LostFocus ()
'    cboAtbbs.ImeMode = IMEVALUE
End Sub

Sub cboToField_GotFocus ()
'    IMEVALUE = IMEStatus
'    cboToField.ImeMode = 2
End Sub

Sub cboToField_LostFocus ()
'    cboToField.ImeMode = IMEVALUE
End Sub

Sub cmdEditor_Click ()
    Dim a

    On Error GoTo cmdEditor_Click_err

    FileSave (USERDIR + EXTFILENAME)
    frmEdit.txtEdit.Text = ""
    a = Shell(EDITOR + " " + USERDIR + EXTFILENAME, 1)
    Exit Sub

cmdEditor_Click_err:
    MsgBox "指定した外部テキストエディタ [ " + EDITOR + " ] が起動できません。", 48, "外部テキストエディタの起動"
    Exit Sub
End Sub

Sub cmdHTMLEdit_Click ()
    Dim a, FileNum, FileNum1, s$
    Dim html_flag
    
    On Error GoTo cmdHTMLEdit_Click_err

    FileSave (USERDIR + "readhtml.txt")
    frmEdit.txtEdit.Text = ""
    
    FileNum = FreeFile
    html_flag = False
    Open USERDIR + "readhtml.txt" For Input As FileNum
    FileNum1 = FreeFile
    Open USERDIR + EXTFILENAME For Output As FileNum1
    Print #FileNum1, "<HTML>"
    While Not EOF(FileNum)
        Line Input #FileNum, s$
        If s$ = "<HTML>" Then
            html_flag = True
        End If
        If s$ = "</HTML>" Then
            html_flag = False
        End If
        If html_flag = False Then
            s$ = s$ + "<BR>"
        End If
        Print #FileNum1, s$
    Wend
    Print #FileNum1, "</HTML>"
    Close #FileNum
    Close #FileNum1
    html_flag = False

    a = Shell(HTMLEDITOR + " " + USERDIR + EXTFILENAME, 1)
    Exit Sub

cmdHTMLEdit_Click_err:
    MsgBox "指定したＨＴＭＬエディタ [ " + HTMLEDITOR + " ] が起動できません。", 48, "ＨＴＭＬエディタの起動"
    Exit Sub
End Sub

Sub Form_Activate ()
    WINDOW_MODE = 2
End Sub

Sub Form_Load ()
    ReadWinPos frmEdit, "Editor"
    txtEdit.Text = ""
'    cboToField.ImeMode = 2
'    cboAtbbs.ImeMode = 2
'    txtEdit.ImeMode = 1
'    txtTitle.ImeMode = 1

    If L_MODE = 1 Then
        lblToField.Caption = "To"
        lblTitle.Caption = "Title"
        cmdEditor.Caption = "Ext.Editor"
        '
        mnuFile.Caption = "File"
        mnuEditNew.Caption = "New"
        mnuEditCancel.Caption = "Edit Abort"
        mnuReserve.Caption = "Reserve"
        mnuEditExit.Caption = "Exit"
        '
        mnuEdit.Caption = "Edit(&E)"
        mnuCut.Caption = "Cut"
        mnuCopy.Caption = "Copy"
        mnuPaste.Caption = "Paste"
        '
        mnuHelp0.Caption = "Help"
        mnuHelp.Caption = "Contents"
        mnuAbout.Caption = "About"
    End If
    If LAPUTAMODE = True Then
        cboNewsgroups.Visible = True
        cboNewsgroups.Clear
        cboNewsgroups.AddItem "version-up"
        cboNewsgroups.AddItem "bug"
        cboNewsgroups.AddItem "subway"
        cboNewsgroups.AddItem "all"
        cboNewsgroups.AddItem "misc"
    Else
        cboNewsgroups.Visible = False
    End If
    EXTFILENAME = Left$(Time$, 2) + Mid$(Time$, 4, 2) + Right$(Time$, 2) + ".HTM"
End Sub

Sub Form_Resize ()
    On Error Resume Next ' 最小時のプロパティエラー発生防止

    If frmEdit.ScaleWidth > 9975 Then
        txtEdit.Width = 9975
    Else
        txtEdit.Width = frmEdit.ScaleWidth
    End If
    txtEdit.Top = picEdit.Height
    txtEdit.Height = frmEdit.ScaleHeight - txtEdit.Top
End Sub

Sub Form_Unload (Cancel As Integer)
    Dim a As Integer
    Dim tocall, title As String

    If frmEdit.WindowState <> 1 Then
        WrtWinPos frmEdit, "Editor"
    End If
    Exit Sub
'
    If txtEdit.Text <> "" Then
        a = MsgBox("メッセージを書き込みますか?", MB_YESNOCANCEL Or MB_ICONINFORMATION Or MB_DEFBUTTON1, "SPG-NEWS")
        Select Case a
        Case IDYES
            If (Len(tocall) = 0) Or Len(title) = 0 Then
                MsgBox "宛て先とタイトルを指定してください。", 48, "書き込み"
                Exit Sub
            End If
            Call mnuEditExit_Click
        Case IDNO
            Unload frmEdit
            ToolBarEnable
        Case IDCANCEL
            Exit Sub
        End Select
    End If
End Sub

Sub mnuAbout_Click ()
    OpenAbout
End Sub

Sub mnuCopy_Click ()
    'テキスト複写
    ClipBoard.SetText frmEdit.ActiveControl.SelText
End Sub

Sub mnuCut_Click ()
    'テキスト複写
    ClipBoard.SetText frmEdit.ActiveControl.SelText
    'テキスト削除
    frmEdit.ActiveControl.SelText = ""
End Sub

Sub mnuEditCancel_Click ()
    Dim a%, s$

    If L_MODE = 0 Then
        s$ = "書き込みを中断しますか？"
    Else
        s$ = "Write abort?"
    End If

    If txtEdit.Text <> "" Then
        If MsgBox(s$, MB_YESNO Or MB_ICONQUESTION) = IDYES Then
            Unload frmEdit
            ToolBarEnable
        Else
            Exit Sub
        End If
    Else
        Unload frmEdit
        ToolBarEnable
    End If
End Sub

Sub mnuEditExit_Click ()
    Dim tocall As String
    Dim field As String
    Dim atbbs As String
    Dim title As String
    Dim wmode As String
    Dim a, i, flag As Integer

    NEWSGROUP = "." + cboNewsgroups.Text
    
    If WRITEMODE = "B" Then
        tocall = cboToField.Text + "@" + cboAtbbs.Text
        ' To FieldがSYSOPだったらWRITEMODE="P"とする
 '       If UCase$(cboToField.Text) = "SYSOP" Then
 '           WRITEMODE = "P"
 '       End If
    Else
        tocall = cboToField.Text
        '@BBSが無い場合は自分と同じHomeRBBSにする
'        If instr(tocall, "@") = 0 Then
'            tocall = cboToField.Text + "@" + HOMERBBS
'        End If
    End If

    title = txtTitle.Text

    If (Len(tocall) = 0) Or Len(title) = 0 Then
        MsgBox "宛て先とタイトルを指定してください。", 48, "書き込み"
        Exit Sub
    End If
    If txtEdit = "" And FileExist(USERDIR + EXTFILENAME) = False Then
        MsgBox "文書を作成してください。", 48, "書き込み"
        Exit Sub
    End If

    ' tocallからfieldとatbbsを分離する
    tocall = SpaceCut(tocall)
    If InStr(tocall, "@") = 0 Then
        field = tocall
        atbbs = ""
    Else
        i = 1
        Do While (Mid$(tocall, i, 1) <> "@")
            field = field + Mid$(tocall, i, 1)
            i = i + 1
        Loop
        atbbs = Mid$(tocall, i + 1)
    End If

    If WRITEMODE = "B" Then
        wmode = "B$"
    Else
        wmode = "P "
    End If

    FileSave (USERDIR + "SEND.TXT")
    If R_MODE = True Then
        flag = False
    Else
        flag = True
    End If
    a = FileWrite(USERDIR + "SEND.TXT", wmode, field, atbbs, title, "*", flag)
    If R_MODE = True Then
        MenuFileKill
        FileListMake "!K"
        R_MODE = False
    End If

    Unload frmEdit
    ToolBarEnable
End Sub

Sub mnuEditNew_Click ()
    txtEdit = ""
End Sub

Sub mnuHelp_Click ()
    OpenHelp
End Sub

Sub mnuPaste_Click ()
    txtEdit.SetFocus
    'テキスト配置
    frmEdit.ActiveControl.SelText = ClipBoard.GetText()
End Sub

Sub mnuReserve_Click ()
    Dim tocall As String
    Dim field As String
    Dim atbbs As String
    Dim title As String
    Dim wmode As String
    Dim a, i, flag As Integer

    NEWSGROUP = "." + cboNewsgroups.Text
    
    If WRITEMODE = "B" Then
        tocall = cboToField.Text + "@" + cboAtbbs.Text
        ' To FieldがSYSOPだったらWRITEMODE="P"とする
        If UCase$(cboToField.Text) = "SYSOP" Then
            WRITEMODE = "P"
        End If
    Else
        tocall = cboToField.Text
        '@BBSが無い場合は自分と同じHomeRBBSにする
'        If instr(tocall, "@") = 0 Then
'            tocall = cboToField.Text + "@" + HOMERBBS
'        End If
    End If

    title = txtTitle.Text

    If (Len(tocall) = 0) Or Len(title) = 0 Then
        MsgBox "宛て先とタイトルを指定してください。", 48, "書き込み"
        Exit Sub
    End If
    If txtEdit = "" And FileExist(USERDIR + EXTFILENAME) = False Then
        MsgBox "文書を作成してください。", 48, "書き込み"
        Exit Sub
    End If

    ' tocallをfieldとatbbsに分離する
    tocall = SpaceCut(tocall)
    If InStr(tocall, "@") = 0 Then
        field = tocall
        atbbs = ""
    Else
        i = 1
        Do While (Mid$(tocall, i, 1) <> "@")
            field = field + Mid$(tocall, i, 1)
            i = i + 1
        Loop
        atbbs = Mid$(tocall, i + 1)
    End If

    If WRITEMODE = "B" Then
        wmode = "B$H"
    Else
        wmode = "P H"
    End If

    FileSave (USERDIR + "SEND.TXT")
    If R_MODE = True Then
        flag = False
    Else
        flag = True
    End If
    a = FileWrite(USERDIR + "SEND.TXT", wmode, field, atbbs, title, "*", flag)
    If R_MODE = True Then
        MenuFileKill
        FileListMake "!K"
        R_MODE = False
    End If

    Unload frmEdit
    ToolBarEnable
End Sub

