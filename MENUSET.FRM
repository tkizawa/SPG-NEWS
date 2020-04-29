VERSION 2.00
Begin Form frmMenuSet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "リストメニューカスタム設定"
   ClientHeight    =   4065
   ClientLeft      =   1155
   ClientTop       =   1515
   ClientWidth     =   7740
   Height          =   4470
   Left            =   1095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   7740
   Top             =   1170
   Width           =   7860
   Begin CommandButton cmdCancel 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin CommandButton cmdDel 
      Caption         =   "削除"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin CommandButton cmdChange 
      Caption         =   "変更"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin CommandButton cmdAdd 
      Caption         =   "追加"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   360
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
      Left            =   6360
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin ComboBox cboLabel 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   390
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin TextBox txtToField 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin TextBox txtAtbbs 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1380
      Width           =   2535
   End
   Begin TextBox txtCall 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label2"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1095
      Left            =   360
      TabIndex        =   13
      Top             =   2640
      Width           =   5775
   End
   Begin Label lblCall 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Call:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   1980
      Width           =   615
   End
   Begin Label lblAtbbs 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "@BBS"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   720
      TabIndex        =   11
      Top             =   1440
      Width           =   480
   End
   Begin Label lblToField 
      BackColor       =   &H00C0C0C0&
      Caption         =   "To:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   720
      TabIndex        =   10
      Top             =   900
      Width           =   330
      WordWrap        =   -1  'True
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ラベル名"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   960
   End
End

Sub cboLabel_Click ()
    Dim i As Integer

    i = cboLabel.ListIndex
    MENUSETINDEX = i
    txtToField.Text = MENUSET$(i, 1)
    txtAtbbs.Text = MENUSET$(i, 2)
    txtCall.Text = MENUSET$(i, 3)
End Sub

Sub cmdAdd_Click ()
    Dim i As Integer

    If SetCheck() = False Then
        MsgBox "すべての項目を入力して下さい", MB_ICONEXCLAMATION, "SPG-NEWS"
        Exit Sub
    End If

    For i = 0 To 100
        If MENUSET$(i, 0) = "" Then
            Exit For
        End If
    Next i
    If i = 100 Then
        MsgBox "これ以上追加できません", MB_ICONEXCLAMATION, "SPG-NEWS"
        Exit Sub
    End If
    MENUSET$(i, 0) = cboLabel.Text
    MENUSET$(i, 1) = UCase$(txtToField.Text)
    MENUSET$(i, 2) = UCase$(txtAtbbs.Text)
    MENUSET$(i, 3) = UCase$(txtCall.Text)
    cboLabel.AddItem cboLabel.Text
End Sub

Sub cmdCancel_Click ()
    Unload Me
End Sub

Sub cmdChange_Click ()
    Dim i As Integer

    If SetCheck() = False Then
        MsgBox "すべての項目を入力して下さい", MB_ICONEXCLAMATION, "SPG-NEWS"
        Exit Sub
    End If

    MENUSET$(MENUSETINDEX, 0) = cboLabel.Text
    MENUSET$(MENUSETINDEX, 1) = UCase$(txtToField.Text)
    MENUSET$(MENUSETINDEX, 2) = UCase$(txtAtbbs.Text)
    MENUSET$(MENUSETINDEX, 3) = UCase$(txtCall.Text)
    
    cboLabel.List(MENUSETINDEX) = cboLabel.Text
    cboLabel.ListIndex = MENUSETINDEX
End Sub

Sub cmdDel_Click ()
    Dim i, c, j, k As Integer

    c = cboLabel.ListCount - 1
    i = cboLabel.ListIndex
    
    For j = i To c
        MENUSET$(j, 0) = MENUSET$(j + 1, 0)
        MENUSET$(j, 1) = MENUSET$(j + 1, 1)
        MENUSET$(j, 2) = MENUSET$(j + 1, 2)
        MENUSET$(j, 3) = MENUSET$(j + 1, 3)
    Next j

    cboLabel.RemoveItem i
    If i <> c Then
        cboLabel.ListIndex = i
        txtToField.Text = MENUSET$(i, 0)
        txtAtbbs.Text = MENUSET$(i, 1)
        txtCall.Text = MENUSET$(i, 2)
    Else
        txtToField.Text = ""
        txtAtbbs.Text = ""
        txtCall.Text = ""
    End If
End Sub

Sub Command1_Click ()
    Dim FileNum As Integer

    If FileExist(USERDIR + "LISTMENU.TXT") = True Then
        Kill USERDIR + "LISTMENU.TXT"
    End If
    
    If MENUSET$(0, 0) <> "" Then
        FileNum = FreeFile
        Open USERDIR + "LISTMENU.TXT" For Output As FileNum
        Do Until MENUSET$(i, 0) = ""
            Print #FileNum, MENUSET$(i, 0) + "," + MENUSET$(i, 1) + "," + MENUSET$(i, 2) + "," + MENUSET$(i, 3)
            i = i + 1
        Loop
        Close #FileNum
    End If

    ' リスト表示コンポボックス削除
    frmMain!cboList.ListIndex = 3
    i = 3
    Do While frmMain!cboList.List(i) <> "ごみ箱"
           frmMain!cboList.RemoveItem i
    Loop
    frmMain!cboList.RemoveItem i

    ' リスト表示コンポボックス追加
    i = 0
    Do While MENUSET$(i, 0) <> ""
        frmMain!cboList.AddItem MENUSET$(i, 0)
        i = i + 1
    Loop
    frmMain!cboList.AddItem "送信保留"
    frmMain!cboList.AddItem "ごみ箱"

    frmMain!cboList.ListIndex = 0
    Unload Me
End Sub

Sub Form_Load ()
    MENUSETINDEX = 0
    Width = 8000
    Height = 5000
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2

    Label2.Caption = "新規追加：各項目を設定して追加ボタンを押す" + Chr$(13) + Chr$(10) + "変更：各項目を設定して追加ボタンを押す" + Chr$(13) + Chr$(10) + "削除：削除したいラベルを指定して削除ボタンを押す"

    If FileExist(USERDIR + "LISTMENU.TXT") = True Then
        i = 0
        Do While MENUSET$(i, 0) <> ""
            cboLabel.AddItem MENUSET$(i, 0)
            i = i + 1
        Loop
        cboLabel.ListIndex = 0
        txtToField.Text = MENUSET$(0, 1)
        txtAtbbs.Text = MENUSET$(0, 2)
        txtCall.Text = MENUSET$(0, 3)
    End If
End Sub

Function SetCheck () As Integer
    If cboLabel.Text = "" Or txtToField.Text = "" Or txtAtbbs.Text = "" Or txtCall.Text = "" Then
        SetCheck = False
    Else
        SetCheck = True
    End If
End Function

