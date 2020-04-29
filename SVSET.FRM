VERSION 2.00
Begin Form frmServerSet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "アーカイブの設定"
   ClientHeight    =   4590
   ClientLeft      =   3570
   ClientTop       =   6495
   ClientWidth     =   7875
   Height          =   4995
   Left            =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   7875
   Top             =   6150
   Width           =   7995
   Begin CommandButton cmdDelete 
      Caption         =   "削除"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   6420
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin CommandButton cmdAdd 
      Caption         =   "追加..."
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5100
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin CommandButton cmdChange 
      Caption         =   "変更"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5100
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin TextBox txtCall 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin TextBox txtAtbbs 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   3180
      Width           =   1215
   End
   Begin TextBox txtToField 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin ListBox lstServer 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1110
      Left            =   240
      TabIndex        =   1
      Top             =   1260
      Width           =   7455
   End
   Begin CheckBox chkServerOn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "アーカイバを使う"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   3255
   End
   Begin CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5100
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
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
      Left            =   6420
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "出力ファイル名選択"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   2655
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
      Left            =   3060
      TabIndex        =   12
      Top             =   2700
      Width           =   330
      WordWrap        =   -1  'True
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
      Left            =   3060
      TabIndex        =   11
      Top             =   3240
      Width           =   480
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
      Left            =   3060
      TabIndex        =   10
      Top             =   3780
      Width           =   615
   End
End
Option Explicit

Sub chkServerOn_Click ()
'    If chkServerOn.Value = 1 Then
'        fraServer.Enabled = True
'    Else
'        fraServer.Enabled = False
'    End If
End Sub

Sub cmdAdd_Click ()
    frmServerSet1.Show 1
End Sub

Sub cmdCancel_Click ()
    Unload frmServerSet
End Sub

Sub cmdChange_Click ()
    Dim a%
    a% = lstServer.ListIndex
    SVTFL$(a%) = UCase$(txtToField.Text)
    SVATBBS$(a%) = UCase$(txtAtbbs.Text)
    SVCALL$(a%) = UCase$(txtCall.Text)
End Sub

Sub cmdDelete_Click ()
    Dim a%, i%
    a% = lstServer.ListIndex

    If UBound(SVNAME$) > 0 Then
        For i% = a% To UBound(SVNAME$) - 1
            SVNAME$(i%) = SVNAME$(i% + 1)
            SVTFL$(i%) = SVTFL$(i% + 1)
            SVATBBS$(i%) = SVATBBS$(i% + 1)
            SVCALL$(i%) = SVCALL$(i% + 1)
        Next i%
        lstServer.RemoveItem a%
        a% = UBound(SVNAME$) - 1
        ReDim Preserve SVNAME$(a%), SVTFL$(a%), SVATBBS$(a%), SVCALL$(a%)
    End If
End Sub

Sub cmdOK_Click ()
    Dim a, i, FileNum As Integer
    Dim s$

    If chkServerOn = 1 Then
        s$ = "ON"
    Else
        s$ = "OFF"
    End If
    SERVERENABLE = s$
    a = WritePrivateProfileString("Server", "Enable", s$ + Chr$(0) + Chr$(0), INIFILE)
    
    If FILEEXist(USERDIR + "SERVER.SPG") = True Then
        Kill USERDIR + "SERVER.SPG"
    End If
    FileNum = FreeFile
    Open USERDIR + "SERVER.SPG" For Output As #FileNum
    For i = 0 To UBound(SVNAME$)
        If SVNAME$(i) <> "" And SVTFL$(i) <> "" And SVATBBS$(i) <> "" And SVCALL$(i) <> "" Then
            Print #FileNum, SVNAME$(i)
            Print #FileNum, SVTFL$(i) + "," + SVATBBS$(i) + "," + SVCALL$(i)
        End If
    Next i
    Close #FileNum
    Unload frmServerSet
End Sub

Sub Form_Load ()
    Dim i, FileNum As Integer
    Dim s As String
    Dim f$, t$, a$, fr$
    
    Width = 8000
    Height = 5000
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2
    cmdOk.Default = True
    If SERVERENABLE = "ON" Then
        chkServerOn.Value = 1
    Else
        chkServerOn.Value = 0
    End If
    
    If L_MODE = 1 Then
        Caption = "Archive Setup"
        chkServerOn.Caption = "Use Archive"
        Label2.Caption = "Select output filename"
        cmdAdd.Caption = "Add..."
        cmdChange.Caption = "Change"
        cmdDelete.Caption = "Delete"
        cmdCancel.Caption = "Cancel"
    End If
    
    ' SERVER.SPG読み込み
    i = 0
    ReDim Preserve SVNAME$(0), SVTFL$(0), SVATBBS$(0), SVCALL$(0)
    FileNum = FreeFile
    lstServer.Clear
    If FILEEXist(USERDIR + "SERVER.SPG") = True Then
        Open USERDIR + "SERVER.SPG" For Input As #FileNum
        ' サーバー稼働
        Do While (EOF(FileNum) = False)
            ReDim Preserve SVNAME$(i), SVTFL$(i), SVATBBS$(i), SVCALL$(i)
            Line Input #FileNum, SVNAME$(i)
            Input #FileNum, SVTFL$(i), SVATBBS$(i), SVCALL$(i)
            lstServer.AddItem SVNAME$(i)
            i = i + 1
        Loop
        Close #FileNum
        If i > 0 Then
            lstServer.ListIndex = 0
            txtToField.Text = SVTFL$(0)
            txtAtbbs.Text = SVATBBS$(0)
            txtCall.Text = SVCALL$(0)
        End If
    End If
End Sub

Sub lstServer_Click ()
    Dim a%
    a% = lstServer.ListIndex
    txtToField.Text = SVTFL$(a%)
    txtAtbbs.Text = SVATBBS$(a%)
    txtCall.Text = SVCALL$(a%)
End Sub

