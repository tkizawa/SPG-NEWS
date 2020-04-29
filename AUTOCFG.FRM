VERSION 2.00
Begin Form frmAutoConfig 
   BackColor       =   &H00C0C0C0&
   Caption         =   "オートパイロット設定"
   ClientHeight    =   5550
   ClientLeft      =   1575
   ClientTop       =   1620
   ClientWidth     =   8610
   FontBold        =   0   'False
   FontItalic      =   0   'False
   FontName        =   "MS Sans Serif"
   FontSize        =   8.25
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   5955
   Left            =   1515
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8610
   Top             =   1275
   Width           =   8730
   Begin CommandButton cmdAutoOption 
      Caption         =   "ｵﾌﾟｼｮﾝ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   435
      Left            =   7740
      TabIndex        =   26
      Top             =   1440
      Width           =   795
   End
   Begin SSFrame Frame3D3 
      Caption         =   "タイトル入力時のプロンプト"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   120
      TabIndex        =   23
      Top             =   3600
      Width           =   8415
      Begin TextBox txtPrompt 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   2535
      End
      Begin Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "よくわからない場合は空欄にしてください"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   360
         Width           =   5295
      End
   End
   Begin SSFrame Frame3D2 
      Caption         =   "取得リスト数"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   4140
      TabIndex        =   21
      Top             =   4560
      Width           =   3975
      Begin TextBox txtListCount 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Text            =   "50"
         Top             =   300
         Width           =   915
      End
      Begin Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "一度に取得するリスト数"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "標準明朝"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   1260
         TabIndex        =   22
         Top             =   360
         Width           =   2055
      End
   End
   Begin SSFrame Frame3D1 
      Caption         =   "ファイル番号"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   3975
      Begin TextBox txtListNum 
         Alignment       =   1  'Right Justify
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Text            =   "100"
         Top             =   300
         Width           =   915
      End
      Begin Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "リストをとりたい番号を入力してください"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "標準明朝"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   435
         Left            =   1320
         TabIndex        =   20
         Top             =   360
         Width           =   2535
      End
   End
   Begin CommandButton cmdAutoSetCan 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   435
      Left            =   7740
      TabIndex        =   12
      Top             =   780
      Width           =   795
   End
   Begin CommandButton cmdAutoSetOK 
      Caption         =   "OK"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   435
      Left            =   7740
      TabIndex        =   11
      Top             =   240
      Width           =   795
   End
   Begin SSFrame fra3DRBBS 
      Caption         =   "ＲＢＢＳ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin TextBox txtReadCmd 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   4920
         TabIndex        =   27
         Text            =   "R"
         Top             =   2280
         Width           =   975
      End
      Begin SSOption optListCommand 
         Caption         =   " L xxxx-yyyy"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Index           =   1
         Left            =   3180
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1755
      End
      Begin SSOption optListCommand 
         Caption         =   " L xxxx yyyy"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Index           =   0
         Left            =   3180
         TabIndex        =   6
         Top             =   1260
         Value           =   -1  'True
         Width           =   1935
      End
      Begin TextBox txtTitleCol 
         Alignment       =   1  'Right Justify
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Text            =   "48"
         Top             =   2400
         Width           =   975
      End
      Begin TextBox txtCallCol 
         Alignment       =   1  'Right Justify
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Text            =   "24"
         Top             =   1980
         Width           =   975
      End
      Begin TextBox txtAtbbsCol 
         Alignment       =   1  'Right Justify
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Text            =   "31"
         Top             =   1560
         Width           =   975
      End
      Begin TextBox txtToCol 
         Alignment       =   1  'Right Justify
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Text            =   "17"
         Top             =   1140
         Width           =   975
      End
      Begin ComboBox cboRBBS 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   390
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "リードコマンド"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   3120
         TabIndex        =   28
         Top             =   2280
         Width           =   1815
      End
      Begin Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "リストコマンド出力形式"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   3180
         TabIndex        =   25
         Top             =   840
         Width           =   2955
      End
      Begin Label lblCol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "桁数（先頭は１）"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   840
         Width           =   2355
      End
      Begin Label lblRbbsRem 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "通信相手のＲＢＢＳのリスト表示のフォーマットを指定します"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   17
         Top             =   3000
         Width           =   7035
      End
      Begin Label lblTitleCol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Title"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2460
         Width           =   975
      End
      Begin Label lblCallCol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "From"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   975
      End
      Begin Label lblAtbbsCol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "@BBS"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1620
         Width           =   975
      End
      Begin Label lblToCol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "To Field"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   975
      End
   End
End
' オートパイロット設定ダイアログ
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Sub cboRBBS_Click ()
    Select Case cboRBBS.Text
        Case "RLI/DJL"
            txtToCol.Text = 17
            txtAtbbsCol.Text = 31
            txtCallCol.Text = 24
            txtTitleCol.Text = 48
            optListCommand(0).Value = True
        Case "FBB"
            txtToCol.Text = 16
            txtAtbbsCol.Text = 23
            txtCallCol.Text = 30
            txtTitleCol.Text = 48
            optListCommand(1).Value = True
        Case "FOS"
            txtToCol.Text = 29
            txtAtbbsCol.Text = 36
            txtCallCol.Text = 22
            txtTitleCol.Text = 50
            optListCommand(0).Value = True
        Case "MSYS"
            txtToCol.Text = 16
            txtAtbbsCol.Text = 30
            txtCallCol.Text = 23
            txtTitleCol.Text = 44
            optListCommand(0).Value = True
    End Select
End Sub

Sub cmdAutoOption_Click ()
    frmAutoOptionSet.Show 1
End Sub

Sub cmdAutoSetCan_Click ()
    Unload frmAutoConfig
End Sub

Sub cmdAutoSetOK_Click ()
    
' SPGNEWS.INI更新
    a = WritePrivateProfileString("AutoPilot", "RBBS", cboRBBS.Text, INIFILE)
    a = WritePrivateProfileString("AutoPilot", "ToFieldColum", txtToCol.Text, INIFILE)
    a = WritePrivateProfileString("AutoPilot", "AtbbsColum", txtAtbbsCol.Text, INIFILE)
    a = WritePrivateProfileString("AutoPilot", "CallColum", txtCallCol.Text, INIFILE)
    a = WritePrivateProfileString("AutoPilot", "TitleColum", txtTitleCol.Text, INIFILE)
    a = WritePrivateProfileString("AutoPilot", "ListNum", UCase$(txtListNum.Text), INIFILE)
    a = WritePrivateProfileString("AutoPilot", "ListCount", UCase$(txtListCount.Text), INIFILE)
    a = WritePrivateProfileString("AutoPilot", "TitlePrompt", txtPrompt.Text, INIFILE) ' UCase$をとる(V3.4.2)
    a = WritePrivateProfileString("AutoPilot", "ReadCommand", txtReadCmd.Text, INIFILE)

    Select Case cboRBBS.Text
        Case "RLI/DJL"
            a = WritePrivateProfileString("AutoPilot", "ListCommand", "0", INIFILE)
        Case "FBB"
            a = WritePrivateProfileString("AutoPilot", "ListCommand", "1", INIFILE)
        Case Else
            If optListCommand(0).Value = True Then
                a = WritePrivateProfileString("AutoPilot", "ListCommand", "0", INIFILE)
            Else
                a = WritePrivateProfileString("AutoPilot", "ListCommand", "1", INIFILE)
            End If
    End Select
    TITLEINPUT = txtPrompt.Text
    READCMD = txtReadCmd.Text

    Unload frmAutoConfig
End Sub

Sub Form_Load ()
    Dim ss As String * 100
    Dim sr As String * 100
    Dim st As String * 300
    Dim sb As String * 300
    Dim sc As String * 300
    Dim ln As String * 40
    Dim lc As String * 40
    Dim pr As String * 200
    Dim listcommand As Integer
    Dim i As Integer
    
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2
    
    TOFIELDCOL = GetPrivateProfileInt("AutoPilot", "ToFieldColum", 17, INIFILE)
    ATBBSCOL = GetPrivateProfileInt("AutoPilot", "AtbbsColum", 31, INIFILE)
    CALLCOL = GetPrivateProfileInt("AutoPilot", "CallColum", 24, INIFILE)
    TITLECOL = GetPrivateProfileInt("AutoPilot", "TitleColum", 48, INIFILE)
    listcommand = GetPrivateProfileInt("AutoPilot", "ListCommand", 0, INIFILE)
    a = GetPrivateProfileString("AutoPilot", "RBBS", "RLI/DJL", ss, 80, INIFILE)
    s$ = RNullCut(ss)
    a = GetPrivateProfileString("AutoPilot", "ListNum", "1000", ln, 20, INIFILE)
    a = GetPrivateProfileString("AutoPilot", "ListCount", "50", lc, 20, INIFILE)
    a = GetPrivateProfileString("AutoPilot", "TitlePrompt", "", pr, 160, INIFILE)

    cboRBBS.Clear
    cboRBBS.AddItem "RLI/DJL"
    cboRBBS.AddItem "FBB"
    cboRBBS.AddItem "FOS"
    cboRBBS.AddItem "MSYS"
    cboRBBS.AddItem "Other"
    
    Select Case s$
        Case "RLI/DJL"
            cboRBBS.ListIndex = 0
        Case "FBB"
            cboRBBS.ListIndex = 1
        Case "FOS"
            cboRBBS.ListIndex = 2
        Case "MSYS"
            cboRBBS.ListIndex = 3
        Case Else
            cboRBBS.ListIndex = 4
    End Select

    txtListNum.Text = RNullCut(ln)
    txtListCount.Text = RNullCut(lc)
'    txtPrompt.Text = RNullCut(pr)
    txtPrompt.Text = TITLEINPUT
    txtToCol.Text = Format$(TOFIELDCOL)
    txtAtbbsCol.Text = Format$(ATBBSCOL)
    txtCallCol.Text = Format$(CALLCOL)
    txtTitleCol = Format$(TITLECOL)

    optListCommand(listcommand).Value = True

    txtReadCmd.Text = READCMD

End Sub

