VERSION 2.00
Begin Form frmDownSet1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '固定(2重線)
   Caption         =   "ダウンロード設定"
   Height          =   3930
   Left            =   3540
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7290
   Top             =   4605
   Width           =   7410
   Begin CommandButton cmdClose 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5940
      TabIndex        =   8
      Top             =   2760
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
      Left            =   4560
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin SSOption optRcv 
      Caption         =   "受信する"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Index           =   0
      Left            =   540
      TabIndex        =   4
      Top             =   2400
      Value           =   -1  'True
      Width           =   1335
   End
   Begin SSOption optRcv 
      Caption         =   "受信拒否"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1395
   End
   Begin TextBox txtToField 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   1740
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin TextBox txtAtbbs 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   1740
      TabIndex        =   1
      Top             =   780
      Width           =   975
   End
   Begin TextBox txtCall 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   1740
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin TextBox txtTitle 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   1740
      TabIndex        =   3
      Top             =   1860
      Width           =   3255
   End
   Begin Label lblDown 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label5"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1455
      Left            =   2940
      TabIndex        =   12
      Top             =   300
      Width           =   4155
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "To Field"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   300
      Width           =   1155
   End
   Begin Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "@BBS"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   780
      Width           =   1155
   End
   Begin Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "From"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   1320
      Width           =   1155
   End
   Begin Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Title"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   1860
      Width           =   795
   End
End
Option Explicit

Sub cmdClose_Click ()
    Unload frmDownSet1
End Sub

Sub cmdOk_Click ()
    Dim i As Integer
    Dim a$, b$, c$, d$, e$, t$, y$, n$, s1$, s2$

    t$ = Chr$(9)

    If L_MODE = 0 Then
        y$ = "する"
        n$ = "拒否"
        s1$ = "追加した項目は設定済みです"
        s2$ = "変更した項目は設定済みです"
    Else
        y$ = "Yes"
        n$ = "No"
        s1$ = "Add item has been set."
        s2$ = "Change item has been set."
    End If

    frmDownSet1!txtToField.Text = UCase$(frmDownSet1!txtToField.Text)
    frmDownSet1!txtAtbbs.Text = UCase$(frmDownSet1!txtAtbbs.Text)
    frmDownSet1!txtCall.Text = UCase$(frmDownSet1!txtCall.Text)
    If frmDownSet1!txtToField.Text = "" Then
        frmDownSet1!txtToField.Text = "*"
    End If
    If frmDownSet1!txtAtbbs.Text = "" Then
        frmDownSet1!txtAtbbs.Text = "*"
    End If
    If frmDownSet1!txtCall.Text = "" Then
       frmDownSet1!txtCall.Text = "*"
    End If
    If frmDownSet1!txtTitle.Text = "" Then
       frmDownSet1!txtTitle.Text = "*"
    End If
    If DOWNSETMODE = "ADD" Then
        If frmDownSet1!optRcv(0).Value = True Then
            a$ = y$
        Else
            a$ = n$
        End If
        If DownSetCheck() = True Then
            Beep
            MsgBox s1$, MB_ICONEXCLAMATION
            Exit Sub
        End If
        b$ = frmDownSet1!txtToField.Text
        c$ = frmDownSet1!txtAtbbs.Text
        d$ = frmDownSet1!txtCall.Text
        e$ = frmDownSet1!txtTitle.Text
        ' 何も設定されていないグリッドに追加された場合は追加せずに変更する
        If frmDownSet!grdDownSet.Rows = 2 And frmDownSet!grdDownSet.Text = "" Then
            i = frmDownSet!grdDownSet.SelStartRow
            frmDownSet!grdDownSet.Row = i
            frmDownSet!grdDownSet.Col = 1
            frmDownSet!grdDownSet.Text = a$
            frmDownSet!grdDownSet.Col = 2
            frmDownSet!grdDownSet.Text = b$
            frmDownSet!grdDownSet.Col = 3
            frmDownSet!grdDownSet.Text = c$
            frmDownSet!grdDownSet.Col = 4
            frmDownSet!grdDownSet.Text = d$
            frmDownSet!grdDownSet.Col = 5
            frmDownSet!grdDownSet.Text = e$
            frmDownSet!grdDownSet.Col = 1
        Else
            frmDownSet!grdDownSet.AddItem "" + t$ + a$ + t$ + b$ + t$ + c$ + t$ + d$ + t$ + e$
        End If
    Else
        If DownSetCheck() = True Then
            Beep
            MsgBox s2$, MB_ICONEXCLAMATION
           Exit Sub
        End If
        b$ = frmDownSet1!txtToField.Text
        c$ = frmDownSet1!txtAtbbs.Text
        d$ = frmDownSet1!txtCall.Text
        e$ = frmDownSet1!txtTitle.Text
        i = frmDownSet!grdDownSet.SelStartRow
        frmDownSet!grdDownSet.Row = i
        frmDownSet!grdDownSet.Col = 2
        frmDownSet!grdDownSet.Text = b$
        frmDownSet!grdDownSet.Col = 3
        frmDownSet!grdDownSet.Text = c$
        frmDownSet!grdDownSet.Col = 4
        frmDownSet!grdDownSet.Text = d$
        frmDownSet!grdDownSet.Col = 5
        frmDownSet!grdDownSet.Text = e$
        frmDownSet!grdDownSet.Col = 1
        If frmDownSet1!optRcv(0).Value = True Then
            frmDownSet!grdDownSet.Text = y$
        Else
            frmDownSet!grdDownSet.Text = n$
        End If
    End If
    Unload Me
End Sub

Function DownSetCheck ()
    Dim a, i As Integer

    a = frmDownSet!grdDownSet.Rows - 1
    For i = 1 To a
        frmDownSet!grdDownSet.Row = i
        frmDownSet!grdDownSet.Col = 2
        If frmDownSet1!txtToField.Text = frmDownSet!grdDownSet.Text Then
            frmDownSet!grdDownSet.Col = 3
            If frmDownSet1!txtAtbbs.Text = frmDownSet!grdDownSet.Text Then
                frmDownSet!grdDownSet.Col = 4
                If frmDownSet1!txtCall.Text = frmDownSet!grdDownSet.Text Then
                    frmDownSet!grdDownSet.Col = 5
                    If frmDownSet1!txtTitle.Text = frmDownSet!grdDownSet.Text Then
                        frmDownSet!grdDownSet.Col = 1
                        If frmDownSet!grdDownSet.Text = "する" And optRcv(0).Value = True Then
                            DownSetCheck = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
    DownSetCheck = False
End Function

Sub Form_Load ()
    Dim a As Integer

    If DOWNSETMODE = "CHG" Then
        If L_MODE = 0 Then
            frmDownSet1.Caption = "ダウンロード設定［変更］"
        Else
            frmDownSet1.Caption = "Down Load Set [Change]"
        End If
        a = frmDownSet!grdDownSet.SelStartRow
        frmDownSet!grdDownSet.Row = a
        frmDownSet!grdDownSet.Col = 2
        txtToField = frmDownSet!grdDownSet.Text
        frmDownSet!grdDownSet.Col = 3
        txtAtbbs = frmDownSet!grdDownSet.Text
        frmDownSet!grdDownSet.Col = 4
        txtCall = frmDownSet!grdDownSet.Text
        frmDownSet!grdDownSet.Col = 5
        txtTitle = frmDownSet!grdDownSet.Text
        frmDownSet!grdDownSet.Col = 1
        If frmDownSet!grdDownSet.Text = "する" Or frmDownSet!grdDownSet.Text = "Yes" Then
            optRcv(0).Value = True
        Else
            optRcv(1).Value = True
        End If
    Else
        If L_MODE = 0 Then
            frmDownSet1.Caption = "ダウンロード設定［追加］"
        Else
            frmDownSet1.Caption = "Down Load Set [Add]"
        End If
        optRcv(0).Value = True
    End If

    If L_MODE = 0 Then
        lblDown.Caption = "ダウンロードしたい宛先を指定して下さい。特に指定しない部分は空欄にしてください。"
    Else
        lblDown.Caption = "Input the address whitch wants to be down-load."
        optRcv(0).Caption = "Receive"
        optRcv(1).Caption = "Reject"
        cmdClose.Caption = "Cancel"
    End If
End Sub

