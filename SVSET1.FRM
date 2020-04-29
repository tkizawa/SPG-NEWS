VERSION 2.00
Begin Form frmServerSet1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Form1"
   ClientHeight    =   2865
   ClientLeft      =   3510
   ClientTop       =   3270
   ClientWidth     =   7365
   Height          =   3270
   Left            =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   7365
   Top             =   2925
   Width           =   7485
   Begin CommandButton cmdCancel 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5700
      TabIndex        =   5
      Top             =   1620
      Width           =   1215
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
      Left            =   5700
      TabIndex        =   4
      Top             =   960
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
      Left            =   3840
      TabIndex        =   1
      Top             =   660
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
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
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
      Left            =   3840
      TabIndex        =   3
      Top             =   1740
      Width           =   1215
   End
   Begin TextBox txtFileName 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   900
      Width           =   2775
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
      Left            =   3180
      TabIndex        =   9
      Top             =   1800
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
      Left            =   3180
      TabIndex        =   8
      Top             =   1260
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
      Left            =   3180
      TabIndex        =   7
      Top             =   720
      Width           =   330
      WordWrap        =   -1  'True
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "出力ファイル名"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   540
      Width           =   2655
   End
End
Option Explicit

Sub cmdCancel_Click ()
    Unload Me
End Sub

Sub cmdOK_Click ()
    Dim i, f As Integer

    f = False

    If txtFileName.Text = "" Or txtToField.Text = "" Or txtAtbbs.Text = "" Or txtCall.Text = "" Then
        Beep
        If L_MODE = 0 Then
            MsgBox "未設定項目があります", MB_ICONEXCLAMATION
        Else
            MsgBox "There is as unset item.", MB_ICONEXCLAMATION
        End If
        Exit Sub
    End If
    

    For i = 0 To UBound(SVNAME$)
        If SVNAME$(i) = UCase$(txtFileName.Text) Then
            If SVTFL$(i) = UCase$(txtToField.Text) Then
                If SVATBBS$(i) = UCase$(txtAtbbs.Text) Then
                    If SVCALL$(i) = UCase$(txtCall.Text) Then
                        f = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next i
    If f = True Then
        Beep
        If L_MODE = 0 Then
            MsgBox "追加した項目は設定済みです", MB_ICONEXCLAMATION
        Else
            MsgBox "Add item has been set.", MB_ICONEXCLAMATION
        End If
        Exit Sub
    Else
        i = UBound(SVNAME$) + 1
        ReDim Preserve SVNAME$(i), SVTFL$(i), SVATBBS$(i), SVCALL$(i)
        SVNAME$(i) = UCase$(txtFileName.Text)
        SVTFL$(i) = UCase$(txtToField.Text)
        SVATBBS$(i) = UCase$(txtAtbbs.Text)
        SVCALL$(i) = UCase$(txtCall.Text)
        frmServerSet!lstServer.AddItem UCase$(txtFileName.Text)
    End If
    Unload Me
End Sub

Sub Form_Load ()
    Top = frmServerSet.Top
    Left = frmServerSet.Left
    If L_MODE = 0 Then
        Caption = "アーカイブ設定追加"
    Else
        Caption = "Add Archive Set"
        Label1.Caption = "Output filename"
        cmdCancel.Caption = "Cancel"
    End If
End Sub

