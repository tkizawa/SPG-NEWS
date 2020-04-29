VERSION 2.00
Begin Form frmTimerSet 
   BackColor       =   &H00C0C0C0&
   Caption         =   "タイマー設定"
   ClientHeight    =   5595
   ClientLeft      =   1035
   ClientTop       =   1650
   ClientWidth     =   6960
   FontBold        =   0   'False
   FontItalic      =   0   'False
   FontName        =   "FixedSys"
   FontSize        =   13.5
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   6000
   Left            =   975
   LinkTopic       =   "Form2"
   ScaleHeight     =   5595
   ScaleWidth      =   6960
   Top             =   1305
   Width           =   7080
   Begin SSFrame fraTimer1 
      Caption         =   "通信起動時間間隔設定"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   360
      TabIndex        =   14
      Top             =   3540
      Width           =   4395
      Begin TextBox txtIntervalMinutes 
         Alignment       =   1  'Right Justify
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   420
         TabIndex        =   6
         Top             =   480
         Width           =   555
      End
      Begin Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "分毎に通信開始"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Top             =   540
         Width           =   2175
      End
   End
   Begin SSOption optTimerMode 
      Caption         =   "指定した時間間隔に通信開始"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Index           =   1
      Left            =   420
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1260
      Width           =   6435
   End
   Begin SSOption optTimerMode 
      Caption         =   "指定した時間に通信開始"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   840
      Value           =   -1  'True
      Width           =   6375
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
      Left            =   5580
      TabIndex        =   9
      Top             =   4980
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
      Left            =   4260
      TabIndex        =   7
      Top             =   4980
      Width           =   1215
   End
   Begin SSFrame fraTimer 
      Caption         =   "通信起動時間設定"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   360
      TabIndex        =   8
      Top             =   1860
      Width           =   4395
      Begin TextBox txtMinutes 
         Alignment       =   1  'Right Justify
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   1260
         TabIndex        =   5
         Top             =   900
         Width           =   555
      End
      Begin TextBox txtEndHour 
         Alignment       =   1  'Right Justify
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   1860
         TabIndex        =   4
         Top             =   360
         Width           =   555
      End
      Begin TextBox txtStartHour 
         Alignment       =   1  'Right Justify
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   420
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
      Begin Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "分に通信開始"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   960
         Width           =   1815
      End
      Begin Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   " 毎時"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   540
         TabIndex        =   12
         Top             =   960
         Width           =   675
      End
      Begin Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "時までの間に"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   480
         Width           =   1515
      End
      Begin Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "時から"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1020
         TabIndex        =   10
         Top             =   480
         Width           =   795
      End
   End
   Begin CheckBox chkTimerOn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "タイマー起動で通信を行う"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   420
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End

Sub chkTimerOn_Click ()
    If optTIMERMODE(0).Value = True Then
        txtStartHour.Enabled = True
        txtEndHour.Enabled = True
        txtMinutes.Enabled = True
        txtIntervalMinutes.Enabled = False
    Else
        txtStartHour.Enabled = False
        txtEndHour.Enabled = False
        txtMinutes.Enabled = False
        txtIntervalMinutes.Enabled = True
    End If
    '
    If chkTimerOn.Value = 1 Then
        fraTimer.Enabled = optTIMERMODE(0).Value
        fraTimer1.Enabled = Not fraTimer.Enabled
        optTIMERMODE(0).Enabled = True
        optTIMERMODE(1).Enabled = True
    Else
        fraTimer.Enabled = False
        fraTimer1.Enabled = False
        optTIMERMODE(0).Enabled = False
        optTIMERMODE(1).Enabled = False
    End If
End Sub

Sub cmdCancel_Click ()
    Unload frmTimerSet
End Sub

Sub cmdOK_Click ()
    If chkTimerOn = 1 Then
        s$ = "ON"
    Else
        s$ = "OFF"
    End If
    If txtStartHour.Text = "24" Then
        txtStartHour.Text = "00"
    End If
    If txtEndHour.Text = "24" Then
        txtEndHour.Text = "00"
    End If
    
    If Val(txtStartHour.Text) < 0 Or Val(txtStartHour.Text) > 24 Then
        MsgBox "時刻設定が間違っています。", MB_OK Or MB_ICONEXCLAMATION, "SPG-NEWS"
        txtStartHour.SetFocus
        Exit Sub
    End If
    If Val(txtEndHour.Text) < 0 Or Val(txtEndHour.Text) > 24 Then
        MsgBox "時刻設定が間違っています。", MB_OK Or MB_ICONEXCLAMATION, "SPG-NEWS"
        txtEndHour.SetFocus
        Exit Sub
    End If
    If Val(txtMinutes.Text) < 0 Or Val(txtMinutes.Text) > 59 Then
        MsgBox "時刻設定が間違っています。", MB_OK Or MB_ICONEXCLAMATION, "SPG-NEWS"
        txtMinutes.SetFocus
        Exit Sub
    End If

    TIMERENABLE = s$
    a = WritePrivateProfileString("Timer", "Enable", s$ + Chr$(0) + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Timer", "StartHour", txtStartHour.Text + Chr$(0) + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Timer", "EndHour", txtEndHour.Text + Chr$(0) + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Timer", "Minutes", txtMinutes.Text + Chr$(0) + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Timer", "IntervalMinutes", txtIntervalMinutes.Text + Chr$(0) + Chr$(0), INIFILE)
    
    ' タイマー起動モードの設定
    If optTIMERMODE(0).Value = True Then
        TIMERMODE = 0
        a = WritePrivateProfileString("Timer", "Mode", "0" + Chr$(0) + Chr$(0), INIFILE)
    Else
        TIMERMODE = 1
        a = WritePrivateProfileString("Timer", "Mode", "1" + Chr$(0) + Chr$(0), INIFILE)
    End If
    
    ' タイマー起動時間の再設定
    TIMESTARTHOUR = txtStartHour.Text
    TIMEENDHOUR = txtEndHour.Text
    TIMEMINUTES = txtMinutes.Text
    TIMERINTERVAL = txtIntervalMinutes.Text

    ' タイマー設定
    ReDim TIMERSET(24)
    i% = Val(txtStartHour.Text)
    j% = Val(txtEndHour.Text)
    k% = i%
    c% = 0
    While (k% <> j%)
        k% = i% Mod 24
        TIMERSET(c%) = Right$("00" + Mid$(Str$(k%), 2), 2) + ":" + Right$("00" + Mid$(Str$(Val(txtMinutes.Text)), 2), 2)
        c% = c% + 1
        i% = i% + 1
    Wend
    TIMERSET(c%) = "-1"
    Unload frmTimerSet
End Sub

Sub Form_Load ()
    Top = (frmMain.Height - frmTimerSet.Height) / 2
    Left = (frmMain.Width - frmTimerSet.Width) / 2
    If Trim$(TIMERENABLE) = "ON" Then
        chkTimerOn.Value = 1
        fraTimer.Enabled = True
    Else
        chkTimerOn.Value = 0
        fraTimer.Enabled = False
    End If
    txtStartHour.Text = TIMESTARTHOUR
    txtEndHour.Text = TIMEENDHOUR
    txtMinutes.Text = TIMEMINUTES
    txtIntervalMinutes.Text = TIMERINTERVAL
    If TIMERMODE = 0 Then
        fraTimer.Enabled = True
        fraTimer1.Enabled = False
        optTIMERMODE(0).Value = True
        txtStartHour.Enabled = True
        txtEndHour.Enabled = True
        txtMinutes.Enabled = True
        txtIntervalMinutes.Enabled = False
    Else
        fraTimer.Enabled = False
        fraTimer1.Enabled = True
        optTIMERMODE(1).Value = True
        txtStartHour.Enabled = False
        txtEndHour.Enabled = False
        txtMinutes.Enabled = False
        txtIntervalMinutes.Enabled = True
    End If

    If L_MODE = 1 Then
        Caption = "Timer Set"
        chkTimerOn.Caption = "Communicates by timer set"
        optTIMERMODE(0).Caption = "Begins the Commuication at the specified time"
        optTIMERMODE(1).Caption = "Begins the Commuication at the specified inteval"
        fraTimer.Caption = "Start Time"
        fraTimer1.Caption = "Interval Time"
        Label1.Caption = "to"
        Label2.Caption = ""
        Label3.Caption = ""
        Label4.Caption = "Minutes"
        Label5.Caption = "Minutes"
        cmdCancel.Caption = "Cancel"
    End If
End Sub

Sub optTIMERMODE_Click (Index As Integer, Value As Integer)
    If optTIMERMODE(0).Value = True Then
        fraTimer.Enabled = True
        fraTimer1.Enabled = False
        txtStartHour.Enabled = True
        txtEndHour.Enabled = True
        txtMinutes.Enabled = True
        txtIntervalMinutes.Enabled = False
    Else
        fraTimer.Enabled = False
        fraTimer1.Enabled = True
        txtStartHour.Enabled = False
        txtEndHour.Enabled = False
        txtMinutes.Enabled = False
        txtIntervalMinutes.Enabled = True
    End If
End Sub

