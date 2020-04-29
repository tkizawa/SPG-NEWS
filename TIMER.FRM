VERSION 2.00
Begin Form frmTimer 
   Caption         =   "Timer"
   ClientHeight    =   1515
   ClientLeft      =   1770
   ClientTop       =   5880
   ClientWidth     =   2775
   Height          =   1920
   Left            =   1710
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   2775
   Top             =   5535
   Visible         =   0   'False
   Width           =   2895
   Begin Timer tmrAuto 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1560
      Top             =   240
   End
   Begin Timer tmrTimer 
      Left            =   300
      Top             =   240
   End
   Begin Timer tmrScadule 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   900
      Top             =   240
   End
End
Option Explicit

Sub tmrAuto_Timer ()
    Dim x%
    Static t%

    If TIMELIMIT = -1 Then
        Exit Sub
    End If

    t% = t% + 1
    If t% >= TIMELIMIT Then
        t% = 0
        If MODE = "FWD" Then
            frmMain!cmdFWD.Enabled = True
            frmMain!mnuComSR.Enabled = True
        Else
            frmMain!cmdAuto.Enabled = True
            frmMain!mnuAutoPilot.Enabled = True
        End If
        Idol_on
        Disconnect
        x% = SerialClose()
        COMFLAG = False
    End If
End Sub

Sub tmrScadule_Timer ()
    Dim i%, flag%, tm%
    Static t%, ts%

    ts% = ts% + 1
    frmMain!pnl3DTime.Caption = Format$(Now, "yy/mm/dd hh:mm")
    ' スプーラ起動
    Spool
    ' 通信起動スケジューラ
    ' Timer起動が有効になっていて､通信中(TIMERCOM=True)でなければ実行
    If TIMERENABLE = "ON" And TIMERCOM <> True And MODE <> "DJLMB" Then
        ' 定刻起動モード
        If TIMERMODE = 0 Then
            flag% = False
            For i% = 0 To 23
                If TIMERSET(i%) = Format$(Now, "hh:mm") Then
                    flag% = True
                End If
            Next i%
            If flag% = True Then
                If MODE = "FWD" Then
                    Forward
                Else
                    AutoPilotStart
                End If
            End If
        Else
        ' インターバル起動モード
            If ts% = TIMERINTERVAL Then
                ts% = 0
                If MODE = "FWD" Then
                    Forward
                Else
                    AutoPilotStart
                End If
            End If
        End If
    End If
End Sub

Sub tmrTimer_Timer ()
    TIMEOUTFLAG = -1
    TIMECOUNT = TIMECOUNT + 1
End Sub

