VERSION 2.00
Begin Form frmDownSet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "ダウンロード設定"
   ClientHeight    =   4260
   ClientLeft      =   1980
   ClientTop       =   3390
   ClientWidth     =   7365
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "FixedSys"
   FontSize        =   13.5
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   4665
   Left            =   1920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3769.912
   ScaleMode       =   0  'User
   ScaleWidth      =   7365
   Top             =   3045
   Width           =   7485
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
      Left            =   5880
      TabIndex        =   4
      Top             =   2940
      Width           =   1215
   End
   Begin Grid grdDownSet 
      Cols            =   6
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   3735
      Left            =   420
      TabIndex        =   0
      Top             =   300
      Width           =   5205
   End
   Begin CommandButton cmdClose 
      Caption         =   "ｷｬﾝｾﾙ"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   3540
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
      Left            =   5880
      TabIndex        =   3
      Top             =   1620
      Width           =   1215
   End
   Begin CommandButton cmdDel 
      Caption         =   "削除"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin CommandButton cmdAdd 
      Caption         =   "追加"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   300
      Width           =   1215
   End
End
Option Explicit

Sub cmdAdd_Click ()
    DOWNSETMODE = "ADD"
    frmDownSet1.Show 1
End Sub

Sub cmdChange_Click ()
    DOWNSETMODE = "CHG"
    frmDownSet1.Show 1
End Sub

Sub cmdClose_Click ()
    Unload Me
End Sub

Sub cmdDel_Click ()
    Dim i As Integer
    
    On Error Resume Next

    If grdDownSet.Rows = 2 Then
        grdDownSet.Row = 1
        For i = 1 To 5
            grdDownSet.Col = i
            grdDownSet.Text = ""
        Next i
        Exit Sub
    End If
    grdDownSet.RemoveItem grdDownSet.SelStartRow
End Sub

Sub cmdOK_Click ()
    Dim i, j, FileNum As Integer
    Dim a$

    If (FILEEXist(USERDIR + "DOWNLOAD.SPG") = True) Then
        Kill USERDIR + "DOWNLOAD.SPG"
    End If
    ' 設定が空かチェックして空ならばDOWNLOAD.SPGを作成しない
    grdDownSet.Row = 1
    grdDownSet.Col = 1
    If grdDownSet.Text <> "" Then
        FileNum = FreeFile
        Open USERDIR + "DOWNLOAD.SPG" For Output As #FileNum
        For i = 1 To grdDownSet.Rows - 1
            grdDownSet.Row = i
            grdDownSet.Col = 1
            If grdDownSet.Text = "する" Then
                a$ = "@"
            Else
                a$ = "!"
            End If
            For j = 2 To 5
                grdDownSet.Col = j
                a$ = a$ + "," + grdDownSet.Text
            Next j
            Print #FileNum, a$
        Next i
        Close #FileNum
    End If
    Unload Me
End Sub

Sub DownSet ()
    Dim FileNum, i As Integer
    Dim a$, b$, tfl$, atbbs$, callsign$, title$, t$

    t$ = Chr$(9)
    FileNum = FreeFile
    i = 1
    grdDownSet.Col = 1
    If FILEEXist(USERDIR + "DOWNLOAD.SPG") = True Then
        Open USERDIR + "DOWNLOAD.SPG" For Input As #FileNum
        Do While (EOF(FileNum) = False)
            Input #FileNum, a$, tfl$, atbbs$, callsign$, title$
    '        grdDownSet.Row = i
            If a$ = "@" Then
                If L_MODE = 0 Then
                    b$ = "する"
                Else
                    b$ = "Yes"
                End If
            Else
                If L_MODE = 0 Then
                    b$ = "拒否"
                Else
                    b$ = "No"
                End If
            End If
    
            grdDownSet.AddItem "" + t$ + b$ + t$ + tfl$ + t$ + atbbs$ + t$ + callsign$ + t$ + title$
            i = i + 1
            If i >= 128 Then Exit Do
        Loop
        Close #FileNum
        grdDownSet.RemoveItem 1
    End If
End Sub

Sub Form_Load ()
    Dim i As Integer

    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2

    grdDownSet.ColWidth(0) = 500
    grdDownSet.ColWidth(1) = 800
    grdDownSet.ColWidth(2) = 800
    grdDownSet.ColWidth(3) = 800
    grdDownSet.ColWidth(4) = 800
    grdDownSet.ColWidth(5) = 5000

    grdDownSet.Col = 0
    grdDownSet.Row = 0
    grdDownSet.Col = 1
    If L_MODE = 0 Then
        grdDownSet.Text = "受信"
    Else
        grdDownSet.Text = "Rcv"
    End If
    grdDownSet.Col = 2
    grdDownSet.Text = "To"
    grdDownSet.Col = 3
    grdDownSet.Text = "@BBS"
    grdDownSet.Col = 4
    grdDownSet.Text = "From"
    grdDownSet.Col = 5
    grdDownSet.Text = "Title"

    DownSet
    
    If L_MODE = 1 Then
        cmdAdd.Caption = "Add"
        cmdDel.Caption = "Delete"
        cmdChange.Caption = "Change"
        cmdClose.Caption = "Cancel"
    End If
End Sub

Sub InputDisEnable ()
'    txtToField.Enabled = False
'    txtAtbbs.Enabled = False
'    txtCall.Enabled = False
'    txtTitle.Enabled = False
'    optRcv(0).Enabled = False
'    optRcv(1).Enabled = False
End Sub

Sub InputEnable ()
'    txtToField.Enabled = True
'    txtAtbbs.Enabled = True
'    txtCall.Enabled = True
'    txtTitle.Enabled = True
'    optRcv(0).Enabled = True
'    optRcv(1).Enabled = True
End Sub

