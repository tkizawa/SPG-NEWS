Option Explicit

'
' スプール
' 一つのファイルを書き出す
' タイマー起動
'
Sub Spool ()
    Dim filename$
    Dim r%

    On Error GoTo Spool_Err

    If SPOOLDIR + "\" = SYSDIR Or SPOOLDIR = MSGDIR Or SPOOLDIR = USERDIR Then
        Exit Sub
    End If
    
    filename$ = Dir$(SPOOLDIR + "\", 0)
    r% = SpoolOut(SPOOLDIR + "\" + filename$)
    Exit Sub

Spool_Err:
    Exit Sub
End Sub

'
' スプーラ
' ディレクトリにある全部のファイルを出力する
' メニューから起動
'
Sub SpoolAll ()
    Dim filename$
    Dim r%

    On Error GoTo SpoolAll_Err

    If SPOOLDIR + "\" = SYSDIR Or SPOOLDIR = MSGDIR Then
        Exit Sub
    End If

    filename$ = Dir$(SPOOLDIR + "\", 0)
    Do While filename$ <> ""
        Debug.Print filename$
        r% = SpoolOut(SPOOLDIR + "\" + filename$)
        filename$ = Dir$
    Loop
    Exit Sub

SpoolAll_Err:
    Exit Sub
End Sub

'
' 指定したファイルに基づき書込みを行う
'
Function SpoolOut (f$) As Integer
    Dim FileNum As Integer
    Dim FileNum2 As Integer
    Dim s$, cmd$
    Dim msg As MSGD2
    Dim HoldFlag As Integer

    On Error GoTo SpoolOut_Error

    HoldFlag = False
    Screen.MousePointer = 11
    FileNum = FreeFile
    Open f$ For Input As #FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, s$
        ' コマンド解析
        cmd$ = UCase$(Left$(s$, 1))
        If cmd$ = "S" Then
            ' コマンドライン解析
            CmdAna s$, msg
            ' タイトル設定
            Line Input #FileNum, s$
            msg.Title = s$
            ' メッセージを保存する
            FileNum2 = FreeFile
            Open USERDIR + "\SPOOLOUT.TXT" For Output As FileNum2
            Line Input #FileNum, s$
            Do While s$ <> "/EX" ' Or s$ <> "/ex"  Or EOF(FileNum) = False
                Print #FileNum2, s$
                Line Input #FileNum, s$
            Loop
            Close #FileNum2
            If Left$(msg.stat, 1) = "B" Then
                If SPOOLUP = True Then
                    msg.stat = "B$"
                Else
                    msg.stat = "BF"
                End If
            Else
                If SPOOLUP = True Then
                    msg.stat = "P "
                Else
                    msg.stat = "PF"
                End If
            End If
            If msg.Call = "" Then
                msg.Call = MYCALL
            End If
            If SPOOLUP = True Then
                If (FileWrite(USERDIR + "\SPOOLOUT.TXT", msg.stat, msg.field, msg.atbbs, msg.Title, "*", True) <> 0) Then
                    SpoolOut = False
                    Exit Function
                End If
            Else
                If (FileWrite2(USERDIR + "\SPOOLOUT.TXT", msg.stat, msg.field, msg.atbbs, msg.Call, msg.Title, msg.bid, "*", True) <> 0) Then
                    SpoolOut = False
                    Exit Function
                End If
            End If
        Else
            SpoolOut = False
            Close #FileNum
            Kill f$
            Screen.MousePointer = 0
            Exit Function
        End If
    Loop
    Close #FileNum
    SpoolOut = True
    Kill f$
    Screen.MousePointer = 0
    Exit Function

SpoolOut_Error:
    SpoolOut = False
    Screen.MousePointer = 0
    Exit Function
End Function

