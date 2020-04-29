Option Explicit

Sub ServerOut (num&, stat$, field$, atbbs$, from$, title$, bid$)
    Dim FileNum, FileNum1, FileNum2 As Integer
    Dim f$, fn%, fname$
    Dim t$, a$, fr$, linetext$

    On Error GoTo Server_Out_Err
    
    If SERVERENABLE = "OFF" Then
        Exit Sub
    End If

    If FILEEXist(USERDIR + "SERVER.SPG") = False Then
        Exit Sub
    End If
    FileNum = FreeFile
    Open USERDIR + "SERVER.SPG" For Input As #FileNum
    ' �T�[�o�[�ғ�
    Do While (EOF(FileNum) = False)
        Line Input #FileNum, f$
        Input #FileNum, t$, a$, fr$
        ' �G���[����
        If f$ = "" Or t$ = "" Or a$ = "" Or fr$ = "" Then
            If L_MODE = 0 Then
                MsgBox "SERVER.SPG�̋L�q�ɖ�肪����܂�", MB_ICONEXCLAMATION
            Else
                MsgBox "Missing SERVER.SPG", MB_ICONEXCLAMATION
            End If
            Close #FileNum
            Exit Sub
        End If

        ' ����
        If field$ Like t$ = True Then
            If atbbs$ Like a$ = True Then
                If from$ Like fr$ = True Then
                    ' ��������
                    FileNum1 = FreeFile
                    fname$ = FileNameMake(Format$(num&))
                    Open fname$ For Input As #FileNum1
                    FileNum2 = FreeFile
                    If FILEEXist(f$) = False Then
                        fn% = FreeFile
                        Open f$ For Output As #fn%
                        Close #fn%
                    End If
                    Open f$ For Append As #FileNum2
                    Print #FileNum2, "S" + Left$(stat$, 1) + " " + field$ + " @ " + atbbs$; " < " + from$;
                    If Mid$(stat, 2, 1) = "$" Then
                        Print #FileNum2, " $" + bid;
                    End If
                    Print #FileNum2, ""
                    Print #FileNum2, title
            
                    Do While Not EOF(FileNum1)
                        Line Input #FileNum1, linetext$
                        Print #FileNum2, linetext$
                    Loop
                    Print #FileNum2, "/EX"
                    Close #FileNum1
                    Close #FileNum2
                End If
            End If
        End If
    Loop
    Close #FileNum
    Exit Sub

Server_Out_Err:
    ErrorProc "ServerOut", Err, Erl, Error
    PrtInfo "�A�[�J�C�o�o�͂ł��܂���" + Chr$(13) + Chr$(10) + "���X�g���������������Ɛ�������܂��B" + Chr$(13) + Chr$(10) + "�ʐM�������I�����܂�"
    FileAdd USERDIR + "LOG.TXT", "ServerOut>Error " + Format$(Err) + ":" + Format$(Erl)
    Disconnect
    Idol_on
    Exit Sub
End Sub
