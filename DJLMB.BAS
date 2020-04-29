Option Explicit

Sub GetDJLMailDir ()
    Dim FileNum, i As Integer
    Dim l$, yn$

    On Error GoTo DJL_ERR

    FileNum = FreeFile
    Open DJLMAILDIR + "\" + "DJL.MB" For Input As FileNum
    While (l$ <> "*** EOF")
        Line Input #FileNum, l$
    Wend
    Line Input #FileNum, l$
    While (l$ <> "*** EOF")
        Line Input #FileNum, l$
    Wend
    For i = 1 To 13
        Line Input #FileNum, l$
    Next i
    Close #FileNum
    yn$ = UCase$(Left$(l$, 1))
    If yn$ = "Y" Then
        ' MultiDir
        DJLMULTIDIR = True
    Else
        DJLMULTIDIR = False
    End If
    Exit Sub

DJL_ERR:
    ErrorProc "DJL-MB", Err, Erl, Error
    MsgBox "DJL-MB�ݒ�t�@�C���̓ǂݍ��݂Ɏ��s���܂����B" + Chr$(13) + Chr$(10) + "���X�g���������������Ɛ�������܂��B" + Chr$(13) + Chr$(10) + "�ʐM�������I�����܂�", MB_ICONEXCLAMATION, "SPG-NEWS"
    FileAdd USERDIR + "LOG.TXT", "DJL-MB>Error " + Format$(Err) + ":" + Format$(Erl)
    Exit Sub
End Sub

Function GetDJLMBSize () As Long
    GetDJLMBSize = FileLen(DJLMAILDIR + "\" + "MAIL.MB")
End Function
