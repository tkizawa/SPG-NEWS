Option Explicit
' SPG-NEWS for Windows�f�o�b�O�p���W���[��
' Copyright (c) 1991-1996 by JO1SPG/T.KIZAWA

' �G���[�\��/���O�L�^
'
' �G���[�\�������ăG���[���O�̋L�^������
'
' 1995.4.15
'
Sub ErrorProc (s As String, ec As Integer, el As Integer, es As String)
    Dim Filenum As Integer
    Dim ErrorName$
    
    Trace "ErrorProc"
    
    On Error GoTo ErrorProc_Error
    
    Beep
    ErrorName$ = s + ":" + " Err=" + Format$(ec) + " Erl=" + Format$(el) + " " + Chr$(34) + es + Chr$(34)
'    MsgBox "�G���[���������܂��� " + Chr$(13) + Chr$(10) + ErrorName$, MB_ICONEXCLAMATION, "SPG-NEWS"
    PrtInfo "�G���[���������܂��� " + Chr$(13) + Chr$(10) + ErrorName$
    ' �f�o�b�O�t���O����True�Ȃ�΃G���[���O���L�^����
    Filenum = FreeFile
    Open USERDIR + "ERROR.LOG" For Append As #Filenum
    Print #Filenum, Date$ + " " + Time$ + " " + ErrorName$
    Close #Filenum
    Exit Sub

ErrorProc_Error:
    MsgBox "�G���[�������ɃG���[���������܂����B�����I�����܂��B", MB_ICONEXCLAMATION
    End
End Sub

' ���W���[���g���[�X
'
' ���s�������W���[�����L�^����
' 1995.4.15
'
Sub Trace (s As String)
    If DEBUGFLAG = False Then
        Exit Sub
    End If
    
    Dim Filenum As Integer

    On Error GoTo Trace_Error

    Filenum = FreeFile
    Open USERDIR + "TRACE.TXT" For Append As Filenum
    If s = "Main" Then
        Print #Filenum, String$(79, "-")
    End If
    Print #Filenum, Format$(Now, "yy/mm/dd hh:mm:ss") + " : " + s
    Close #Filenum
    Exit Sub

Trace_Error:
    ErrorProc "Trace", Err, Erl, Error
    Close #Filenum
    Exit Sub
End Sub
