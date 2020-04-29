' メッセージ管理
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

Function DataGetNum (num&)
    Dim i As Integer
    Dim j&
    Dim Maxnum As Integer

    Maxnum = UBound(MESSAGE)
    
    For i = Maxnum To 0 Step -1
        j& = Val(MESSAGE(i).fname)
        If j& = num& Then
            DataGetNum = i
            Exit Function
        End If
    Next i
    DatGetNum = -1
End Function

Function GetMessageNum () As Integer
    GetMessageNum = MESSAGERECORDNUM
End Function

Sub SetMessageNum (i As Integer)
    MESSAGERECORDNUM = i
End Sub

' 文字列カット
'　　NULLコード以降の文字列をカットする
' dummy : 元文字列
' ss : 変換後文字列
'
Function StrCut (dummy As String) As String
    Dim i As Integer

    s$ = ""
    For i = 1 To Len(dummy)
        tmp$ = Mid$(dummy, i, 1)
        If tmp$ = Chr$(0) Then
            GoTo out
        Else
            s$ = s$ + tmp$
        End If
    Next i
out:
    StrCut = s$
End Function

