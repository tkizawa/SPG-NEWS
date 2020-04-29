' 変数型定義
' Copyright (c) 1993-1996 by JO1SPG/T.KIZAWA

' ファイル管理情報
' SPG-NEWS for Windowsオリジナル

Type MSGD2
    fname As String
    mediatype As String
    savemode As String
    stat As String
    field As String
    atbbs As String
    call As String
    size As String
    date As String
    time As String
    Title As String
    bid As String
    fstat As String
End Type

' DJL-MBのMAIL.MB構造
Type DJLMSG
    id As String * 2
    readnum As String * 2
    num As String * 2
    size As String * 2
    stat1 As String * 1
    stat2 As String * 1
    field As String * 7
    call1 As String * 7
    atbbs As String * 7
    rdate As String * 7
    rtime As String * 5
    call2 As String * 7
    msgnum As String * 2
    Title As String * 81
    sdate As String * 7
    stime As String * 5
    ccnum As String * 1
    bid As String * 13
    cccall As String * 246
    cccallstat As String * 107
End Type

