Option Explicit

'
' ゲージ表示プログラム
'(VisualBasicパワフルテクニック P.295)
'
Sub GaugeDisplay (p!)
    Dim Percent$
    Dim OldScaleMode%, r%

    Percent$ = "メッセージ読み込み中 " + Format$(p!, "0%")
    frmOpen!picCount.Cls
    frmOpen!picCount1.Cls
    frmOpen!picCount.CurrentX = (frmOpen!picCount.Width - frmOpen!picCount.TextWidth(Percent$)) / 2
    frmOpen!picCount1.CurrentX = frmOpen!picCount.CurrentX
    frmOpen!picCount.CurrentY = (frmOpen!picCount.Height - frmOpen!picCount.TextHeight(Percent$)) / 2
    frmOpen!picCount1.CurrentY = frmOpen!picCount.CurrentY
    frmOpen!picCount.Print Percent$
    frmOpen!picCount1.Print Percent$
    OldScaleMode% = frmOpen!picCount.Parent.ScaleMode
    frmOpen!picCount.Parent.ScaleMode = 3
    r% = BitBlt(frmOpen!picCount.hDC, 0, 0, frmOpen!picCount1.Width * p!, frmOpen!picCount1.Height, frmOpen!picCount1.hDC, 0, 0, SRCCOPY)
    frmOpen!picCount.Parent.ScaleMode = OldScaleMode%
End Sub

