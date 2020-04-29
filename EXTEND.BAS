Option Explicit

Sub FontListMenu ()
    Trace "FontListMenu"
    ' [フォントの指定] ダイアログ ボックスのフラグ
    Const CF_BOTH = &H3&
    Const CF_EFFECTS = &H100&
    ' フラグをセットします.
'    CMDlgFont.Flags = CF_BOTH Or CF_EFFECTS
    ' ダイアログ ボックスの初期値を設定します.
'    CMDlgFont.FontSize = frmMain!lstMenu.FontSize
'    CMDlgFont.FontName = frmMain!lstMenu.FontName
    ' [フォントの指定] ダイアログ ボックスを表示します.
'    CMDlgFont.Action = 4
    ' 選択したオプションに従って, テキスト ボックスの
    ' テキストのフォントを変更します.
'    frmMain!lstMenu.FontName = CMDlgFont.FontName
'    frmMain!lstMenu.FontSize = CMDlgFont.FontSize

'    frmMain!lstMenu.FontBold = CMDlgFont.FontBold
'    frmMain!lstMenu.FontItalic = CMDlgFont.FontItalic
'    frmMain!lstMenu.FontUnderline = CMDlgFont.FontUnderLine
'    frmMain!lstMenu.FontStrikethru = CMDlgFont.FontStrikeThru
'    frmMain!lstMenu.ForeColor = CMDlgFont.Color
    
'    frmMain!lstListTitle.FontName = CMDlgFont.FontName
'    frmMain!lstListTitle.FontSize = CMDlgFont.FontSize
'    a = WritePrivateProfileString("Font", "ListMenuFontName", frmMain!lstListTitle.FontName, INIFILE)
'    a = WritePrivateProfileString("Font", "ListMenuFontSize", Format$(frmMain!lstListTitle.FontSize), INIFILE)
End Sub

