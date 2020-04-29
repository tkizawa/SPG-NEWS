' ＷｉｎｄｏｗｓＡＰＩ定義

' User Profile Routines
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFilename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFilename As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFilename As String) As Integer
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

'NOTE: Catch/Throw expect a long pointer to an 18-byte buffer (lpCatchBuf)
' eg:
'   Dim Buffer(1 To 9) As Integer
'   result% = Catch (Buffer(1))

Declare Function Catch Lib "Kernel" (lpCatchBuf As Any) As Integer
Declare Sub Throw Lib "Kernel" (lpCatchBuf As Any, ByVal nThrowBack As Integer)

Declare Sub FatalExit Lib "Kernel" (ByVal Code As Integer)
Declare Sub SwapRecording Lib "Kernel" (ByVal wFlag As Integer)

' 動作中のアプリケーションのインスタンス数チェック

Declare Function GetModuleHandle Lib "Kernel" (ByVal lpProgName As String) As Integer
Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer

' BitBlt API
Declare Function BitBlt% Lib "Gdi" (ByVal destDC%, ByVal X%, ByVal Y%, ByVal W%, ByVal h%, ByVal srcDC%, ByVal xSrc%, ByVal ySrc%, ByVal RasterOp&)
Global Const SRCCOPY = &HCC0020

' SendMessage
Declare Function SendMESSAGE Lib "User" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal LParam&) As Integer

Global Const WM_USER = &H400
Global Const EM_LINESCROLL = WM_USER + 6

' NKF
Declare Function SetNkfOption Lib "nkfwin.dll" (ByVal optStr$) As Integer
Declare Sub NkfConvert Lib "nkfwin.dll" (ByVal outstring As Any, ByVal instring As Any)

