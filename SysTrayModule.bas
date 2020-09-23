Attribute VB_Name = "SysTrayModule"
'      Need to add to form using this:
'      PS: Also remember to "RemoveFromTray" when your form unloads
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Dim Message As Long
'   On Error Resume Next
'    Message = x / Screen.TwipsPerPixelX
'    Select Case Message
'        'Your Choice:
'        Case WM_RBUTTONUP
'            SetForegroundWindow Me.hwnd
'            PopupMenu [Menu]
'        Case WM_RBUTTONDOWN
'            SetForegroundWindow Me.hwnd
'            PopupMenu [Menu]
'    End Select
'End Sub

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public TaskBarhWnd As Long
Public TaskBarCrashDetected As Boolean
Public TrayIconCrashDetected As Boolean

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const WM_MOUSEMOVE = &H200

'Left-click constants.
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up

'Right-click constants.
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up

Dim TrayIcon As NOTIFYICONDATA
Public Sub AddToTray(frm As Form, ToolTip As String, Icon)
    Dim Retval As Long
    Call RemoveFromTray
    TrayIconCrashDetected = False
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = frm.hwnd
    TrayIcon.szTip = ToolTip & vbNullChar
    TrayIcon.hIcon = Icon
    TrayIcon.uID = vbNull
    TrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    TrayIcon.uCallbackMessage = WM_MOUSEMOVE
    Retval = Shell_NotifyIcon(NIM_ADD, TrayIcon)
    If Format(Retval) = 0 Then
        TrayIconCrashDetected = True
        'TaskBarCrashDetected = True
    Else
        TrayIconCrashDetected = False
        'TaskBarCrashDetected = False
    End If
End Sub
Public Sub ModifyTray(frm As Form, ToolTip As String, Icon)
    Dim Retval As Long
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = frm.hwnd
    TrayIcon.szTip = ToolTip & vbNullChar
    TrayIcon.hIcon = Icon
    TrayIcon.uID = vbNull
    TrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    TrayIcon.uCallbackMessage = WM_MOUSEMOVE
    Retval = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
    If FindWindow("Shell_traywnd", vbNullString) = 0 Then TaskBarCrashDetected = True
    If Format(Retval) = 0 And TaskBarCrashDetected = False Then
        TrayIconCrashDetected = True
    Else
        TrayIconCrashDetected = False
    End If
End Sub
Public Sub RemoveFromTray()
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub
Public Function GetY()
    Dim Point As POINTAPI, Retval As Long
    Retval = GetCursorPos(Point)
    GetY = Point.Y
End Function
