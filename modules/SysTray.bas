Attribute VB_Name = "SysTray"
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Tray_Icon As StdPicture

Public Const GWL_USERDATA = (-21&)
Public Const GWL_WNDPROC = (-4&)
Public Const WM_USER = &H400&

Public Const WM_MOUSEMOVE = &H200&
Public Const WM_LBUTTONDOWN = &H201&
Public Const WM_LBUTTONUP = &H202&
Public Const WM_LBUTTONDBLCLK = &H203&
Public Const WM_RBUTTONDOWN = &H204&
Public Const WM_RBUTTONUP = &H205&
Public Const WM_RBUTTONDBLCLK = &H206&


Public Const NIM_ADD = &H0&
Public Const NIM_MODIFY = &H1&
Public Const NIM_DELETE = &H2&
Public Const NIF_MESSAGE = &H1&
Public Const NIF_ICON = &H2&
Public Const NIF_TIP = &H4&

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Sub TRAY_Create(TRAY_Object As Object)
 Dim TrayData As NOTIFYICONDATA
 
 TrayData.hwnd = TRAY_Object.hwnd
 TrayData.szTip = "      NetMSG v1.0" & vbCrLf & "Right Click for Menu" & vbNullChar
 TrayData.hIcon = Tray_Icon.Handle
 TrayData.cbSize = Len(TrayData)
 TrayData.uCallbackMessage = WM_MOUSEMOVE
 TrayData.uID = vbNull
 TrayData.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
 dReturn = Shell_NotifyIcon(NIM_ADD, TrayData)
End Sub

Sub TRAY_SetTip(TRAY_Object As Object, Tip As String)
 Dim TrayData As NOTIFYICONDATA
 
 TrayData.hwnd = TRAY_Object.hwnd
 TrayData.szTip = Tip & vbNullChar
 TrayData.cbSize = Len(TrayData)
 TrayData.uID = vbNull
 TrayData.uFlags = NIF_TIP
 
 dReturn = Shell_NotifyIcon(NIM_MODIFY, TrayData)
End Sub

Sub TRAY_SetIcon(TRAY_Object As Object, Icon As StdPicture)
 Dim TrayData As NOTIFYICONDATA
 
 TrayData.hwnd = TRAY_Object.hwnd
 TrayData.hIcon = Icon.Handle
 TrayData.cbSize = Len(TrayData)
 TrayData.uID = vbNull
 TrayData.uFlags = NIF_ICON
 
 dReturn = Shell_NotifyIcon(NIM_MODIFY, TrayData)
End Sub

Sub TRAY_Close(TRAY_Object As Object)
 Dim TrayData As NOTIFYICONDATA
 
 TrayData.uID = vbNull
 TrayData.cbSize = Len(TrayData)
 TrayData.hwnd = TRAY_Object.hwnd
 TrayData.uFlags = 0
 
 dReturn = Shell_NotifyIcon(NIM_DELETE, TrayData)
End Sub

