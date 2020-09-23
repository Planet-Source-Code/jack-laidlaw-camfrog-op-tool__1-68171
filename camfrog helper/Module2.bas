Attribute VB_Name = "Module2"
Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type

Public Const NIM_ADD = &H0 'Add to Tray
Public Const NIM_MODIFY = &H1 'Modify Details
Public Const NIM_DELETE = &H2 'Remove From Tray
Public Const NIF_MESSAGE = &H1 'Message
Public Const NIF_ICON = &H2 'Icon
Public Const NIF_TIP = &H4 'TooTipText
Public Const WM_MOUSEMOVE = &H200 'On Mousemove
Public Const WM_LBUTTONDOWN = &H201 'Left Button Down
Public Const WM_LBUTTONUP = &H202 'Left Button Up
Public Const WM_LBUTTONDBLCLK = &H203 'Left Double Click
Public Const WM_RBUTTONDOWN = &H204 'Right Button Down
Public Const WM_RBUTTONUP = &H205 'Right Button Up
Public Const WM_RBUTTONDBLCLK = &H206 'Right Double Click

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
      
Public nid As NOTIFYICONDATA

