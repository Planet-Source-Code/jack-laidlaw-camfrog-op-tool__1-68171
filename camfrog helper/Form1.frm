VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Camfrog Helper"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Unblock"
      Height          =   255
      Left            =   1080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Block Mic"
      Height          =   255
      Left            =   120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Always On Top"
      Height          =   495
      Left            =   1680
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Send To Tray"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Warn"
      Height          =   255
      Left            =   120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":002B
      TabIndex        =   6
      Text            =   "With Reason"
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Banip"
      Height          =   255
      Left            =   1080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Ban"
      Height          =   255
      Left            =   120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Punish"
      Height          =   255
      Left            =   2040
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Kick"
      Height          =   255
      Left            =   1080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1200
      Top             =   2280
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Status: "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   540
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageLONG Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Const WM_SETTEXT = &HC
Private Const BM_CLICK = &HF5

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Sub About_Click()
frmAbout.Show vbModal
End Sub


Private Sub Command1_Click()

Dim OurParent As Long
Dim OurHandle As Long
Dim OurButton As Long
Dim sCaption As String
Dim kick As String
Dim Reason As String
    
'This Finds any window with "video Chat Room" in the title
OurParent = FindWindowWild("*Video Chat Room*", False)
'This statement is used for when camfrog rooms have topic set
'as that would make the wildcard above void
OurParent = FindWindowWild("*Topic*", False)
'This finds the message box inside the camfrog room
OurHandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
'This finds the send message button inside the camfrog room
OurButton& = FindWindowEx(OurParent&, 0, "Button", vbNullString)

kick = "/kick "
'This tells the program that what ever is typed in the textbox
'that is what sCaption is equal to
sCaption = Text1 + " "
Reason = Combo1
    
'This sends the text in the text box to the text box in camfrog
SendMessageSTRING OurHandle, WM_SETTEXT, 256, kick + sCaption + Reason
'This click's the "send" button in the camfrog room
SendMessageSTRING OurButton, BM_CLICK, 0, 0

'moves curser back into textbox
Text1.SetFocus
End Sub

Private Sub Command2_Click()

Dim OurParent As Long
Dim OurHandle As Long
Dim OurButton As Long
Dim sCaption As String
Dim Punish As String
Dim Reason As String
    
'This Finds any window with "video Chat Room" in the title
OurParent = FindWindowWild("*Video Chat Room*", False)
'This statement is used for when camfrog rooms have topic set
'as that would make the wildcard above void
OurParent = FindWindowWild("*Topic*", False)
'This finds the message box inside the camfrog room
OurHandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
'This finds the send message button inside the camfrog room
OurButton& = FindWindowEx(OurParent&, 0, "Button", vbNullString)

Punish = "/punish "
'This tells the program that what ever is typed in the textbox
'that is what sCaption is equal to
sCaption = Text1 + " "
Reason = Combo1
'This sends the text in the text box to the text box in camfrog
SendMessageSTRING OurHandle, WM_SETTEXT, 256, Punish + sCaption + Reason
'This click's the "send" button in the camfrog room
SendMessageSTRING OurButton, BM_CLICK, 0, 0

'moves curser back into textbox
Text1.SetFocus

End Sub

Private Sub Command3_Click()

Dim OurParent As Long
Dim OurHandle As Long
Dim OurButton As Long
Dim sCaption As String
Dim Ban As String
Dim Reason As String
    
'This Finds any window with "video Chat Room" in the title
OurParent = FindWindowWild("*Video Chat Room*", False)
'This statement is used for when camfrog rooms have topic set
'as that would make the wildcard above void
OurParent = FindWindowWild("*Topic*", False)
'This finds the message box inside the camfrog room
OurHandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
'This finds the send message button inside the camfrog room
OurButton& = FindWindowEx(OurParent&, 0, "Button", vbNullString)

Ban = "/ban "
'This tells the program that what ever is typed in the textbox
'that is what sCaption is equal to
sCaption = Text1 + " "
Reason = Combo1

'This sends the text in the text box to the text box in camfrog
SendMessageSTRING OurHandle, WM_SETTEXT, 256, Ban + sCaption + Reason
'This click's the "send" button in the camfrog room
SendMessageSTRING OurButton, BM_CLICK, 0, 0

'moves curser back into textbox
Text1.SetFocus

End Sub

Private Sub Command4_Click()

Dim OurParent As Long
Dim OurHandle As Long
Dim OurButton As Long
Dim sCaption As String
Dim Banip As String
Dim Reason As String
    
'This Finds any window with "video Chat Room" in the title
OurParent = FindWindowWild("*Video Chat Room*", False)
'This statement is used for when camfrog rooms have topic set
'as that would make the wildcard above void
OurParent = FindWindowWild("*Topic*", False)
'This finds the message box inside the camfrog room
OurHandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
'This finds the send message button inside the camfrog room
OurButton& = FindWindowEx(OurParent&, 0, "Button", vbNullString)

Banip = "/banip "
'This tells the program that what ever is typed in the textbox
'that is what sCaption is equal to
sCaption = Text1 + " "
Reason = Combo1
    
'This sends the text in the text box to the text box in camfrog
SendMessageSTRING OurHandle, WM_SETTEXT, 256, Banip + sCaption + Reason
'This click's the "send" button in the camfrog room
SendMessageSTRING OurButton, BM_CLICK, 0, 0

'moves curser back into textbox
Text1.SetFocus

End Sub

Private Sub Command5_Click()

Dim OurParent As Long
Dim OurHandle As Long
Dim OurButton As Long
Dim sCaption As String
Dim Warn As String
Dim Reason As String
    
'This Finds any window with "video Chat Room" in the title
OurParent = FindWindowWild("*Video Chat Room*", False)
'This statement is used for when camfrog rooms have topic set
'as that would make the wildcard above void
OurParent = FindWindowWild("*Topic*", False)
'This finds the message box inside the camfrog room
OurHandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
'This finds the send message button inside the camfrog room
OurButton& = FindWindowEx(OurParent&, 0, "Button", vbNullString)

'This tells the program that what ever is typed in the textbox
'that is what sCaption is equal to
sCaption = Text1 + " "
Reason = Combo1
Warn = "(ST) " + sCaption + "(ST) "
    
'This sends the text in the text box to the text box in camfrog
SendMessageSTRING OurHandle, WM_SETTEXT, 256, Warn + Reason
'This click's the "send" button in the camfrog room
SendMessageSTRING OurButton, BM_CLICK, 0, 0

'moves curser back into textbox
Text1.SetFocus

End Sub

Private Sub Command6_Click()
Me.Hide
End Sub

Private Sub Command7_Click()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Command8_Click()
Dim OurParent As Long
Dim OurHandle As Long
Dim OurButton As Long
Dim Block As String
Dim sCaption As String

'This Finds any window with "video Chat Room" in the title
OurParent = FindWindowWild("*Video Chat Room*", False)
'This statement is used for when camfrog rooms have topic set
'as that would make the wildcard above void
OurParent = FindWindowWild("*Topic*", False)
'This finds the message box inside the camfrog room
OurHandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
'This finds the send message button inside the camfrog room
OurButton& = FindWindowEx(OurParent&, 0, "Button", vbNullString)

Block = "/blockmic "
sCaption = Text1

'This sends the text in the text box to the text box in camfrog
SendMessageSTRING OurHandle, WM_SETTEXT, 256, Block + sCaption
'This click's the "send" button in the camfrog room
SendMessageSTRING OurButton, BM_CLICK, 0, 0

'moves curser back into textbox
Text1.SetFocus


End Sub

Private Sub Command9_Click()
Dim OurParent As Long
Dim OurHandle As Long
Dim OurButton As Long
Dim Unblock As String
Dim sCaption As String

'This Finds any window with "video Chat Room" in the title
OurParent = FindWindowWild("*Video Chat Room*", False)
'This statement is used for when camfrog rooms have topic set
'as that would make the wildcard above void
OurParent = FindWindowWild("*Topic*", False)
'This finds the message box inside the camfrog room
OurHandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
'This finds the send message button inside the camfrog room
OurButton& = FindWindowEx(OurParent&, 0, "Button", vbNullString)

Unblock = "/unblockmic "
sCaption = Text1

'This sends the text in the text box to the text box in camfrog
SendMessageSTRING OurHandle, WM_SETTEXT, 256, Unblock + sCaption
'This click's the "send" button in the camfrog room
SendMessageSTRING OurButton, BM_CLICK, 0, 0

'moves curser back into textbox
Text1.SetFocus

End Sub

Private Sub Exit_Click()
End
End Sub


Private Sub Form_Load()
  Me.Show 'form must be fully visible
    Me.Refresh
        
        With nid 'with system tray
            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon 'use form's icon in tray
            .szTip = "System Tray Example" & vbNullChar 'tooltip text
        End With
        
    Shell_NotifyIcon NIM_ADD, nid 'add to tray
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Result, Action As Long
    
    'there are two display modes and we need to find out
    'which one the application is using
    
    If Me.ScaleMode = vbPixels Then
        Action = x
    Else
        Action = x / Screen.TwipsPerPixelX
    End If
    
Select Case Action

    Case WM_LBUTTONDBLCLK 'Left Button Double Click
        Me.WindowState = vbNormal 'put into taskbar
            Result = SetForegroundWindow(Me.hwnd)
        Me.Show 'show form
    
    Case WM_RBUTTONUP 'Right Button Up
        Result = SetForegroundWindow(Me.hwnd)
        PopupMenu File 'popup menu, cool eh?
    
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer) 'on form unload
    Shell_NotifyIcon NIM_DELETE, nid 'remove from tray
End Sub

Private Sub mnuExit_Click() 'exit
    Unload Me: End
End Sub

Private Sub Timer1_Timer()
Dim OurParent As Long
Dim sCaption As String * 256

'This Finds any window with "video Chat Room" in the title
OurParent = FindWindowWild("*Video Chat Room*", False)
'This statement is used for when camfrog rooms have topic set
'as that would make the wildcard above void
OurParent = FindWindowWild("*Topic*", False)

GetWindowText OurParent, sCaption, 256

If OurParent Then
    lblStatus.Caption = sCaption
    Else
    lblStatus.Caption = "Not in room"
    End If

End Sub
