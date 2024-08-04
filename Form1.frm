VERSION 5.00
Object = "{F1E2AAA6-E1B5-4648-9ED5-180ECE147792}#1.0#0"; "HookControl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSPaint Add-In System"
   ClientHeight    =   7320
   ClientLeft      =   8685
   ClientTop       =   4410
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   4440
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   4920
   End
   Begin HookCtrl.HookControl HookControl 
      Left            =   3120
      Top             =   3360
      _ExtentX        =   1058
      _ExtentY        =   582
   End
   Begin VB.TextBox Text1 
      Height          =   2235
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4320
      Width           =   3975
   End
   Begin VB.CommandButton cmdUnhook 
      Caption         =   "UnHook"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   1155
   End
   Begin VB.CommandButton cmdHook 
      Caption         =   "Hook Add-In"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2160
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   2895
      Left            =   960
      ScaleHeight     =   2895
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vintagers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   765
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unhooked"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vintagers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.Menu add 
      Caption         =   "addins"
      Visible         =   0   'False
      Begin VB.Menu dd 
         Caption         =   "DialogBox Draw"
      End
      Begin VB.Menu layers 
         Caption         =   "Word Paint"
      End
      Begin VB.Menu olepaint 
         Caption         =   "OLE Paint"
      End
      Begin VB.Menu htmlp 
         Caption         =   "HTML Paint"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hsubmenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Dim Msg As String
Const ParentCaption = "untitled - Paint"
Const ChildClassName = "Afx:1000000:8"
Const ChildCaption = ""
Const MF_BYCOMMAND As Long = &H0
Const MF_ENABLED As Long = &H0

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_SETTEXT = &HC
Private Const WM_CHAR = &H102
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_COMMAND = &H111
Private Const WM_NULL = 0

Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME

Private Const WS_CHILD = &H40000000

Private Const WS_VISIBLE = &H10000000
Private Const SS_LEFT = &H0
Private Const STATIC_CLASS = "Static"

Private Const WS_VSCROLL = &H200000
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Const GWL_EXSTYLE = -20
Const WS_EX_LAYERED = &H80000
Const LWA_ALPHA = &H2
 Const WM_SETBKCOLOR = &HC

Private Const GWL_STYLE = -16
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Dim hWndButton As Long
Private Const BS_PUSHBUTTON As Long = &H0
Private Const WM_PAINT = &HF
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Const GCL_HBRBACKGROUND = (-10)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long


























Private Sub cmdHook_Click()

'This routine sets the hook to monitor specified messages

With HookControl
    'handle to main window
    lhWnd = .GetTopLevelHandle(ParentCaption)
       
    
    While lhWnd = 0
        If MsgBox("Open Notepad", vbOKCancel) = vbCancel Then Exit Sub
        lhWnd = .GetTopLevelHandle(ParentCaption)
    Wend
    'monitor the following message(s)
    .AddMessage WM_CHAR, "WM_CHAR"
    .AddMessage &H111, "WM_COMMAND"
     .AddMessage WM_PAINT, "WM_PAINT"
    'handle to the textbox we want to hook

    
    .TargethWnd = lhWnd
    'Set the hooks
    If .SetHook Then
        cmdHook.Enabled = False
        cmdUnhook.Enabled = True
        lblStatus = "Hooked"
    Else
        cmdUnhook.Enabled = False
        cmdHook.Enabled = True
        lblStatus = "Unhooked"
    End If
End With


'     hWndButton = CreateWindowEx(0, "BUTTON", "My Button", WS_CHILD Or WS_VISIBLE Or BS_PUSHBUTTON, 1000, 50, 100, 30, HookControl.TargethWnd, 1000, App.hInstance, ByVal 0&)
   
'BringWindowToTop (hWndButton)
'SetParent Me.hwnd, lhWnd
'ChildWindowFromPoint(lhWnd, 480, 480)
 '   Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
  '  Dim hFont As Long

    ' Get a handle to the default GUI font
   ' hFont = GetStockObject(17)

    ' Set the button's font to the default GUI font
    'SendMessage hWndButton, WM_SETFONT, hFont, True
Dim hMenu As Long
Dim hsubmenu As Long
Dim hsubsubmenu As Long

' Get the menu
hMenu = GetMenu(lhWnd)

' Get the submenu
hsubmenu = GetSubMenu(hMenu, 5)

' Check if the submenu exists
        Call ModifyMenu(hsubmenu, 1, &H400& Or &H0&, 0, "Add-Ins")

End Sub

Public Sub cmdUnHook_Click()

'This routine removes the hook
If HookControl.RemoveAllHooks Then
    cmdUnhook.Enabled = False
    cmdHook.Enabled = True
    lblStatus = "Unhooked"
Else
    cmdHook.Enabled = False
    cmdUnhook.Enabled = True
    lblStatus = "Hooked"
End If

End Sub


Private Sub dd_Click()
Dim bla As RECT
Call GetWindowRect(lhWnd, bla)

 ShowWindow (ChildWindowFromPoint(lhWnd, 20, 50)), 0
                ShowWindow (ChildWindowFromPoint(lhWnd, 60, 50)), 0
 ShowWindow (ChildWindowFromPoint(lhWnd, 150, bla.Bottom - bla.Top - 35)), 0
Dim style As Long
    ' Update the window's position to apply the new style

        Call MoveWindow(lhWnd, bla.Left, bla.Top, bla.Right - bla.Left, 1, True)

    
  ' Define background color (vbCyan)
'Picture1.Top = 0
 ' Picture1.Left = 0
 ' Picture1.width = Screen.width
 '  Picture1.height = Screen.height

'SetParent Picture1.hwnd, lhWnd
 ' SendMessage lhWnd, WM_SETBKCOLOR, 1, vbCyan

  
    
    
     hWndButton = CreateWindowEx(0, "BUTTON", "Insert everything inside", WS_CHILD Or WS_VISIBLE Or BS_PUSHBUTTON, 500, 50, 150, 30, HookControl.TargethWnd, 1000, App.hInstance, ByVal 0&)
   
BringWindowToTop (hWndButton)
'SetParent Me.hwnd, lhWnd
'ChildWindowFromPoint(lhWnd, 480, 480)
    'Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Dim hFont As Long

    ' Get a handle to the default GUI font
    hFont = GetStockObject(17)

    ' Set the button's font to the default GUI font
    SendMessage hWndButton, WM_SETFONT, hFont, True
    
'      Call SetWindowPos(Picture1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
  ' Create the static control window
  'Dim hwndControl As Long
  'hwndControl = CreateWindowEx(0&, STATIC_CLASS, "", WS_CHILD Or WS_VISIBLE Or SS_LEFT, 0, 0, 1024, 768, lhwnd, 0&, App.hInstance, vbCyan)

  ' Error handling (optional)
  'If hwndControl = 0 Then
  '  MsgBox "Error creating static control!", vbExclamation
  'End If
' Example usage (assuming you have a form with a handle in hForm)
    'Dim exstyle As Long
    ' Get the extended window style
    'exstyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    ' Add the layered window style
   ' SetWindowLong lhWnd, GWL_EXSTYLE, exstyle Or WS_EX_LAYERED
    ' Set the color key to the specified color
   ' SetLayeredWindowAttributes lhWnd, Color, 0, &H1
   
   
'cmdUnHook_Click
'Call SetParent(mainMenu.hwnd, lhWnd)
'Call SetParent(dialogTemplate.hwnd, lhWnd)
'Call SetParent(Properties.hwnd, lhWnd)
'Call SetParent(ToolBox.hwnd, lhWnd)
frmSplash.Show
End Sub

Private Sub HookControl_SentMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
If uMsg = &HF Then
Text1.Text = Text1.Text & "paint"
End If
If uMsg = 273 Then
Call SetParent(mainMenu.hwnd, lhWnd)
Call SetParent(dialogTemplate.hwnd, lhWnd)
Call SetParent(Properties.hwnd, lhWnd)
Call SetParent(ToolBox.hwnd, lhWnd)
End If

Exit Sub
End Sub

Private Sub HookControl_UnHook()
  
cmdHook.Enabled = True
cmdUnhook.Enabled = False
lblStatus = "Unhooked"

End Sub
Private Sub HookControl_PostedMessage(uMsg As Long, wParam As Long, lParam As Long)

If uMsg = &HF Then
  Dim hdc As Long
    Dim hBrush As Long
    Dim RECT As RECT
    
    ' Get the device context
    hdc = GetDC(lhWnd)
    
    ' Set the RECT to the size of the window's client area
    ' (You'll need to fill in these values)
    RECT.Left = 0
    RECT.Top = 0
    RECT.Right = 1366  ' Replace with the width of the window
    RECT.Bottom = 768 ' Replace with the height of the window
    
    ' Create a solid brush
    hBrush = CreateSolidBrush(GetSysColor(15))    ' Red
    
    ' Fill the window with the brush
    FillRect hdc, RECT, hBrush
    
    ' Clean up
    DeleteObject hBrush
    ReleaseDC lhWnd, hdc
    
Else
'display the messages as ,

    If wParam = 0 Then
     Dim pt As POINTAPI

    ' Get the current mouse position in screen coordinates
    GetCursorPos pt

    ' Convert the screen coordinates to client (form) coordinates
    ScreenToClient Me.hwnd, pt
    
    Call Me.PopupMenu(add, , pt.x * Screen.TwipsPerPixelX, pt.Y * Screen.TwipsPerPixelY)
        End If
        Text1.Text = Text1.Text & wParam & vbNewLine
        End If
    ' an example of how to change a message
'Change all a's  to ‘X’
'If uMsg = WM_CHAR And wParam = Asc("a") Then wParam = Asc("X")
   

'change message to WM_NULL if key is "s" so Notepad ignores it
'If uMsg = WM_CHAR And wParam = Asc("s") Then uMsg = WM_NULL


End Sub

Private Sub htmlp_Click()
Dim bla As RECT
Call GetWindowRect(lhWnd, bla)

 ShowWindow (ChildWindowFromPoint(lhWnd, 20, 50)), 0
                ShowWindow (ChildWindowFromPoint(lhWnd, 60, 50)), 0
 ShowWindow (ChildWindowFromPoint(lhWnd, 150, bla.Bottom - bla.Top - 35)), 0
Form4.Show
Form4.width = Screen.width
Form4.height = Screen.height
Form4.Left = Form5.width
Form4.Top = 0
SetParent Form4.hwnd, lhWnd
Form5.Show
Form5.Top = 0
Form5.Left = 0
Form5.height = Screen.height
SetParent Form5.hwnd, lhWnd

Call SetWindowPos(Form5.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub layers_Click()
'Form2.Show vbModal, Me
Dim bla As RECT
Call GetWindowRect(lhWnd, bla)

 ShowWindow (ChildWindowFromPoint(lhWnd, 20, 50)), 0
                ShowWindow (ChildWindowFromPoint(lhWnd, 60, 50)), 0
 ShowWindow (ChildWindowFromPoint(lhWnd, 150, bla.Bottom - bla.Top - 35)), 0
 Form3.Tag = "word"
Form3.Show
Form3.width = Screen.width
Form3.height = Screen.height
Form3.Left = 0
Form3.Top = 0
SetParent Form3.hwnd, lhWnd

End Sub

Private Sub MimeEdit1_GotFocus()

End Sub

Private Sub olepaint_Click()
cmdUnHook_Click
Dim bla As RECT
Call GetWindowRect(lhWnd, bla)

 ShowWindow (ChildWindowFromPoint(lhWnd, 20, 50)), 0
                ShowWindow (ChildWindowFromPoint(lhWnd, 60, 50)), 0
 ShowWindow (ChildWindowFromPoint(lhWnd, 150, bla.Bottom - bla.Top - 35)), 0
 

 Form3.Show
Form3.width = Screen.width
Form3.height = Screen.height
Form3.Left = 0
Form3.Top = 0
SetParent Form3.hwnd, lhWnd

End Sub

Private Sub Timer1_Timer()
Me.Refresh
End Sub

Private Sub WindowsMediaPlayer1_OpenStateChange(ByVal NewState As Long)

End Sub
