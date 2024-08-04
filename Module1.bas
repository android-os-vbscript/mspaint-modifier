Attribute VB_Name = "Module1"
 Option Explicit
 Public lhWnd As Long
 Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

 Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
 Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
 Const WM_COMMAND As Long = &H111
 Const GWL_WNDPROC As Long = -4
Public Const MF_STRING As Long = &H0&
 
 Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal uIDEnableItem As Long, ByVal uEnable As Long) As Long
Const MF_ENABLED As Long = &H0&

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

 Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long


Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Public Const WM_SETFONT = &H30
Const DEFAULT_GUI_FONT = 17


Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long

Const MF_BYPOSITION = &H400&

Public Const mf_popup = &H10

