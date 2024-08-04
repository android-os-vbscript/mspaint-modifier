Attribute VB_Name = "Module2"
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    ' Destroy the child window
    DestroyWindow (hwnd)
    
    EnumChildProc = 1 'Continue enumeration
    MsgBox "del " & hwnd
End Function

Sub DeleteChildWindows(Whwnd As Long)
    'EnumChildWindows Whwnd, AddressOf EnumChildProc, ByVal 0&
End Sub

