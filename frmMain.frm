VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mainMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "MSPaint Dialog Editor"
   ClientHeight    =   405
   ClientLeft      =   2640
   ClientTop       =   2850
   ClientWidth     =   4260
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   284
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   -75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":049C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   6120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":06C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":07D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":08E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_NewDialog 
         Caption         =   "&New Dialog"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_File_OpenDialog 
         Caption         =   "&Open Dialog"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_File_Save 
         Caption         =   "&Save Dialog"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_File_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Op_ShowGrid 
         Caption         =   "&Show Grid"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_Op_Snap 
         Caption         =   "S&nap to Grid"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "mainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDialogBaseUnits Lib "user32" () As Long

Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function MapDialogRect Lib "user32" (ByVal hDlg As Long, lpRect As rect) As Long

Private Function PixelToDialogUnits(ByVal pixelX As Long, ByVal pixelY As Long) As rect
   
        Dim baseUnits As Long
    baseUnits = GetDialogBaseUnits()

    Dim baseunitX As Long
    Dim baseunitY As Long
    baseunitX = baseUnits \ &H10000  ' low-order word
    'baseunitY = baseUnits And &HFFFF& ' high-order word

    Dim dialogunitX As Long
    Dim dialogunitY As Long
    dialogunitX = pixelX * 8 / baseunitX
    dialogunitY = pixelY * 8 / baseunitX
    Dim dialogrect As rect
    dialogrect.Left = 0
    dialogrect.Top = 0
    dialogrect.Right = dialogunitX
    dialogrect.Bottom = dialogunitY
PixelToDialogUnits = dialogrect
    ' Now pixelX and pixelY contain the size in pixels
End Function
Private Sub Form_Load()
'    SetParent Me.hwnd, lhWnd
 '       SetParent Properties.hwnd, lhWnd
  '          SetParent dialogTemplate.hwnd, lhWnd
   '             SetParent ToolBox.hwnd, lhWnd

 'Dim style As Long

    ' Get the current window style
  '  style = GetWindowLong(hwnd, GWL_STYLE)

    ' Remove the caption bar and border from the style
   ' style = style And Not WS_CAPTION

    ' Set the new window style
    'Call SetWindowLong(hwnd, GWL_STYLE, style)

    ' Update the window's position to apply the new style
    'Call SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
   
'Me.width = Screen.width

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload ToolBox
    Unload Me
    End
End Sub

Private Sub Form_Resize()
    ToolBox.Show
End Sub


Private Sub mnu_File_Exit_Click()
    Unload Me
End Sub

Private Sub mnu_File_NewDialog_Click()
    Dim newDialog As New dialogTemplate
    newDialog.Show
    'SetParent newDialog.hwnd, lhWnd
End Sub
Private Function insertctrl(ByVal ctrl As Control) As String
  Dim dlgBaseUnits As Long
    Dim XBaseUnits As Integer
    Dim YBaseUnits As Integer
    Dim XUnitPixels As Single
    Dim YUnitPixels As Single

    ' Get the dialog base units
    dlgBaseUnits = GetDialogBaseUnits()

    ' Extract the X and Y base units
    XBaseUnits = dlgBaseUnits And &HFFFF&
    YBaseUnits = dlgBaseUnits \ &H10000

    ' Calculate the width of one dialog unit in pixels (X axis)
    XUnitPixels = XBaseUnits / 4 ' typical conversion factor for X-axis

    ' Calculate the height of one dialog unit in pixels (Y axis)
    YUnitPixels = YBaseUnits / 8 ' typical conversion factor for Y-axis


 If TypeOf ctrl Is CommandButton Then
        insertctrl = " CONTROL """ & ctrl.Caption & "" & ", " & ctrl.Index & " , BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, " & ctrl.Left / XUnitPixels & ", " & ctrl.Top / YUnitPixels & ", " & ctrl.width / XUnitPixels & ", " & ctrl.height / YUnitPixels
    ElseIf TypeOf ctrl Is TextBox Then
    insertctrl = " CONTROL """ & ctrl.Text & "" & ", " & ctrl.Index & " , EDIT, WS_BORDER | WS_CHILD | WS_VISIBLE | WS_TABSTOP, " & ctrl.Left / XUnitPixels & ", " & ctrl.Top / YUnitPixels & ", " & ctrl.width / XUnitPixels & ", " & ctrl.height / YUnitPixels
    ElseIf TypeOf ctrl Is Label Then
        insertctrl = " CONTROL """ & ctrl.Caption & "" & ", " & ctrl.Index & " , STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, " & ctrl.Left / XUnitPixels & ", " & ctrl.Top / YUnitPixels & ", " & ctrl.width / XUnitPixels & ", " & ctrl.height / YUnitPixels
    ElseIf TypeOf ctrl Is Frame Then
        Debug.Print ctrl.Name & " is a Frame."
    Else
        ' Handle other control types if necessary
    End If
End Function


Private Sub mnu_File_Save_Click()
Dim ctrl As String
Dim ctrlo As Control

For Each ctrlo In selectedForm.Controls
    ' Perform actions on each control
    ' For example:
    ' Debug.Print ctrl.Name
    If ctrlo.Visible = True Then
    ctrl = ctrl & insertctrl(ctrlo) & vbNewLine
    End If
Next ctrlo

Dim formrect As rect
  Dim dlgBaseUnits As Long
    Dim XBaseUnits As Integer
    Dim YBaseUnits As Integer
    Dim XUnitPixels As Single
    Dim YUnitPixels As Single
Dim style As String
    ' Get the dialog base units
    dlgBaseUnits = GetDialogBaseUnits()

    ' Extract the X and Y base units
    XBaseUnits = dlgBaseUnits And &HFFFF&
    YBaseUnits = dlgBaseUnits \ &H10000

    ' Calculate the width of one dialog unit in pixels (X axis)
    XUnitPixels = XBaseUnits / 4 ' typical conversion factor for X-axis

    ' Calculate the height of one dialog unit in pixels (Y axis)
    YUnitPixels = YBaseUnits / 8 ' typical conversion factor for Y-axis
Dim resp As Integer
Dim rc As String
If selectedForm.BorderStyle = 2 Or selectedForm.BorderStyle = 1 Then
style = style & " | " & "WS_MAXIMIZEBOX | WS_MINIMIZEBOX"
End If
If Not selectedForm.BorderStyle = 0 Then
style = style & " | " & "WS_CAPTION"
End If
If selectedForm.BorderStyle = 2 Or selectedForm.BorderStyle = 5 Then
style = style & " | WS_THICKFRAME"
End If
If selectedForm.BorderStyle = 4 Or selectedForm.BorderStyle = 5 Then
style = style & vbNewLine & "EXSTYLE WS_EX_PALETTEWINDOW"
End If
rc = "1 DIALOGEX 0, 0, " & (selectedForm.ScaleWidth) / XUnitPixels & ", " & (selectedForm.ScaleHeight) / YUnitPixels & vbNewLine & "STYLE DS_SETFONT | DS_MODALFRAME | WS_POPUP | WS_SYSMENU" & style & vbNewLine & "CAPTION """ & selectedForm.Caption & """" & vbNewLine & "LANGUAGE LANG_NEUTRAL, SUBLANG_NEUTRAL" & vbNewLine & "FONT 8, ""MS Sans Serif""" & vbNewLine & "{" & vbNewLine & ctrl & "}"
 resp = MsgBox(rc)
 If resp = vbOK Then
 Clipboard.Clear
 Clipboard.SetText (rc)
 End If
End Sub
Private Sub mnu_Op_ShowGrid_Click()
    With mnu_Op_ShowGrid
        .Checked = Not .Checked
        ShowGrid = .Checked
    End With
    
    Dim x As Form
    For Each x In Forms
        If x.Name = "dialogTemplate" Then DrawTheGrid x
    Next
End Sub


Private Sub mnu_Op_Snap_Click()
    With mnu_Op_Snap
        .Checked = Not .Checked
        useGrid = .Checked
    End With
End Sub


Private Sub toolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  
    Select Case Button.Index
        Case 1
            Call mnu_File_NewDialog_Click
            On Error Resume Next
            Case 2
            mnu_File_Save_Click
              End Select
            
End Sub
