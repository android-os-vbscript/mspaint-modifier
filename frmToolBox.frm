VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form ToolBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Toolbox"
   ClientHeight    =   3780
   ClientLeft      =   2625
   ClientTop       =   3675
   ClientWidth     =   1155
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   77
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar toolBar 
      Align           =   1  'Align Top
      Height          =   1110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1958
      ButtonWidth     =   926
      ButtonHeight    =   900
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   0
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   2280
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolBox.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolBox.frx":0982
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolBox.frx":1304
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolBox.frx":1C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolBox.frx":2608
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Initialize()
Dim lol As Integer
lol = Me.width
Me.width = 1000
Me.width = lol
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.Tag = "" Then Cancel = 1
End Sub


Private Sub Form_Resize()
    toolBar.Move toolBar.Left, toolBar.Top, Me.ScaleWidth - 4, Me.ScaleHeight - 6
    'toolBar.Move 0, 0, Me.width, Me.width
End Sub

Private Sub Timer1_Timer()
Dim lol As Integer
lol = Me.width
Me.width = 100000
Me.width = lol

Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()

Me.Refresh
toolBar.Refresh
End Sub

Private Sub toolBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim x As Form
    For Each x In Forms
        If x.Name = "dialogTemplate" Then
            If Button.Index > 1 Then
                x.MousePointer = 2
            Else
                x.MousePointer = 0
            End If
        End If
    Next x
End Sub
