VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form Form5 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   810
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   3870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   6826
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "B"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "I"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Btn"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "TxT"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "H1"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Font"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Style"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   1800
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Me.width = Me.width
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim bla As String
Select Case Button.Index
Case 1
Form4.DHTMLEdit1.ExecCommand (DECMD_BOLD)
Case 2
Form4.DHTMLEdit1.ExecCommand (DECMD_ITALIC)
Case 3
Form4.DHTMLEdit1.DocumentHTML = Form4.DHTMLEdit1.DocumentHTML & "<input type=submit value=""Button""</input>"
Case 4
Form4.DHTMLEdit1.DocumentHTML = Form4.DHTMLEdit1.DocumentHTML & "<input type=input value=""TextBox""</input>"
Case 5

bla = InputBox("Header size(only number)", "Vintagers")
Form4.DHTMLEdit1.DocumentHTML = "<h" & bla & ">" & "Header" & "</h" & bla & ">" & Form4.DHTMLEdit1.DocumentHTML
Case 6
bla = InputBox("Font Name", "Vintagers")
Form4.DHTMLEdit1.ExecCommand (DECMD_SETFONTNAME), , bla
Case 7
bla = InputBox("0(Help and Support theme) or 1(XP Help for programs)")
If bla = "1" Then

Else
Form4.DHTMLEdit1.DOM.bgColor = "#6375D6"
Form4.DHTMLEdit1.DOM.fgColor = "#FFFFFF"

Form4.DHTMLEdit1.DocumentHTML = Replace(Replace(Replace(Replace(Form4.DHTMLEdit1.DocumentHTML, "<P", "<font face=Tahoma color=#FFFFFF size=1><P"), "</P>", "</P></FONT>"), "<H1", "<FONT color=#D6DFF5 size=5 face=""Franklin Gothic Medium""><P"), "</H1>", "</P></FONT>")

End If
End Select

End Sub
