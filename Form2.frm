VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Layers"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   4680
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Layers"
      Height          =   5655
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command3 
         Caption         =   "Launch"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Launch Background Paint"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   5760
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Shell "C:/WINDOWS/SYSTEM32/MSPaint.exe"
End Sub

Private Sub Command4_Click()
MsgBox "Save both of the pictures"
End Sub
