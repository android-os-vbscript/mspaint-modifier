VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DHTMLEDLibCtl.DHTMLEdit DHTMLEdit1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
DHTMLEdit1.width = Form1.width
DHTMLEdit1.height = Form1.height
End Sub
