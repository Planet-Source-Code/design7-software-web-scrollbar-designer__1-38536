VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmPreview 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scrollbar Preview"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton close 
      Caption         =   "Close"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   4335
   End
   Begin SHDocVwCtl.WebBrowser preview 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Load the HTML file that contains the scrollbar design code
preview.Navigate App.Path & "\scpreview.htm"
End Sub
