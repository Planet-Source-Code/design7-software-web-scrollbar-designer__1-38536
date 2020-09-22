VERSION 5.00
Begin VB.Form frmCode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scrollbar HTML Code"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton close 
      Caption         =   "Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton copy 
      Caption         =   "Copy to the clipboard"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox thehtmlcode 
      Height          =   3375
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Insert the following code between the <HEAD> tags"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
Unload Me
End Sub

Private Sub copy_Click()
Clipboard.SetText thehtmlcode.Text
End Sub
