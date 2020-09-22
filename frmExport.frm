VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Scrollbar Design"
   ClientHeight    =   2175
   ClientLeft      =   3885
   ClientTop       =   3165
   ClientWidth     =   4380
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4380
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   3000
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton close 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton export 
      Caption         =   "Export"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton saveas 
      Caption         =   "Save As"
      Height          =   280
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox exportfilename 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox filetype 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Choose the type of exported Scrollbar Design file"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
Unload Me
End Sub

Private Sub export_Click()
Dim fileoutput As String
Dim f As Object
Dim exportfiles(0 To 1) As String
exportfiles(0) = ".css"
exportfiles(1) = ".html"
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(exportfilename & exportfiles(filetype.ListIndex), 2, True)
Dim htmlcode As String
htmlcode = frmMain.GenerateHTMLCode
Select Case filetype.ListIndex
Case 0:
fileoutput = htmlcode
Case 1:
fileoutput = "<html>" & vbCrLf & "<head>" & vbCrLf & _
"<style type=""text/css"">" & vbCrLf & "<!--" & vbCrLf & _
htmlcode & vbCrLf & "-->" & vbCrLf & "</style>" & _
vbCrLf & "<title>My Page</title>" & vbCrLf & "</head>" & _
vbCrLf & "<body>" & vbCrLf & "</body>" & vbCrLf & "</html>"
End Select
f.Write fileoutput
f.close
Unload Me
End Sub

Private Sub Form_Load()
filetype.AddItem "Cascading Style Sheet (*.css)"
filetype.AddItem "HTML File (*.html)"
filetype.ListIndex = 0
End Sub

Private Sub saveas_Click()
Dim exportfiles(0 To 1) As String
exportfiles(0) = ".css"
exportfiles(1) = ".html"
Select Case filetype.ListIndex
Case 0:
cmdialog.Filter = "Cascading Style Sheet (*.css)"
Case 1:
cmdialog.Filter = "HTML File (*.html)"
End Select
cmdialog.ShowSave
If cmdialog.filename <> "" Then
exportfilename.Text = cmdialog.filename & exportfiles(filetype.ListIndex)
End If
cmdialog.filename = ""
End Sub
