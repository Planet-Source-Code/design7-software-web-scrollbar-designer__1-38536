VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Scrollbar Designer"
   ClientHeight    =   7305
   ClientLeft      =   1695
   ClientTop       =   915
   ClientWidth     =   8265
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8265
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   6600
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Scrollbar Designer File (*.scd) | *.scd"
   End
   Begin VB.CommandButton loadscd 
      Caption         =   "Load Scrollbars Design"
      Height          =   615
      Left            =   3240
      TabIndex        =   53
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton savescd 
      Caption         =   "Save Scrollbars Design"
      Height          =   615
      Left            =   5760
      TabIndex        =   52
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton previewscd 
      Caption         =   "Preview Scrollbars"
      Height          =   615
      Left            =   3240
      TabIndex        =   51
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton generatecode 
      Caption         =   "Generate Code"
      Height          =   615
      Left            =   5760
      TabIndex        =   50
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton choosecolor 
      Caption         =   "Choose Color"
      Height          =   255
      Index           =   6
      Left            =   6960
      TabIndex        =   49
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton choosecolor 
      Caption         =   "Choose Color"
      Height          =   255
      Index           =   5
      Left            =   6960
      TabIndex        =   48
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton choosecolor 
      Caption         =   "Choose Color"
      Height          =   255
      Index           =   4
      Left            =   6960
      TabIndex        =   47
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton choosecolor 
      Caption         =   "Choose Color"
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   46
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton choosecolor 
      Caption         =   "Choose Color"
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   45
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton choosecolor 
      Caption         =   "Choose Color"
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   44
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton choosecolor 
      Caption         =   "Choose Color"
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   43
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox scpcolor 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   5040
      TabIndex        =   41
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox scpcolor 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   5040
      TabIndex        =   40
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox scpcolor 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   39
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox scpcolor 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   5040
      TabIndex        =   38
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox scpcolor 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   5040
      TabIndex        =   37
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox scpcolor 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   36
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox scpcolor 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   35
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   9
         Left            =   240
         ScaleHeight     =   2775
         ScaleWidth      =   255
         TabIndex        =   22
         Top             =   2535
         Width           =   255
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   495
         ScaleHeight     =   255
         ScaleWidth      =   1695
         TabIndex        =   26
         Top             =   5055
         Width           =   1695
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   13
         Left            =   2175
         ScaleHeight     =   2775
         ScaleWidth      =   255
         TabIndex        =   23
         Top             =   2535
         Width           =   255
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2175
         TabIndex        =   21
         Top             =   5295
         Width           =   2175
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Index           =   12
         Left            =   720
         ScaleHeight     =   2295
         ScaleWidth      =   1455
         TabIndex        =   27
         Top             =   2760
         Width           =   1455
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Index           =   11
         Left            =   495
         ScaleHeight     =   2295
         ScaleWidth      =   255
         TabIndex        =   25
         Top             =   2775
         Width           =   255
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   495
         ScaleHeight     =   240
         ScaleWidth      =   1695
         TabIndex        =   24
         Top             =   2535
         Width           =   1695
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2175
         TabIndex        =   19
         Top             =   2280
         Width           =   2175
         Begin VB.PictureBox Picture23 
            Height          =   15
            Left            =   240
            ScaleHeight     =   15
            ScaleWidth      =   1095
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2175
         TabIndex        =   18
         Top             =   2040
         Width           =   2175
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   1
         Left            =   240
         ScaleHeight     =   1815
         ScaleWidth      =   255
         TabIndex        =   6
         Top             =   255
         Width           =   255
         Begin VB.PictureBox Picture7 
            Height          =   975
            Left            =   240
            ScaleHeight     =   975
            ScaleWidth      =   15
            TabIndex        =   9
            Top             =   480
            Width           =   15
         End
         Begin VB.PictureBox Picture15 
            Height          =   255
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   8
            Top             =   1560
            Width           =   15
         End
         Begin VB.PictureBox Picture17 
            Height          =   15
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   255
            TabIndex        =   7
            Top             =   1800
            Width           =   255
         End
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   3
         Left            =   480
         ScaleHeight     =   1095
         ScaleWidth      =   255
         TabIndex        =   11
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   495
         ScaleHeight     =   255
         ScaleWidth      =   1695
         TabIndex        =   10
         Top             =   495
         Width           =   1695
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   6
         Left            =   2415
         ScaleHeight     =   5295
         ScaleWidth      =   255
         TabIndex        =   1
         Top             =   255
         Width           =   255
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   495
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   2
         Top             =   255
         Width           =   1935
         Begin VB.PictureBox Picture4 
            Height          =   15
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   1335
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.PictureBox Picture5 
            Height          =   15
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   855
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin VB.PictureBox Picture13 
            Height          =   15
            Left            =   1680
            ScaleHeight     =   15
            ScaleWidth      =   255
            TabIndex        =   3
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   20
         Left            =   240
         ScaleHeight     =   1185
         ScaleWidth      =   2385
         TabIndex        =   42
         Top             =   5520
         Width           =   2415
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   495
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   17
         Top             =   1815
         Width           =   1935
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Index           =   5
         Left            =   2175
         ScaleHeight     =   1335
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   495
         Width           =   255
      End
      Begin VB.PictureBox scrollbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   4
         Left            =   735
         ScaleHeight     =   1095
         ScaleWidth      =   1455
         TabIndex        =   12
         Top             =   735
         Width           =   1455
         Begin VB.PictureBox scrollbar 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   16
            Left            =   600
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   15
            Top             =   360
            Width           =   135
         End
         Begin VB.PictureBox scrollbar 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   17
            Left            =   480
            ScaleHeight     =   135
            ScaleWidth      =   375
            TabIndex        =   14
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox scrollbar 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   18
            Left            =   360
            ScaleHeight     =   135
            ScaleWidth      =   615
            TabIndex        =   13
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Shape Shape1 
         Height          =   5310
         Left            =   230
         Top             =   240
         Width           =   2445
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":1272
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3240
      TabIndex        =   55
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scrollbar Designer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDD89F&
      Height          =   615
      Left            =   3240
      TabIndex        =   54
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label scpart 
      Caption         =   "Track Color:"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   34
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label scpart 
      Caption         =   "Dark shadow Color:"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   33
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label scpart 
      Caption         =   "Shadow Color:"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   32
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label scpart 
      Caption         =   "3D light Color:"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   31
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label scpart 
      Caption         =   "Highlight Color:"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   30
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label scpart 
      Caption         =   "Arrow Color:"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   29
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label scpart 
      Caption         =   "Face Color:"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   28
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   6885
      Left            =   105
      Top             =   230
      Width           =   2925
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save Design"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "Export Design"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "Preview Scrollbars"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileGenerateCode 
         Caption         =   "Generate Code"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuFileAbout 
      Caption         =   "About"
      Begin VB.Menu mnuFilePage 
         Caption         =   "Scrollbar Designer Page"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim scparts(0 To 20) As Integer
Dim scrollbarcolors(0 To 20) As Long
Dim scdesc(0 To 6) As String
Const scrollbarobjects = 6
Dim modified As Boolean
Dim fso As Object

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Sub HighlightLabel(Index As Integer)
For i = 0 To Index - 1
scpart(i).ForeColor = vbBlack
Next
For j = Index + 1 To scpart.Count - 1
scpart(j).ForeColor = vbBlack
Next
scpart(Index).ForeColor = vbRed
End Sub

Function RGBToHex(RGBColor)
'Function to convert a RGB color to hex
Dim RGBRed As String, RGBGreen As String, RGBBlue As String
RGBRed = Hex(RGBColor And &HFF)
RGBGreen = Hex(((RGBColor And &H100FF00) / &H100))
RGBBlue = Hex((RGBColor And &HFF0000) / &H10000)
RGBRed = String(2 - Len(RGBRed), "0") & RGBRed
RGBGreen = String(2 - Len(RGBGreen), "0") & RGBGreen
RGBBlue = String(2 - Len(RGBBlue), "0") & RGBBlue
RGBToHex = RGBRed & RGBGreen & RGBBlue
End Function

Private Sub choosecolor_Click(Index As Integer)
On Error GoTo exitme
cmdialog.ShowColor
With scpcolor(Index)
Dim hexcolor
hexcolor = RGBToHex(cmdialog.Color)
.Text = "#" & hexcolor
modified = True
End With
For i = 0 To UBound(scparts)
If scparts(i) = Index Then
scrollbar(i).BackColor = cmdialog.Color
End If
Next
exitme:
Exit Sub
End Sub

Function GenerateHTMLCode()
'Generate HTML code for the scrollbars
Dim thehtmlcode As String
thehtmlcode = "BODY{" & vbCrLf
For i = 0 To scrollbarobjects
thehtmlcode = thehtmlcode & "scrollbar-" & scdesc(i) & _
"-color: " & scpcolor(i).Text & ";" & vbCrLf
Next
thehtmlcode = thehtmlcode & "}"
GenerateHTMLCode = thehtmlcode
End Function

Private Sub Form_Load()
scparts(0) = 3
scparts(1) = 3
scparts(2) = 2
scparts(3) = 2
scparts(4) = 0
scparts(5) = 4
scparts(6) = 5
scparts(7) = 4
scparts(8) = 3
scparts(9) = 3
scparts(10) = 2
scparts(11) = 2
scparts(12) = 0
scparts(13) = 4
scparts(14) = 4
scparts(16) = 1
scparts(17) = 1
scparts(18) = 1
scparts(15) = 5
scparts(19) = 5
scparts(20) = 6

scdesc(0) = "face"
scdesc(1) = "arrow"
scdesc(2) = "highlight"
scdesc(3) = "3dlight"
scdesc(4) = "shadow"
scdesc(5) = "darkshadow"
scdesc(6) = "track"

'Load default colors
For i = 0 To UBound(scparts)
Dim hexcolor As String
hexcolor = RGBToHex(scrollbar(i).BackColor)
scpcolor(scparts(i)) = "#" & hexcolor
scrollbarcolors(i) = scrollbar(i).BackColor
Next

modified = False

Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To scpart.Count - 1
scpart(i).ForeColor = vbBlack
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Check if the scrollbar design was modified
If modified = True Then
If MsgBox("Do you want to save changes in the current scrollbar design?", vbInformation + vbYesNo, "Scrollbar Designer") = vbYes Then
savescd_Click
End If
End If
End Sub

Private Sub generatecode_Click()
frmCode.thehtmlcode.Text = GenerateHTMLCode()
frmCode.Show
End Sub

Private Sub loadscd_Click()
On Error GoTo exitme
cmdialog.DialogTitle = "Select a file to open"
cmdialog.ShowOpen
Dim f As Object
Dim hexcolor As String
Dim vbhexcolor As String
Set f = fso.OpenTextFile(cmdialog.FileName, 1, False)
cmdialog.FileName = ""
modified = False
'If file does not starts with the correct header then it is not valid
If f.Read(3) <> "scd" Then
MsgBox "The file is not a valid Scrollbar Designer file", vbExclamation, App.Title
Exit Sub
f.close
End If
For i = 0 To scrollbarobjects
'Read HTML hex color
hexcolor = f.Read(6)
'Read VB hex color
vbhexcolor = "&H" & f.Read(6)
scpcolor(scparts(i)) = "#" & hexcolor
For j = 0 To UBound(scparts)
If scparts(j) = i Then
scrollbar(j).BackColor = vbhexcolor
End If
Next
Next
f.close
exitme:
Exit Sub
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuFileExport_Click()
frmExport.Show
End Sub

Private Sub mnuFileGenerateCode_Click()
generatecode_Click
End Sub

Private Sub mnuFileNew_Click()
'Check if the scrollbar design was modified
If modified = True Then
If MsgBox("Do you want to save changes in the current scrollbar design?", vbInformation + vbYesNo, "Scrollbar Designer") = vbYes Then
savescd_Click
End If
End If
'Load default colors
For i = 0 To UBound(scparts)
Dim hexcolor As String
scrollbar(i).BackColor = scrollbarcolors(i)
hexcolor = RGBToHex(scrollbarcolors(i))
scpcolor(scparts(i)) = "#" & hexcolor
Next
modified = False
End Sub

Private Sub mnuFileOpen_Click()
loadscd_Click
End Sub

Private Sub mnuFilePage_Click()
ShellExecute &O0, "OPEN", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=38536&lngWId=1", "", "", vbNormal
End Sub

Private Sub mnuFilePreview_Click()
previewscd_Click
End Sub

Private Sub mnuFileSave_Click()
savescd_Click
End Sub

Private Sub savescd_Click()
On Error GoTo exitme
cmdialog.DialogTitle = "Save Scrollbar Design"
cmdialog.ShowSave
'Check if file already exists
If Dir(cmdialog.FileName) <> "" Then
If MsgBox("The file already exists, do you want to replace it?", vbExclamation + vbYesNo, "Scrollbar Designer") = vbYes Then
fso.DeleteFile cmdialog.FileName
Else
Exit Sub
End If
End If
Dim f As Object
Dim vbhexcolor As String
Set f = fso.OpenTextFile(cmdialog.FileName, 2, True)
cmdialog.FileName = ""
modified = False
'Write file header
f.Write "scd"
For i = 0 To scrollbarobjects
'Write hex colors
f.Write Right(scpcolor(i).Text, 6)
For j = 0 To UBound(scparts)
If scparts(j) = i Then
'Get VB hex color
vbhexcolor = Hex(scrollbar(j).BackColor)
'If the hex color contatains less than 6 chars then fill it with ceros
If Len(vbhexcolor) < 6 Then
vbhexcolor = String(6 - Len(vbhexcolor), "0") & vbhexcolor
End If
'Write VB hex color
f.Write vbhexcolor
Exit For
End If
Next
Next
f.close
exitme:
Exit Sub
End Sub

Private Sub previewscd_Click()
Dim f As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(App.Path & "\scpreview.htm", 2, True)
'Write HTML file containing the code
f.Write "<html>" & vbCrLf & "<head>" & vbCrLf & _
"<style type=""text/css"">" & vbCrLf & "<!--" & vbCrLf & _
GenerateHTMLCode() & vbCrLf & "-->" & vbCrLf & "</style>" & _
vbCrLf & "<title>Scrollbar" & "Design Preview</title>" & _
vbCrLf & "</head>" & vbCrLf & "<body><p>&nbsp;</p>" & _
"<p><b><font face=""Arial"" color=""#000080"" size=""3"">" & _
"Check it, this is your new scrollbar<br></font></b>" & _
"<font face=""Arial"" size=""1"" color=""#FF0000"">" & _
"Note: This preview can only be viewed<br>with Internet" & _
"Explorer 5.5 or later</font></p><p><img border=""0""" & _
"src=""arrowanim.gif"" width=""136"" height=""62""></p>" & _
"<p>&nbsp;</p>" & vbCrLf & "</body>" & vbCrLf & "</html>"
f.close
frmPreview.Show
End Sub

Private Sub scrollbar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightLabel (scparts(Index))
End Sub
