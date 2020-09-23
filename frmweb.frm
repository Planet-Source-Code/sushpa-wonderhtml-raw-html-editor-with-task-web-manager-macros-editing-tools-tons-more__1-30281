VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWeb 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WonderHTML: Browse Location"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmweb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.FileListBox fTmp 
      Height          =   870
      Left            =   2925
      TabIndex        =   9
      Top             =   945
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   2745
      Top             =   2385
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmweb.frx":2D2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txName 
      Height          =   315
      Left            =   3060
      TabIndex        =   6
      Top             =   2025
      Width           =   2355
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   4365
      TabIndex        =   5
      Tag             =   "Close this dialog without loading or creating the web."
      Top             =   2520
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   390
      Left            =   3330
      TabIndex        =   4
      Tag             =   "Continue to create or load the selected web."
      Top             =   2520
      Width           =   960
   End
   Begin VB.DirListBox Folder 
      Height          =   1890
      Left            =   135
      TabIndex        =   0
      Tag             =   "The FolderList displays all folders on your system."
      Top             =   1035
      Width           =   2220
   End
   Begin VB.DriveListBox Drv 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Tag             =   "The DriveList shows the drives on your computer."
      Top             =   450
      Width           =   2220
   End
   Begin VB.DirListBox drTmp 
      Height          =   540
      Left            =   3465
      TabIndex        =   8
      Top             =   3060
      Visible         =   0   'False
      Width           =   870
   End
   Begin MSComctlLib.TreeView tvWebs 
      Height          =   1545
      Left            =   2520
      TabIndex        =   10
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2725
      _Version        =   393217
      Indentation     =   317
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "iml32"
      Appearance      =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Click 'Continue' to proceed."
      Height          =   195
      Left            =   3105
      TabIndex        =   12
      Top             =   1250
      Width           =   1545
   End
   Begin VB.Label lbTemplate 
      Caption         =   " Template (optional):"
      Height          =   240
      Left            =   2520
      TabIndex        =   11
      Top             =   135
      Width           =   1590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "N&ame:"
      Height          =   195
      Left            =   2565
      TabIndex        =   7
      Top             =   2070
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "P&ath:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   855
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "D&rive:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   225
      Width           =   435
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'######################################
'WonderHTML 0.90 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
'nicer dialog this time: 15 Sep
Option Explicit

Private Sub cmdNo_Click()
ReturnedPath = ""
Unload Me
End Sub

Private Sub cmdOK_Click()
If txName.Enabled = True And txName.Text = "" And Height > 3500 Then MsgBox "The Name field is required.", vbExclamation: Exit Sub
Dim StrPS As String
If Right(Folder.Path, 1) <> "\" Then StrPS = "\"
If txName.Enabled And Height > 3500 And Dir(Folder.Path & StrPS & txName.Text, vbDirectory) <> "" Then MsgBox "Directory already exists." & vbNewLine & "Use the open web feature.", vbExclamation: Exit Sub
ReturnedPath = Folder.Path & StrPS & txName.Text
CopyTemplateFiles
Unload Me
End Sub

Private Sub drTmp_Change()
fTmp.Path = drTmp.Path
End Sub

Private Sub Drv_Change()
On Error GoTo hell
Folder.Path = Drive(Drv.Drive)
Exit Sub
hell:
MsgBox Error, vbExclamation
End Sub

Private Sub Form_Load()
SetFont Me
AddTemplates
SetWindowPos hwnd, -1, Left / 15, Top / 15, Width / 15, Height / 15, 0&
End Sub

Sub AddTemplates()
On Error Resume Next
Dim i As Long
drTmp.Path = FullPath(App.Path, "Templates")
For i = 0 To drTmp.ListCount - 1
tvWebs.Nodes.Add , , drTmp.list(i), GetFile(drTmp.list(i)), 1
Next i
End Sub

Sub CopyTemplateFiles()
On Error Resume Next
If tvWebs.SelectedItem Is Nothing Or tvWebs.SelectedItem.Index = 1 Then tvWebs.SelectedItem = tvWebs.Nodes(2)
drTmp.Path = tvWebs.SelectedItem.Key
Dim i As Integer
For i = 0 To fTmp.ListCount - 1
FileCopy FullPath(fTmp.Path, fTmp.list(i)), tvWebs.SelectedItem.Key
Next i
End Sub

Private Sub tvWebs_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Index = 1 Then txName.Text = "": Exit Sub
txName.Text = Replace(Node.Text, " ", "")
Dim i As Integer
Do
i = i + 1
If Dir(FullPath(Folder.Path, txName.Text & i), vbDirectory) = "" Then Exit Do
'find index which is available. e.g. if myweb1 exists this helps to make myweb2.
Loop
txName.Text = txName.Text & i
End Sub
