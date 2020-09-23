VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConv 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML to Text Converter"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin TabDlg.SSTab SSTab1 
      Height          =   1815
      Left            =   60
      TabIndex        =   4
      Top             =   45
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   3201
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Single File"
      TabPicture(0)   =   "frmConv.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbFile"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Batch"
      TabPicture(1)   =   "frmConv.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lsFiles"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmAdd"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmRem"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmRem 
         Height          =   375
         Left            =   3345
         Picture         =   "frmConv.frx":0044
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1350
         Width           =   375
      End
      Begin VB.CommandButton cmAdd 
         Height          =   375
         Left            =   3345
         Picture         =   "frmConv.frx":018E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   945
         Width           =   375
      End
      Begin VB.ListBox lsFiles 
         Height          =   1335
         IntegralHeight  =   0   'False
         Left            =   90
         TabIndex        =   9
         Top             =   375
         Width           =   3210
      End
      Begin VB.TextBox lbFile 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -74910
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Idle"
         Top             =   450
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Options"
         Height          =   915
         Left            =   -74865
         TabIndex        =   5
         Top             =   765
         Width           =   3525
         Begin VB.CheckBox chConvEntities 
            Caption         =   "&Convert numerical HTML entities"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   495
            Value           =   1  'Checked
            Width           =   3075
         End
         Begin VB.CheckBox chHR 
            Caption         =   "&Replace [HR] tags with hyphens"
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   270
            Value           =   1  'Checked
            Width           =   3075
         End
      End
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   2940
      TabIndex        =   3
      Top             =   1935
      Width           =   915
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Convert"
      Default         =   -1  'True
      Height          =   375
      Left            =   1905
      TabIndex        =   2
      Top             =   1935
      Width           =   960
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   1980
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Max             =   32767
   End
   Begin RichTextLib.RichTextBox rtfTemp 
      Height          =   1275
      Left            =   -68
      TabIndex        =   0
      Top             =   4365
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   2249
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmConv.frx":0718
   End
End
Attribute VB_Name = "frmConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I put this new. It provides some options and a better view
'especially with the progress bar.
'Now also supports batch conversion.
Option Explicit
Public FileL As Long
Dim pos As Long

Private Sub cmAdd_Click()
On Error GoTo hell
If lsFiles.ListCount = 20 Then
  If MsgBox("That's too many files. Lots of memory and resources" & vbCrLf & "may get used up because of it. Continue?", vbYesNo + vbCritical, "Convert batch") = vbNo Then Exit Sub
End If
Dim s As String
s = frmMain.CD.Filter
frmMain.CD.Filter = "HTML Files (*.html, *.htm)|*.html;*.htm"
frmMain.CD.ShowOpen
lsFiles.AddItem frmMain.CD.Filename
hell:
frmMain.CD.Filter = s
lsFiles.SetFocus
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
On Error Resume Next
cmOK.Enabled = False
cmNo.Enabled = False
chHR.Enabled = False
chConvEntities.Enabled = False

If lbFile.Enabled = True Then
  start lbFile.Text
Else
  Dim i As Integer
  For i = 0 To lsFiles.ListCount - 1
    start lsFiles.list(i)
  Next i
End If

cmOK.Enabled = True
cmNo.Enabled = True
chHR.Enabled = True
chConvEntities.Enabled = True
MousePointer = 0: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
Unload Me
End Sub

Private Sub cmRem_Click()
On Error Resume Next
lsFiles.RemoveItem lsFiles.ListIndex
lsFiles.SetFocus
End Sub

Private Sub Form_Load()
SetFont Me
End Sub

Private Sub rtfTemp_SelChange()
On Error Resume Next
If Height > 3500 Then Exit Sub
PB.Value = rtfTemp.SelStart
pos = (rtfTemp.SelStart * 100) / FileL
frmMain.SB.Panels(1).Text = "Converting " & GetFile(lbFile.Text) & ": " & pos & "% completed..."
End Sub

Sub start(File As String)
On Error Resume Next
lbFile.Text = File
rtfTemp.LoadFile File, rtfText
rtfTemp.Text = Replace(rtfTemp.Text, Chr(13), "")
rtfTemp.Text = Replace(rtfTemp.Text, Chr(10), "")
FileL = Len(rtfTemp.Text)
PB.Max = Len(rtfTemp.Text)
MousePointer = 11: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(2).Picture
CleanUp rtfTemp, CBool(chHR.Value)
If CBool(chConvEntities.Value) = True Then ConvEntities rtfTemp
ReplaceStuff rtfTemp


Dim lpD As New frmChild
Load lpD
lpD.Icon = lpD.p2.Picture
lpD.RTF1.Text = rtfTemp.Text
lpD.RTF1.SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
lbFile.Enabled = (SSTab1.Tab = 0)
End Sub
