VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WonderHTML: Find"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmstart.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox txF 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   630
      Width           =   2940
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   18
      Top             =   2250
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   "Seek - and thou shalt find."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1710
      Width           =   1005
   End
   Begin VB.CommandButton cmFind 
      Caption         =   "&Search..."
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   450
      Width           =   1005
   End
   Begin VB.CheckBox chWhole 
      Caption         =   "Find wh&ole words only"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   1890
      Width           =   3030
   End
   Begin VB.CheckBox chCase 
      Caption         =   "Mat&ch case"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   1665
      Width           =   3075
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   1500
      Left            =   75
      TabIndex        =   23
      Top             =   2295
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   2646
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "iml13"
      SmallIcons      =   "iml13"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7039
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2175
      Left            =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   45
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3836
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " Search"
      TabPicture(0)   =   "frmstart.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " Replace"
      TabPicture(1)   =   "frmstart.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txR"
      Tab(1).Control(1)=   "cmRepThis"
      Tab(1).Control(2)=   "cmRepAll"
      Tab(1).Control(3)=   "Label1(2)"
      Tab(1).Control(4)=   "Label1(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Find in files"
      TabPicture(2)   =   "frmstart.frx":02C0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmSearchPath"
      Tab(2).Control(1)=   "txFindFiles"
      Tab(2).Control(2)=   "cmCancelPath"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "chSub"
      Tab(2).Control(4)=   "cmBrowse"
      Tab(2).Control(5)=   "txLoc"
      Tab(2).Control(6)=   "txPattern"
      Tab(2).Control(7)=   "chFileCase"
      Tab(2).Control(8)=   "Label4"
      Tab(2).Control(9)=   "Label3"
      Tab(2).Control(10)=   "Label2"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Go to line"
      TabPicture(3)   =   "frmstart.frx":02DC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lbG"
      Tab(3).Control(1)=   "Label6"
      Tab(3).Control(2)=   "lbDetPos"
      Tab(3).Control(3)=   "txG"
      Tab(3).Control(4)=   "opG(0)"
      Tab(3).Control(5)=   "opG(1)"
      Tab(3).Control(6)=   "opG(2)"
      Tab(3).Control(7)=   "Frame1"
      Tab(3).Control(8)=   "cbAbsRel"
      Tab(3).Control(9)=   "cbAB"
      Tab(3).Control(10)=   "cmG"
      Tab(3).Control(11)=   "chCloseGo"
      Tab(3).ControlCount=   12
      Begin VB.CheckBox chCloseGo 
         Caption         =   "Close"
         Height          =   195
         Left            =   -72570
         TabIndex        =   36
         Top             =   1755
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CommandButton cmG 
         Caption         =   "&Go"
         Height          =   375
         Left            =   -71760
         TabIndex        =   35
         Top             =   1665
         Width           =   960
      End
      Begin VB.ComboBox cbAB 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmstart.frx":02F8
         Left            =   -71580
         List            =   "frmstart.frx":0302
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1215
         Width           =   780
      End
      Begin VB.ComboBox cbAbsRel 
         Height          =   315
         ItemData        =   "frmstart.frx":0315
         Left            =   -72570
         List            =   "frmstart.frx":031F
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1215
         Width           =   960
      End
      Begin VB.Frame Frame1 
         Height          =   60
         Left            =   -74730
         TabIndex        =   31
         Top             =   1485
         Width           =   1950
      End
      Begin VB.OptionButton opG 
         Caption         =   "Line..."
         Height          =   240
         Index           =   2
         Left            =   -74685
         TabIndex        =   30
         Top             =   1665
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton opG 
         Caption         =   "End of Document"
         Height          =   240
         Index           =   1
         Left            =   -74685
         TabIndex        =   29
         Top             =   1125
         Width           =   1815
      End
      Begin VB.OptionButton opG 
         Caption         =   "Start of Document"
         Height          =   240
         Index           =   0
         Left            =   -74685
         TabIndex        =   28
         Top             =   855
         Width           =   1815
      End
      Begin VB.TextBox txG 
         Height          =   315
         Left            =   -72570
         TabIndex        =   26
         Text            =   "0"
         Top             =   630
         Width           =   1770
      End
      Begin VB.CommandButton cmSearchPath 
         Caption         =   "&Search"
         Height          =   375
         Left            =   -72525
         TabIndex        =   13
         Top             =   1635
         Width           =   870
      End
      Begin VB.ComboBox txFindFiles 
         Height          =   315
         Left            =   -74415
         TabIndex        =   8
         Top             =   450
         Width           =   3615
      End
      Begin VB.ComboBox txR 
         Height          =   315
         Left            =   -74860
         TabIndex        =   5
         Top             =   1170
         Width           =   2940
      End
      Begin VB.CommandButton cmCancelPath 
         Caption         =   "S&top"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71625
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1635
         Width           =   825
      End
      Begin VB.CheckBox chSub 
         Caption         =   "Scan subfolders"
         Height          =   195
         Left            =   -74865
         TabIndex        =   11
         Top             =   1395
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton cmBrowse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -71130
         Picture         =   "frmstart.frx":0337
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1035
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox txLoc 
         Height          =   315
         Left            =   -74865
         TabIndex        =   9
         Top             =   1035
         Width           =   3705
      End
      Begin VB.CommandButton cmRepThis 
         Caption         =   "&Replace"
         Height          =   375
         Left            =   -71805
         TabIndex        =   6
         Top             =   810
         Width           =   1005
      End
      Begin VB.CommandButton cmRepAll 
         Caption         =   "Replace &all"
         Height          =   375
         Left            =   -71805
         TabIndex        =   7
         Top             =   1215
         Width           =   1005
      End
      Begin MSComctlLib.ImageList iml13 
         Left            =   -180
         Top             =   -360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmstart.frx":0481
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox txPattern 
         Height          =   315
         ItemData        =   "frmstart.frx":14D5
         Left            =   -74280
         List            =   "frmstart.frx":14F7
         Sorted          =   -1  'True
         TabIndex        =   12
         Text            =   "*.html"
         Top             =   1755
         Width           =   855
      End
      Begin VB.CheckBox chFileCase 
         Caption         =   "Ignor&e case"
         Height          =   195
         Left            =   -73065
         TabIndex        =   24
         Top             =   1395
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Label lbDetPos 
         Caption         =   "Determine position:"
         Height          =   195
         Left            =   -72570
         TabIndex        =   34
         Top             =   990
         Width           =   1770
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Go to what:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   27
         Top             =   495
         Width           =   855
      End
      Begin VB.Label lbG 
         AutoSize        =   -1  'True
         Caption         =   "Enter Line Number:"
         Height          =   195
         Left            =   -72570
         TabIndex        =   25
         Top             =   405
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "In files:"
         Height          =   195
         Left            =   -74865
         TabIndex        =   22
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
         Height          =   195
         Left            =   -74865
         TabIndex        =   21
         Top             =   825
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Find:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   19
         Top             =   495
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Search for:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Search for:"
         Height          =   195
         Index           =   2
         Left            =   -74865
         TabIndex        =   16
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "R&eplace with:"
         Height          =   195
         Index           =   1
         Left            =   -74865
         TabIndex        =   15
         Top             =   945
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmFind"
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
'Made this form better using the tab control
Option Explicit
Dim vCase As Long, vWord As Long
Dim lpStart As Long
Private Const HEIGHT_NORMAL = 2910
Private Const HEIGHT_LVIEW = 4530

Sub cbAbsRel_Click()
cbAB.Enabled = CBool(cbAbsRel.ListIndex = 1)
End Sub

Private Sub chCase_Click()
On Error Resume Next
txF.SetFocus
If chCase.Value = 1 Then
vCase = rtfMatchCase
Else
vCase = 0
End If
End Sub

Private Sub chFileCase_Click()
On Error Resume Next
txFindFiles.SetFocus
End Sub

Private Sub chWhole_Click()
On Error Resume Next
txF.SetFocus
If chWhole.Value = 1 Then
vWord = rtfWholeWord
Else
vWord = 0
End If
End Sub

Private Sub cmBrowse_Click()
Dim s As String
s = SelectDir(True, 3465)
If s <> "" Then txLoc.Text = s
End Sub

Private Sub cmCancelPath_Click()
cmCancelPath.Enabled = False
cmSearchPath.Enabled = True
cmNo.Cancel = True
ExitFlag = True
SearchFlag = False
SB.SimpleText = "Stopped searching."
frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
End Sub

Private Sub cmFind_Click()
On Error Resume Next
Dim Fin As Long
If txF.Text = "" Then Exit Sub
Fin = frmMain.ActiveForm.RTF1.Find(txF.Text, lpStart, , vCase + vWord)
If Fin > 0 Then
lpStart = Fin + 1
If IsContained(txF.Text, txF) = False Then txF.AddItem txF.Text, 0
LastFindText = txF.Text
frmMain.ActiveForm.RTF1.SetFocus
Me.Move 0, 0
Else
SB.SimpleText = "'" & txF.Text & "' cannot be found."
End If
End Sub

Private Sub cmG_Click()
On Error Resume Next
Dim lngStart As Long
With frmMain.ActiveForm.RTF1
If opG(2).Value = True Then 'If Go To line is checked
  If cbAbsRel.ListIndex = 0 Then
    'Get pos of start of the line
    lngStart = SendMessage(.hwnd, EM_LINEINDEX, CLng(txG.Text) - 1, 0&)
    If lngStart = -1 Then 'Invalid line number
        SB.SimpleText = "Can't go. Line number is invalid."
        Exit Sub
    End If
    .SelStart = lngStart 'Go To line
  Else
    If cbAB.ListIndex = 0 Then
      lngStart = SendMessage(.hwnd, EM_LINEINDEX, CLng(txG.Text) + GetCurrentLine(frmMain.ActiveForm.RTF1) - 1, 0&)
      If lngStart = -1 Then 'Invalid line number
        SB.SimpleText = "Can't go. Line number is invalid."
        Exit Sub
      End If
    Else
      lngStart = SendMessage(.hwnd, EM_LINEINDEX, GetCurrentLine(frmMain.ActiveForm.RTF1) - CLng(txG.Text) - 1, 0&)
      If lngStart = -1 Then 'Invalid line number
        SB.SimpleText = "Can't go. Line number is invalid."
        Exit Sub
      End If
    End If
  .SelStart = lngStart
  End If
ElseIf opG(0).Value = True Then 'Go To start of the document
    .SelStart = 0
ElseIf opG(1).Value = True Then 'Go To end of the document
    .SelStart = Len(.Text)
End If
End With
  If chCloseGo.Value = 1 Then Unload Me
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmRepAll_Click()
On Error Resume Next
If txF.Text = "" Then Exit Sub
Dim l As Long
l = frmMain.ActiveForm.RTF1.SelStart
frmMain.ActiveForm.RTF1.Text = Replace(frmMain.ActiveForm.RTF1.Text, txF.Text, txR.Text)
frmMain.ActiveForm.RTF1.SelStart = l
If IsContained(txR.Text, txR) = False Then txR.AddItem txR.Text, 0
frmMain.ActiveForm.RTF1.SetFocus
End Sub

Private Sub cmRepThis_Click()
On Error Resume Next
If txF.Text = "" Then Exit Sub
If frmMain.ActiveForm.RTF1.SelLength = 0 Then Exit Sub
frmMain.ActiveForm.RTF1.SelText = txR.Text
frmMain.ActiveForm.RTF1.Find txF.Text, lpStart, , vCase + vWord
If IsContained(txR.Text, txR) = False Then txR.AddItem txR.Text, 0
frmMain.ActiveForm.RTF1.SetFocus
Me.Move 0, 0
End Sub

Private Sub cmSearchPath_Click()
On Error Resume Next
    Dim dC As Integer
    If txFindFiles.Text = "" Then SB.SimpleText = "Seek 'something', only then thou shalt find.": Exit Sub
    cmSearchPath.Enabled = False
    cmCancelPath.Enabled = True
    cmCancelPath.Cancel = True
    frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(2).Picture
    If txLoc.Text = "" Then txLoc.Text = App.Path
    lvFiles.ListItems.Clear
    Height = HEIGHT_LVIEW
    ExitFlag = False
    frmMain.Fil.Pattern = txPattern.Text
    DirDiver txLoc.Text, txLoc.Text, CBool(chSub.Value), ACTION_FINDFILES, 0, txFindFiles.Text, IIf(chFileCase.Value = 1, vbTextCompare, vbBinaryCompare)
    frmMain.Fil.Pattern = "*.*"
    ExitFlag = True
    SearchFlag = False
    SB.SimpleText = "Finished searching. " & lvFiles.ListItems.count & " matches found."
    frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
    cmSearchPath.Enabled = True
    cmCancelPath.Enabled = False
    If Not IsContained(txFindFiles.Text, txFindFiles) Then txFindFiles.AddItem txFindFiles.Text, 0
    cmNo.Cancel = True
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
SetWindowPos hwnd, -1, Left / 15, Top / 15, Width / 15, Height / 15, 0&
txLoc.Text = IIf(frmMain.tvW.Nodes.count > 0, (frmMain.tvW.Nodes(1).Text), (App.Path))
lpStart = 1
cbAbsRel.ListIndex = 0
cbAB.ListIndex = 0
LoadListItems
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveListItems
End Sub

Private Sub lvFiles_DblClick()
On Error Resume Next
If lvFiles.ListItems.count = 0 Then Exit Sub
Dim lpF As New frmChild
Load lpF
lpF.LoadHTMLFile lvFiles.SelectedItem.Key
lpF.RTF1.SelStart = InStr(1, lpF.RTF1.Text, txFindFiles.Text, vbTextCompare) - 1
'lpF.RTF1.SelLength = Len(txFindFiles.Text)
Me.Move 0, 0
End Sub

Private Sub lvFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
SB.SimpleText = Item.Key
End Sub

Private Sub opG_Click(Index As Integer)
On Error Resume Next
txG.Enabled = CBool(Index = 2)
cbAB.Enabled = txG.Enabled
cbAbsRel.Enabled = txG.Enabled
lbG.Enabled = txG.Enabled
lbDetPos.Enabled = txG.Enabled
txG.SetFocus
End Sub

Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 2 Or SSTab1.Tab = 3 Then SSTab1.ZOrder vbBringToFront Else SSTab1.ZOrder vbSendToBack
Caption = "WonderHTML: " & SSTab1.Caption
If SSTab1.Tab = 2 Then txFindFiles.SetFocus Else If SSTab1.Tab = 3 Then txG.SetFocus Else txF.SetFocus
End Sub

Private Sub txF_Change()
lpStart = 1
End Sub

Sub LoadListItems()
Dim i As Integer, s As String
For i = 0 To 9
s = ReadValue("FindText" & i, "", "Find")
If s <> "" Then txF.AddItem s
Next i
For i = 0 To 9
s = ReadValue("FindRepl" & i, "", "Find")
If s <> "" Then txR.AddItem s
Next i
For i = 0 To 9
s = ReadValue("FindFiles" & i, "", "Find")
If s <> "" Then txFindFiles.AddItem s
Next i
End Sub

Sub SaveListItems()
On Error Resume Next
Dim i As Integer
For i = 0 To 9
SaveValue "FindText" & i, txF.list(i), "Find"
Next i
For i = 0 To 9
SaveValue "FindRepl" & i, txR.list(i), "Find"
Next i
For i = 0 To 9
SaveValue "FindFiles" & i, txFindFiles.list(i), "Find"
Next i
End Sub

Private Sub txG_GotFocus()
txG.SelStart = 0
txG.SelLength = Len(txG.Text)
End Sub

Private Sub txLoc_GotFocus()
txLoc.SelStart = 0
txLoc.SelLength = Len(txLoc.Text)
End Sub
