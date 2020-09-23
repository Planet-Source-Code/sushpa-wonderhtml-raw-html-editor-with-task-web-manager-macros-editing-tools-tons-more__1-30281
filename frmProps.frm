VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProps 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Document Properties"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "frmProps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ProgressBar PB 
      Height          =   240
      Left            =   135
      TabIndex        =   31
      Top             =   3510
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncel"
      Height          =   375
      Left            =   4680
      TabIndex        =   30
      Top             =   3420
      Width           =   960
   End
   Begin VB.CommandButton cmOk 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   3645
      TabIndex        =   29
      Top             =   3420
      Width           =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3300
      Left            =   60
      TabIndex        =   13
      Top             =   45
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   5821
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " General"
      TabPicture(0)   =   "frmProps.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   " Declarations"
      TabPicture(1)   =   "frmProps.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmRem"
      Tab(1).Control(1)=   "cmNew"
      Tab(1).Control(2)=   "cmMetaID"
      Tab(1).Control(3)=   "cmMETAType"
      Tab(1).Control(4)=   "lvMETA"
      Tab(1).Control(5)=   "Line1"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   " Advanced"
      TabPicture(2)   =   "frmProps.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame7 
         Caption         =   "Page Margins"
         Height          =   1770
         Left            =   -71850
         TabIndex        =   46
         Top             =   1395
         Width           =   2355
         Begin VB.TextBox txMrgnWidth 
            Height          =   315
            Left            =   675
            TabIndex        =   53
            Top             =   1305
            Width           =   1545
         End
         Begin VB.TextBox txMrgnHeight 
            Height          =   315
            Left            =   675
            TabIndex        =   52
            Top             =   960
            Width           =   1545
         End
         Begin VB.TextBox txMrgnLeft 
            Height          =   315
            Left            =   675
            TabIndex        =   51
            Top             =   600
            Width           =   1545
         End
         Begin VB.TextBox txMrgnTop 
            Height          =   315
            Left            =   675
            TabIndex        =   50
            Top             =   255
            Width           =   1545
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Width:"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   54
            Top             =   1365
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Height:"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   49
            Top             =   1005
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Left:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   48
            Top             =   660
            Width           =   345
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Top:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   47
            Top             =   300
            Width           =   330
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Background sound"
         Height          =   1140
         Left            =   -74865
         TabIndex        =   39
         Top             =   405
         Width           =   2940
         Begin VB.CheckBox chLoop 
            Caption         =   "Loop sound continuously"
            Height          =   195
            Left            =   90
            TabIndex        =   45
            Top             =   810
            Width           =   2715
         End
         Begin VB.CommandButton cmBrowseBGSND 
            Caption         =   "..."
            Height          =   330
            Left            =   2475
            TabIndex        =   41
            Top             =   450
            Width           =   330
         End
         Begin VB.TextBox txBGSOUNDsrc 
            Height          =   315
            Left            =   90
            TabIndex        =   40
            Top             =   450
            Width           =   2355
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Source file:"
            Height          =   195
            Left            =   105
            TabIndex        =   42
            Top             =   245
            Width           =   810
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "BASE Target frame"
         Height          =   960
         Left            =   -71850
         TabIndex        =   37
         Top             =   405
         Width           =   2355
         Begin VB.ComboBox cbBaseTarget 
            Height          =   315
            ItemData        =   "frmProps.frx":0060
            Left            =   135
            List            =   "frmProps.frx":0070
            Sorted          =   -1  'True
            TabIndex        =   38
            Top             =   465
            Width           =   2070
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Default frame:"
            Height          =   195
            Left            =   180
            TabIndex        =   55
            Top             =   270
            Width           =   1050
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "InterDev-compatible options"
         Height          =   1545
         Left            =   -74865
         TabIndex        =   32
         Top             =   1620
         Width           =   2940
         Begin VB.ComboBox cbDTC_Server 
            Height          =   315
            ItemData        =   "frmProps.frx":0092
            Left            =   990
            List            =   "frmProps.frx":009F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1035
            Width           =   1770
         End
         Begin VB.ComboBox cbDTC_Client 
            Height          =   315
            ItemData        =   "frmProps.frx":00C4
            Left            =   990
            List            =   "frmProps.frx":00D1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   667
            Width           =   1770
         End
         Begin VB.ComboBox cbDTC_Platform 
            Height          =   315
            ItemData        =   "frmProps.frx":00F6
            Left            =   990
            List            =   "frmProps.frx":0103
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   300
            Width           =   1770
         End
         Begin VB.Label LBDTC 
            AutoSize        =   -1  'True
            Caption         =   "Server:"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   44
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label LBDTC 
            AutoSize        =   -1  'True
            Caption         =   "Client:"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   35
            Top             =   720
            Width           =   465
         End
         Begin VB.Label LBDTC 
            AutoSize        =   -1  'True
            Caption         =   "Platform:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   33
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.CommandButton cmRem 
         Caption         =   "Remove"
         Height          =   375
         Left            =   -73920
         TabIndex        =   10
         Top             =   2790
         Width           =   960
      End
      Begin VB.CommandButton cmNew 
         Caption         =   "New..."
         Height          =   375
         Left            =   -74865
         TabIndex        =   9
         Top             =   2790
         Width           =   915
      End
      Begin VB.CommandButton cmMetaID 
         Caption         =   "Edit"
         Height          =   375
         Left            =   -71490
         TabIndex        =   11
         Top             =   2790
         Width           =   780
      End
      Begin VB.CommandButton cmMETAType 
         Caption         =   "Edit value..."
         Height          =   375
         Left            =   -70680
         TabIndex        =   12
         Top             =   2790
         Width           =   1185
      End
      Begin MSComctlLib.ListView lvMETA 
         Height          =   2235
         Left            =   -74910
         TabIndex        =   8
         Top             =   405
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3942
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "iml16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Identifier"
            Object.Width           =   3242
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Content"
            Object.Width           =   3242
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   2435
         EndProperty
      End
      Begin VB.Frame Frame3 
         Caption         =   "Background Properties"
         Height          =   1815
         Left            =   2790
         TabIndex        =   24
         Top             =   1215
         Width           =   2715
         Begin VB.ComboBox txBGCol 
            Height          =   315
            ItemData        =   "frmProps.frx":013D
            Left            =   120
            List            =   "frmProps.frx":0171
            TabIndex        =   5
            Top             =   495
            Width           =   2175
         End
         Begin VB.TextBox txBGImage 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   1125
            Width           =   2175
         End
         Begin VB.CommandButton cmBGImage 
            Caption         =   "..."
            Height          =   330
            Left            =   2325
            TabIndex        =   28
            Top             =   1125
            Width           =   315
         End
         Begin VB.CommandButton cmBGColor 
            Caption         =   "..."
            Height          =   330
            Left            =   2325
            TabIndex        =   26
            Top             =   495
            Width           =   315
         End
         Begin VB.CheckBox chWatermark 
            Caption         =   "Non-scrolling background"
            Height          =   240
            Left            =   120
            TabIndex        =   7
            Top             =   1460
            Width           =   2490
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Image:"
            Height          =   195
            Left            =   135
            TabIndex        =   27
            Top             =   900
            Width           =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Colour:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Page Title"
         Height          =   645
         Left            =   2790
         TabIndex        =   23
         Top             =   495
         Width           =   2715
         Begin VB.TextBox txTitle 
            Height          =   315
            Left            =   90
            TabIndex        =   4
            Top             =   225
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Element Colours"
         Height          =   2715
         Left            =   135
         TabIndex        =   14
         Top             =   405
         Width           =   2580
         Begin VB.ComboBox txColVlink 
            Height          =   315
            ItemData        =   "frmProps.frx":01E3
            Left            =   90
            List            =   "frmProps.frx":0217
            TabIndex        =   3
            Top             =   2295
            Width           =   2040
         End
         Begin VB.ComboBox txColActive 
            Height          =   315
            ItemData        =   "frmProps.frx":0289
            Left            =   90
            List            =   "frmProps.frx":02BD
            TabIndex        =   2
            Top             =   1680
            Width           =   2040
         End
         Begin VB.ComboBox txColLink 
            Height          =   315
            ItemData        =   "frmProps.frx":032F
            Left            =   90
            List            =   "frmProps.frx":0363
            TabIndex        =   1
            Top             =   1065
            Width           =   2040
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   330
            Left            =   2160
            TabIndex        =   22
            Top             =   2295
            Width           =   315
         End
         Begin VB.CommandButton cmBrowseActiveLink 
            Caption         =   "..."
            Height          =   330
            Left            =   2160
            TabIndex        =   20
            Top             =   1680
            Width           =   315
         End
         Begin VB.CommandButton cmBrowseNormLink 
            Caption         =   "..."
            Height          =   330
            Left            =   2160
            TabIndex        =   18
            Top             =   1065
            Width           =   315
         End
         Begin VB.CommandButton cmBrowseColNorm 
            Caption         =   "..."
            Height          =   330
            Left            =   2160
            TabIndex        =   16
            Top             =   450
            Width           =   315
         End
         Begin VB.ComboBox txColNorm 
            Height          =   315
            ItemData        =   "frmProps.frx":03D5
            Left            =   90
            List            =   "frmProps.frx":0409
            TabIndex        =   0
            Top             =   450
            Width           =   2040
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Visited link colour:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   2070
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Active link colour:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1455
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Link colour:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Normal text colour:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   225
            Width           =   1380
         End
      End
      Begin VB.Line Line1 
         X1              =   -74865
         X2              =   -69555
         Y1              =   2700
         Y2              =   2700
      End
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   495
      Top             =   3015
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProps.frx":047B
            Key             =   "def"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProps.frx":0A17
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProps.frx":0FB3
            Key             =   "description"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProps.frx":154F
            Key             =   "keywords"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProps.frx":1AEB
            Key             =   "generator"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProps.frx":2B3F
            Key             =   "author"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProps.frx":30DB
            Key             =   "vi60_defaultclientscript"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProps.frx":3477
            Key             =   "vi60_dtcscriptingplatform"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmBGColor_Click()
BrowseClr txBGCol
End Sub

Private Sub cmBGImage_Click()
On Error GoTo hell
Dim s As String
s = frmMain.CD.Filter
frmMain.CD.Filter = "Images (*.bmp, *.jpg, *.gif)|*.bmp;*.gif;*.jpg;*.jpeg;*.jpe"
frmMain.CD.ShowOpen
frmMain.CD.Filter = s
s = frmMain.CD.Filename
If frmMain.ActiveForm.Caption = "Untitled" Then GoTo absolute
'get relative link
s = Replace(s, Up1Level(frmMain.ActiveForm.Caption), "", , , vbTextCompare)
s = FirstSlash(s)
absolute:
s = Replace(s, "\", "/")
txBGImage.Text = s
txBGImage.SetFocus
hell:
End Sub

Private Sub cmBrowseActiveLink_Click()
BrowseClr txColActive
End Sub

Private Sub cmBrowseBGSND_Click()
On Error GoTo hell
Dim s As String
s = frmMain.CD.Filter
frmMain.CD.Filter = "Microsoft WaveForm Audio (*.wav)|*.wav"
frmMain.CD.ShowOpen
frmMain.CD.Filter = s
s = frmMain.CD.Filename
If frmMain.ActiveForm.Caption = "Untitled" Then GoTo absolute
'get relative link
s = Replace(s, Up1Level(frmMain.ActiveForm.Caption), "", , , vbTextCompare)
s = FirstSlash(s)
absolute:
s = Replace(s, "\", "/")
txBGSOUNDsrc.Text = s
txBGSOUNDsrc.SetFocus
hell:
End Sub

Private Sub cmBrowseColNorm_Click()
BrowseClr txColNorm
End Sub

Private Sub cmBrowseNormLink_Click()
BrowseClr txColLink
End Sub


Private Sub cmMETAType_Click()
On Error Resume Next
Dim s As String
If lvMETA.SelectedItem Is Nothing Then Exit Sub
s = lvMETA.SelectedItem.Text
s = GetMetaType(s, lvMETA.SelectedItem.ListSubItems(2).Text, lvMETA.SelectedItem.ListSubItems(1).Text)
If s <> "" Then
  Dim lp() As String
  lp = Split(s, "ÿþýüûú") 'junk chars form a delimiter
  lvMETA.SelectedItem.ListSubItems(2).Text = lp(0)
  lvMETA.SelectedItem.ListSubItems(1).Text = lp(1)
End If
lvMETA.SetFocus
End Sub

Private Sub cmNew_Click()
On Error Resume Next
lvMETA.ListItems.Add , "tmp", , , 1
lvMETA.ListItems("tmp").ListSubItems.Add 1, , "(none)"
lvMETA.ListItems("tmp").ListSubItems.Add 2, , "name"
lvMETA.SelectedItem = lvMETA.ListItems("tmp")
lvMETA.ListItems("tmp").Key = ""
lvMETA.SetFocus
lvMETA.StartLabelEdit
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
On Error Resume Next
Dim l As Long
frmMain.ActiveForm.RTF1.Visible = False
Screen.MousePointer = 11: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(2).Picture
frmMain.SB.Style = sbrSimple
frmMain.SB.SimpleText = "Please wait, applying changes..."
pB.Visible = True
pB.Max = 15 + 2 * lvMETA.ListItems.count '15 standard actions + metas add/remove
l = frmMain.ActiveForm.RTF1.SelStart
SetDTCInfo
SetMetaTags
SetDocInfo
frmMain.ActiveForm.RTF1.SelStart = l
frmMain.SB.Style = sbrNormal
Screen.MousePointer = 0: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(0).Picture
frmMain.ActiveForm.RTF1.Visible = True
Unload Me
End Sub

Private Sub cmRem_Click()
On Error Resume Next
lvMETA.ListItems.Remove lvMETA.SelectedItem.Index
lvMETA.SetFocus
End Sub

Private Sub Command1_Click()
BrowseClr txColVlink
End Sub

Private Sub cmMetaID_Click()
On Error Resume Next
lvMETA.SetFocus
lvMETA.StartLabelEdit
End Sub

Private Sub Form_Load()
SetFont Me
DoEvents
cbDTC_Client.ListIndex = 0
cbDTC_Platform.ListIndex = 0
cbDTC_Server.ListIndex = 0
LoadDocInfo
GetFrames
End Sub

Sub LoadDocInfo()
On Error Resume Next
Dim whole As String, attrib As String
Dim pos As Long, pos2 As Long
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<BODY", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
whole = Mid$(frmMain.ActiveForm.RTF1.Text, pos, pos2 - pos)
'bgcolor
attrib = ReadAttrib("bgcolor", whole)
If attrib <> whole Then txBGCol.Text = attrib
'text col
attrib = ReadAttrib("text", whole)
If attrib <> whole Then txColNorm.Text = attrib
'link col
attrib = ReadAttrib("link", whole)
If attrib <> whole Then txColLink.Text = attrib
'alink col
attrib = ReadAttrib("alink", whole)
If attrib <> whole Then txColActive.Text = attrib
'vlink col
attrib = ReadAttrib("vlink", whole)
If attrib <> whole Then txColVlink.Text = attrib
'bg image
attrib = ReadAttrib("background", whole)
If attrib <> whole Then txBGImage.Text = attrib
'leftmargin
attrib = ReadAttrib("leftmargin", whole)
If attrib <> whole Then txMrgnLeft.Text = attrib
'topmargin
attrib = ReadAttrib("topmargin", whole)
If attrib <> whole Then txMrgnTop.Text = attrib
'marginheight
attrib = ReadAttrib("marginheight", whole)
If attrib <> whole Then txMrgnHeight.Text = attrib
'marginwidth
attrib = ReadAttrib("marginwidth", whole)
If attrib <> whole Then txMrgnWidth.Text = attrib
'watermark
attrib = ReadAttrib("bgproperties", whole)
chWatermark.Value = CBinary(LCase(attrib) = "fixed")
'title
txTitle.Text = GetTitleFromText(frmMain.ActiveForm.RTF1.Text)
'bgsound
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<BGSOUND", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
whole = Mid$(frmMain.ActiveForm.RTF1.Text, pos, pos2 - pos)
attrib = ReadAttrib("src", whole)
If attrib <> whole Then txBGSOUNDsrc.Text = attrib
attrib = ReadAttrib("loop", whole)
If attrib = "-1" Then chLoop.Value = 1
'base target
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<BASE", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
whole = Mid$(frmMain.ActiveForm.RTF1.Text, pos, pos2 - pos)
attrib = ReadAttrib("target", whole)
If attrib <> whole Then cbBaseTarget.Text = attrib
'load other things
LoadMetaTags
LoadDTCInfo
End Sub

Private Sub lvMETA_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error Resume Next
lvMETA.SelectedItem.SmallIcon = IIf(InStr(" keywords description refresh author vi60_defaultclientscript vi60_dtcscriptingplatform ", " " & LCase(NewString) & " "), LCase(NewString), "def")
End Sub

Private Sub lvMETA_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvMETA.Sorted = False
lvMETA.SortKey = ColumnHeader.Index - 1
lvMETA.Sorted = True
End Sub

Sub LoadMetaTags()
On Error Resume Next
lvMETA.ListItems.Clear
Dim i As Long
Dim whole As String, attrib As String
Dim typeid As String, Content As String, ID As String
Dim pos As Long, pos2 As Long
Do
pos = InStr(pos2 + 1, frmMain.ActiveForm.RTF1.Text, "<META", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
If pos = 0 Or pos2 = 0 Then Exit Do
whole = Mid$(frmMain.ActiveForm.RTF1.Text, pos, pos2 - pos)
ID = ReadAttrib("name", whole): typeid = "name"
If ID = whole Then ID = ReadAttrib("http-equiv", whole): typeid = "http-equiv"
If ID = whole Then GoTo nxt 'skip this process and try next tag. no ID found.
Content = ReadAttrib("content", whole)
If Content = whole Then Content = ""
lvMETA.ListItems.Add , "tmp", ID, , IIf(InStr(" keywords description refresh author vi60_defaultclientscript vi60_dtcscriptingplatform ", " " & LCase(ID) & " "), LCase(ID), "def")
lvMETA.ListItems("tmp").ListSubItems.Add 1, , Content
lvMETA.ListItems("tmp").ListSubItems.Add 2, , typeid
If LCase(lvMETA.ListItems("tmp").Text) = "vi60_defaultclientscript" Then
lvMETA.ListItems("tmp").Key = "dtc_client"
lvMETA.ListItems("dtc_client").Bold = True
ElseIf LCase(lvMETA.ListItems("tmp").Text) = "vi60_dtcscriptingplatform" Then
lvMETA.ListItems("tmp").Key = "dtc_platform"
lvMETA.ListItems("dtc_platform").Bold = True
Else
lvMETA.ListItems("tmp").Key = ""
End If
nxt:
Loop
End Sub

Sub SetMetaTags()
On Error Resume Next
Dim i As Long
Dim pos As Long, pos2 As Long
Dim sels As Long
sels = InStr(1, frmMain.ActiveForm.RTF1.Text, "<META", vbTextCompare)
If sels = 0 Then
  pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<HEAD>" & vbCrLf, vbTextCompare) + Len("<HEAD>" & vbCrLf)
  If pos = Len("<HEAD>" & vbCrLf) Then pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<HEAD>", vbTextCompare) + Len("<HEAD>")
  If pos = Len("<HEAD>") Then pos = 1
  frmMain.ActiveForm.RTF1.SelStart = pos - 1
  frmMain.ActiveForm.RTF1.SelText = "<META>" & vbCrLf
  sels = InStr(1, frmMain.ActiveForm.RTF1.Text, "<META>", vbTextCompare)
End If
frmMain.ActiveForm.RTF1.SelStart = sels - 1
'Find the positions of the META tags, select them and overwrite each value.
Do
pos = InStr(frmMain.ActiveForm.RTF1.SelStart + 1, frmMain.ActiveForm.RTF1.Text, "<META", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
If pos = 0 Then Exit Do
frmMain.ActiveForm.RTF1.SelStart = pos - 1
frmMain.ActiveForm.RTF1.SelLength = pos2 - pos + 1
frmMain.ActiveForm.RTF1.SelText = "" 'clear tags one by one
frmMain.ActiveForm.RTF1.SelLength = Len(vbCrLf)
If frmMain.ActiveForm.RTF1.SelText = vbCrLf Then frmMain.ActiveForm.RTF1.SelText = ""
pB.Value = pB.Value + 1
Loop
For i = 1 To lvMETA.ListItems.count
'add tags
frmMain.ActiveForm.RTF1.SelText = "<META " & lvMETA.ListItems(i).ListSubItems(2).Text & "=" & Chr(34) & lvMETA.ListItems(i).Text & Chr(34) & " content=" & Chr(34) & lvMETA.ListItems(i).ListSubItems(1).Text & Chr(34) & ">" & vbCrLf
pB.Value = pB.Value + 1
Next i
End Sub

Private Sub lvMETA_DblClick()
If cmMETAType.Enabled = True Then cmMETAType_Click
End Sub

Private Sub lvMETA_ItemClick(ByVal Item As MSComctlLib.ListItem)
cmMetaID.Enabled = Not Item.Bold
cmMETAType.Enabled = cmMetaID.Enabled
lvMETA.LabelEdit = 1 - CBinary(cmMETAType.Enabled)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Select Case SSTab1.Tab
Case 0
txColNorm.SetFocus
Case 1
lvMETA.SetFocus
Case 2
txBGSOUNDsrc.SetFocus
End Select
End Sub

Sub SetDocInfo()
On Error Resume Next
Dim pos As Long, pos2 As Long, whole As String
'set title
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<title>", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, "</title>", vbTextCompare)
pB.Value = pB.Value + 1
If (pos = 0 Or pos2 = 0) And txTitle.Text <> "" Then
  pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<HEAD>" & vbCrLf, vbTextCompare) + Len("<HEAD>" & vbCrLf)
  If pos = Len("<HEAD>" & vbCrLf) Then pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<HEAD>", vbTextCompare) + Len("<HEAD>")
  If pos = Len("<HEAD>") Then pos = 1
  frmMain.ActiveForm.RTF1.SelStart = pos - 1
  frmMain.ActiveForm.RTF1.SelText = "<TITLE>Applying...</TITLE>" & vbCrLf
  pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<title>", vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, "</title>", vbTextCompare)
End If
frmMain.ActiveForm.RTF1.SelStart = pos + 7 - 1 '7 is len(title)
frmMain.ActiveForm.RTF1.SelLength = pos2 - pos - 7
frmMain.ActiveForm.RTF1.SelText = txTitle.Text
nxt:
'set BGSOUND
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<bgsound", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
pB.Value = pB.Value + 1
If (pos = 0 Or pos2 = 0) And txBGSOUNDsrc.Text <> "" Then
  pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<HEAD>" & vbCrLf, vbTextCompare) + Len("<HEAD>" & vbCrLf)
  If pos = Len("<HEAD>" & vbCrLf) Then pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<HEAD>", vbTextCompare) + Len("<HEAD>")
  If pos = Len("<HEAD>") Then pos = 1
  frmMain.ActiveForm.RTF1.SelStart = pos - 1
  frmMain.ActiveForm.RTF1.SelText = "<BGSOUND src=(Wait...)>" & vbCrLf
  pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<bgsound", vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
End If
If txBGSOUNDsrc.Text = "" Then GoTo nxt2
Dim strLoop As String
strLoop = IIf(chLoop.Value = 1, " loop=" & Chr(34) & "-1" & Chr(34) & " ", "")
frmMain.ActiveForm.RTF1.SelStart = pos - 1
frmMain.ActiveForm.RTF1.SelLength = pos2 - pos + 1
frmMain.ActiveForm.RTF1.SelText = "<BGSOUND src=" & Chr(34) & txBGSOUNDsrc.Text & Chr(34) & strLoop & ">"
nxt2:
'set BASE target
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<base", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
If (pos = 0 Or pos2 = 0) And cbBaseTarget.Text <> "" Then
  pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<HEAD>" & vbCrLf, vbTextCompare) + Len("<HEAD>" & vbCrLf)
  If pos = Len("<HEAD>" & vbCrLf) Then pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<HEAD>", vbTextCompare) + Len("<HEAD>")
  If pos = Len("<HEAD>") Then pos = 1
  frmMain.ActiveForm.RTF1.SelStart = pos - 1
  frmMain.ActiveForm.RTF1.SelText = "<BASE target=(Wait...)>" & vbCrLf
  pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<base", vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
End If
pos2 = pos2 + 1 '1=len(">")
whole = Mid$(frmMain.ActiveForm.RTF1.Text, pos, pos2 - pos)
If cbBaseTarget.Text = "" Then DelAttrib whole, "target", pos Else SaveAttrib "target", cbBaseTarget.Text, "<BASE", ">"
pB.Value = pB.Value + 1
'set body things. Complicated...
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<body", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
If pos = 0 Or pos2 = 0 Then
  pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "</HTML>", vbTextCompare)
  If pos = 0 Then pos = Len(frmMain.ActiveForm.RTF1.Text)
  frmMain.ActiveForm.RTF1.SelStart = pos - 1
  frmMain.ActiveForm.RTF1.SelText = "<BODY>" & vbCrLf & vbCrLf & "</BODY>" & vbCrLf
  pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<body", vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">", vbTextCompare)
End If
pos2 = pos2 + 1
whole = Mid$(frmMain.ActiveForm.RTF1.Text, pos, pos2 - pos)
'whole is now the body tag (w/ attributes etc.)
'If any textbox is empty, delete the attribute in the tag if it already
'exists. Else just don't add.
If txColNorm.Text = "" Then DelAttrib whole, "text", pos Else SaveAttrib "text", txColNorm.Text
pB.Value = pB.Value + 1
If txBGCol.Text = "" Then DelAttrib whole, "bgcolor", pos Else SaveAttrib "bgcolor", txBGCol.Text
pB.Value = pB.Value + 1
If txBGImage.Text = "" Then DelAttrib whole, "background", pos Else SaveAttrib "background", txBGImage.Text
pB.Value = pB.Value + 1
If chWatermark.Value = 0 Then DelAttrib whole, "bgproperties", pos Else SaveAttrib "bgproperties", "fixed"
pB.Value = pB.Value + 1
If txColLink.Text = "" Then DelAttrib whole, "link", pos Else SaveAttrib "link", txColLink.Text
pB.Value = pB.Value + 1
If txColVlink.Text = "" Then DelAttrib whole, "vlink", pos Else SaveAttrib "vlink", txColVlink.Text
pB.Value = pB.Value + 1
If txColActive.Text = "" Then DelAttrib whole, "alink", pos Else SaveAttrib "alink", txColActive.Text
pB.Value = pB.Value + 1
If txMrgnLeft.Text = "" Then DelAttrib whole, "leftmargin", pos Else SaveAttrib "leftmargin", txMrgnLeft.Text
pB.Value = pB.Value + 1
If txMrgnTop.Text = "" Then DelAttrib whole, "topmargin", pos Else SaveAttrib "topmargin", txMrgnTop.Text
pB.Value = pB.Value + 1
If txMrgnHeight.Text = "" Then DelAttrib whole, "marginheight", pos Else SaveAttrib "marginheight", txMrgnHeight.Text
pB.Value = pB.Value + 1
If txMrgnWidth.Text = "" Then DelAttrib whole, "marginwidth", pos Else SaveAttrib "marginwidth", txMrgnWidth.Text
pB.Value = pB.Value + 1
End Sub

Function SaveAttrib(ID As String, Value As String, Optional tagStart As String = "<BODY", Optional tagEnd As String = ">") As String
On Error Resume Next
Dim pos As Long, pos2 As Long
Dim lStart As Long
Dim Where As String
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, tagStart, vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, tagEnd, vbTextCompare)
If pos = 0 Or pos2 = 0 Then Exit Function
lStart = pos 'save the position of tag in lstart
Where = Mid$(frmMain.ActiveForm.RTF1.Text, pos, pos2 + 1 - pos)
'Where contains the BODY tag now
pos = 0: pos2 = 0
'every time this function is called, a new BODY tag has to be
'calculated as a previous call might have altered it.
'where contains the entire body tag e.g. <BODY a="b" c="d">
pos = InStr(1, Where, " " & ID & "=", vbTextCompare)
'find attrib, e.g. attrib 'name' is searched for as " name="
If pos = 0 Then GoTo not_existing Else pos = pos + Len(ID) + 2 'skip over the id and get the value part
pos2 = InStr(pos + 1, Where, " ") - 1 'it'll be -1 if no space exists
'by default, value ends at the next space. If it is enclosed in " or ', calculate
'the position of the closing " or '.
If Mid$(Where, pos, 1) = Chr(34) Then pos2 = InStr(pos + 1, Where, Chr(34))
If Mid$(Where, pos, 1) = "'" Then pos2 = InStr(pos + 1, Where, "'")
If pos2 = -1 Then pos2 = Len(Where) Else pos2 = pos2 + 1
frmMain.ActiveForm.RTF1.SelStart = lStart + pos - 2
frmMain.ActiveForm.RTF1.SelLength = pos2 - pos
frmMain.ActiveForm.RTF1.SelText = Chr(34) & Value & Chr(34)
Exit Function
not_existing:
frmMain.ActiveForm.RTF1.SelStart = lStart + Len(Where) - 2 'to end of string
frmMain.ActiveForm.RTF1.SelText = " " & ID & "=" & Chr(34) & Value & Chr(34)
End Function

Function DelAttrib(Where As String, ID As String, StartPos As Long) As String
On Error Resume Next
Dim pos As Long, pos2 As Long
pos = InStr(1, Where, " " & ID & "=", vbTextCompare)
If pos = 0 Then Exit Function
pos2 = InStr(pos + 1, Where, " ")
If Mid$(Where, pos + Len(ID) + 3, 1) = Chr(34) Then pos2 = InStr(pos + 1, Where, Chr(34))
If Mid$(Where, pos + Len(ID) + 3, 1) = "'" Then pos2 = InStr(pos + 1, Where, "'")
If pos2 = 0 Then pos2 = Len(Where)
frmMain.ActiveForm.RTF1.SelStart = StartPos + pos - 2
frmMain.ActiveForm.RTF1.SelLength = pos2 - pos
frmMain.ActiveForm.RTF1.SelText = ""
End Function

Sub LoadDTCInfo()
On Error GoTo nxt
Dim pos As Long, pos2 As Long, tmp As String
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<% Language=", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, "%>")
If pos = 0 Or pos2 = 0 Then GoTo nxt Else pos = pos + Len("<% Language=")
tmp = Mid$(frmMain.ActiveForm.RTF1.Text, pos, pos2 - pos)
tmp = Trim(LCase(tmp))
If tmp = "javascript" Then
cbDTC_Server.ListIndex = 1
ElseIf tmp = "vbscript" Then
cbDTC_Server.ListIndex = 2
Else
nxt: 'default scripting server language
cbDTC_Server.ListIndex = 0
End If
On Error GoTo nxt2
If LCase(lvMETA.ListItems("dtc_client").ListSubItems(1).Text) = "vbscript" Then
cbDTC_Client.ListIndex = 2
ElseIf LCase(lvMETA.ListItems("dtc_client").ListSubItems(1).Text) = "javascript" Then
cbDTC_Client.ListIndex = 1
End If
nxt2:
On Error GoTo hell
If LCase(lvMETA.ListItems("dtc_platform").ListSubItems(1).Text) = "server (asp)" Then
cbDTC_Platform.ListIndex = 2
ElseIf LCase(lvMETA.ListItems("dtc_platform").ListSubItems(1).Text) = "client (ie 4.0 dhtml)" Then
cbDTC_Platform.ListIndex = 1
End If
hell:
End Sub

Sub SetDTCInfo()
On Error Resume Next
Dim bCreatedOurselves As Boolean
If cbDTC_Client.ListIndex = 0 Then lvMETA.ListItems.Remove "dtc_client":  GoTo nxt1
lvMETA.ListItems.Add , "dtc_client", "VI60_defaultClientScript", , 1
lvMETA.ListItems("dtc_client").Ghosted = True
lvMETA.ListItems("dtc_client").ListSubItems.Add 1
lvMETA.ListItems("dtc_client").ListSubItems(1).Text = cbDTC_Client.Text
lvMETA.ListItems("dtc_client").ListSubItems.Add 2
lvMETA.ListItems("dtc_client").ListSubItems(2).Text = "name"
nxt1:
If cbDTC_Platform.ListIndex = 0 Then lvMETA.ListItems.Remove "dtc_platform": GoTo nxt2
lvMETA.ListItems.Add , "dtc_platform", "VI60_DTCScriptingPlatform", , 1
lvMETA.ListItems("dtc_platform").Ghosted = True
lvMETA.ListItems("dtc_platform").ListSubItems.Add 1
lvMETA.ListItems("dtc_platform").ListSubItems(1).Text = cbDTC_Platform.Text
lvMETA.ListItems("dtc_platform").ListSubItems.Add 2
lvMETA.ListItems("dtc_platform").ListSubItems(2).Text = "name"
nxt2:
'OK the fun starts here. We now need to manipulate the server script indicator
'manually. Let's see how it goes through...
Dim pos As Long, pos2 As Long, tmp As String
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<% Language=", vbTextCompare)
bCreatedOurselves = False
If pos = 0 Then
frmMain.ActiveForm.RTF1.Text = "<% Language=(Default) %>" & vbCrLf & vbCrLf & frmMain.ActiveForm.RTF1.Text
pos = InStr(1, frmMain.ActiveForm.RTF1.Text, "<% Language=", vbTextCompare)
bCreatedOurselves = True
End If
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, "%>")
frmMain.ActiveForm.RTF1.SelStart = pos - 1
frmMain.ActiveForm.RTF1.SelLength = pos2 - pos + IIf(bCreatedOurselves, Len("%>" & vbCrLf & vbCrLf), Len("%>"))
If cbDTC_Server.ListIndex = 0 Then tmp = ""
If cbDTC_Server.ListIndex = 1 Then tmp = "<% Language=JavaScript %>" & vbCrLf & vbCrLf
If cbDTC_Server.ListIndex = 2 Then tmp = "<% Language=VBScript %>" & vbCrLf & vbCrLf
frmMain.ActiveForm.RTF1.SelText = tmp
End Sub

Sub GetFrames()
Dim whole As String, s As String
Dim pos As Long, pos2 As Long
Do
pos = InStr(pos2 + 1, frmMain.ActiveForm.RTF1.Text, "<FRAME", vbTextCompare)
pos2 = InStr(pos + 1, frmMain.ActiveForm.RTF1.Text, ">")
If pos = 0 Or pos2 = 0 Then Exit Do
whole = Mid$(frmMain.ActiveForm.RTF1.Text, pos, pos2 - pos)
s = ReadAttrib("name", whole)
If s <> whole Then cbBaseTarget.AddItem s
Loop
End Sub
