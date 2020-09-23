VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRept 
   AutoRedraw      =   -1  'True
   Caption         =   "WonderHTML: Viewing Report"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   Icon            =   "frmRept.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4770
      Left            =   0
      ScaleHeight     =   4770
      ScaleWidth      =   60
      TabIndex        =   5
      Top             =   0
      Width           =   60
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   4770
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4234
            MinWidth        =   4234
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5715
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   " OK: 0 "
            TextSave        =   " OK: 0 "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   " Err: 0 "
            TextSave        =   " Err: 0 "
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1085
            MinWidth        =   1058
            Text            =   " Ext: 0 "
            TextSave        =   " Ext: 0 "
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pBack 
      BackColor       =   &H80000004&
      Height          =   4695
      Left            =   945
      ScaleHeight     =   4635
      ScaleWidth      =   6750
      TabIndex        =   1
      Top             =   45
      Width           =   6810
      Begin MSComctlLib.ImageList imlTB 
         Left            =   540
         Top             =   3690
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":0E42
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":1B1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":27FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":34D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":535A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":6036
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":6D12
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":702E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":8EB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":AD36
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRept.frx":CBBA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   45
         Picture         =   "frmRept.frx":DC0E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Top             =   15
         Width           =   480
      End
      Begin MSComctlLib.ListView lvRept 
         Height          =   4095
         Left            =   0
         TabIndex        =   0
         Top             =   540
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   7223
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         SmallIcons      =   "imlTV"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Modified"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Page Title"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label lbWeb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Web"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   585
         TabIndex        =   2
         Top             =   90
         Width           =   1005
      End
   End
   Begin MSComctlLib.ImageList imlTV 
      Left            =   360
      Top             =   4050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":10908
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":1195C
            Key             =   "otherfile"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":137E0
            Key             =   "audio"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":15664
            Key             =   "program"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":166B8
            Key             =   "shellscript"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":1770C
            Key             =   "script"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":17AA8
            Key             =   "winword"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":18AFC
            Key             =   "image"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":19950
            Key             =   "Unknown"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":19EEC
            Key             =   "Broken"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":1A488
            Key             =   "pdf"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":1B4DC
            Key             =   "psd"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":1C530
            Key             =   "archive"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":1C8CC
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRept.frx":1CE68
            Key             =   "css"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   3  'Align Left
      Height          =   4770
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   8414
      ButtonWidth     =   1376
      ButtonHeight    =   1376
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      ImageList       =   "imlTB"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Files   "
            Object.ToolTipText     =   "Files & File Information"
            ImageIndex      =   11
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Links"
            Object.ToolTipText     =   "Scan for doubtful links"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Large   "
            Object.ToolTipText     =   "View files that are large in size"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Media  "
            Object.ToolTipText     =   "Images, Audio/Video, Programs"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   600
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFileS 
      Caption         =   "&Files"
      Visible         =   0   'False
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Edit page"
      End
      Begin VB.Menu mnuFileDel 
         Caption         =   "&Delete..."
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "&File Info"
      End
   End
   Begin VB.Menu mnuLinks 
      Caption         =   "&Links"
      Visible         =   0   'False
      Begin VB.Menu mnuEditPage 
         Caption         =   "&Edit page"
      End
      Begin VB.Menu mnuOpenLink 
         Caption         =   "L&aunch"
      End
   End
End
Attribute VB_Name = "frmRept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportMode As Long
Dim LinksOK As Long, LinksErr As Long, LinksUnknown As Long

Sub InitializeRept(Web As String)
On Error Resume Next
ArrangeLV 1
lbWeb.Caption = Web
Caption = "Report for " & Web
TB_ButtonClick TB.Buttons(2)
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
lbWeb.FontName = "Tahoma"
lbWeb.FontSize = 14
TB.Style = ReadValue("FlatBar")
lvRept.GridLines = ReadValue("Gridlines", False, "Reports")
End Sub

Private Sub Form_Resize()
On Error Resume Next
TB.Visible = False
pBack.Width = ScaleWidth - pBack.Left - 45
pBack.Height = ScaleHeight - pBack.Top - SB.Height - 45
lvRept.Height = pBack.ScaleHeight - lvRept.Top
lvRept.Width = pBack.ScaleWidth
TB.Visible = True
End Sub

Private Sub lvRept_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvRept.SortKey = ColumnHeader.Index - 1
lvRept.Sorted = True
End Sub

Private Sub lvRept_ItemClick(ByVal Item As MSComctlLib.ListItem)
SB.Panels(2).Text = Item.Key
End Sub

Private Sub lvRept_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If lvRept.ListItems.count = 0 Or Button <> 2 Then Exit Sub
If lvRept.SelectedItem Is Nothing Then lvRept.SelectedItem = lvRept.ListItems(1)
Select Case ReportMode
Case 1, 3, 4
PopupMenu mnuFileS
Case 2
PopupMenu mnuLinks
End Select
End Sub

Private Sub mnuEditPage_Click()
Dim l As New frmChild
Load l
l.LoadHTMLFile FullPath(lbWeb.Caption, lvRept.SelectedItem.ListSubItems(1).Text)
l.RTF1.SetFocus
End Sub

Private Sub mnuFileDel_Click()
On Error Resume Next
Dim s As Boolean, F As String
F = FullPath(lbWeb.Caption, lvRept.SelectedItem.Text)
MousePointer = 11
s = DeleteFile(F)
If s = True Then
frmMain.tvW.Nodes.Remove lvRept.SelectedItem.Key
lvRept.ListItems.Remove lvRept.SelectedItem.Index
End If
MousePointer = 0
End Sub

Private Sub mnuFileInfo_Click()
FileInfo FullPath(lbWeb.Caption, lvRept.SelectedItem.Text)
End Sub

Private Sub mnuFileOpen_Click()
Dim lpF As New frmChild
Load lpF
lpF.LoadHTMLFile lvRept.SelectedItem.Key
End Sub

Private Sub mnuOpenLink_Click()
Dim URL As String
URL = FullPath(lbWeb.Caption, lvRept.SelectedItem.ListSubItems(2).Text)
URL = Replace(URL, "/", "\")
ShellExecute hwnd, "open", URL, "", "", 10
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "Exit" Then Unload Me
lvRept.ListItems.Clear
ArrangeLV Button.Index - 1
'index-1, cause first is separator
ReportMode = Button.Index - 1
If ReportMode = 2 Then LinksOK = 0: LinksErr = 0: LinksUnknown = 0
frmMain.Fldr.Path = frmMain.tvW.Nodes(1).Key
DirDiver frmMain.CurrentWeb, frmMain.Fldr.Path, True, ACTION_REPORT, ReportMode, "", vbBinaryCompare 'dummy
SB.Panels(2).Text = GetFinalText(Button.Index - 1)
End Sub

Sub ListLinksInDoc(File As String)
On Error Resume Next

Dim s As String, this As String
Dim count As Long, lRepImg As String
Dim href As String, orgHref As String


If Ext(File) <> "html" And Ext(File) <> "htm" Then Exit Sub
SB.Panels(2).Text = File
Open File For Binary Access Read As #1
s = Space(LOF(1))
Get #1, , s
Close #1
Dim l As Long, l2 As Long
Do
l = InStr(l2 + 1, s, "<a ", vbTextCompare)
If l = 0 Then Exit Do
l2 = InStr(l + 1, s, ">", vbTextCompare)
If l2 = 0 Then Exit Do
this = Mid$(s, l, l2 - l + 1)
href = ReadAttrib("href", this)
If href = this Then GoTo n
orgHref = href 'original loc as specified
count = count + 1
href = Replace(href, "%20", " ")
If IsLink(href) Then lRepImg = "Unknown": LinksUnknown = LinksUnknown + 1: GoTo start
ChDrive Left(lbWeb.Caption, 2)
ChDir Up1Level(File, "\")
ChDir Up1Level(href, "/")
If Left(href, 1) = "#" Then GoTo n 'anchors, bookmarks
href = FullPath(CurDir(Left(lbWeb.Caption, 2)), GetFile(href))
If Dir(href) = "" Then
  lRepImg = "Broken"
  LinksErr = LinksErr + 1
Else
  lRepImg = "Active"
  LinksOK = LinksOK + 1
  If ReadValue("BrokenOnly", True, "Reports") = True Then GoTo n
End If
start:
href = Replace(href, "%20", " ")
If Not IsLink(href) Then
  href = FirstSlash(Replace(href, lbWeb.Caption, "", , , vbTextCompare))
  href = Replace(href, "\", "/")
End If
If href = this Then href = ""
If href <> "" Then
  lvRept.ListItems.Add , "az", lRepImg, , lRepImg
  lvRept.ListItems("az").ListSubItems.Add 1, , FirstSlash(Replace(File, lbWeb.Caption, ""))
  lvRept.ListItems("az").ListSubItems.Add 2, , href
  lvRept.ListItems("az").ListSubItems(2).Tag = orgHref
'  lvRept.ListItems("az").EnsureVisible
  lvRept.ListItems("az").Key = ""
End If
n:
Loop
l = 0: l2 = 0
Do
l = InStr(l2 + 1, s, "<img ", vbTextCompare) 'image sources
If l = 0 Then Exit Do
l2 = InStr(l + 1, s, ">", vbTextCompare)
If l2 = 0 Then Exit Do
this = Mid$(s, l, l2 - l + 1)
href = ReadAttrib("src", this)
If href = this Then GoTo n2
orgHref = href
count = count + 1
If IsLink(href) Then lRepImg = "Unknown": LinksUnknown = LinksUnknown + 1: GoTo start2
ChDrive Left(lbWeb.Caption, 2)
ChDir Up1Level(File, "\")
ChDir Up1Level(href, "/")
href = FullPath(CurDir(Left(lbWeb.Caption, 2)), GetFile(href))
If Dir(href) = "" Then
  lRepImg = "Broken"
  LinksErr = LinksErr + 1
Else
  lRepImg = "Active"
  LinksOK = LinksOK + 1
  If ReadValue("BrokenOnly", True, "Reports") = True Then GoTo n2
End If
start2:
If Not IsLink(href) Then
  href = FirstSlash(Replace(href, lbWeb.Caption, "", , , vbTextCompare))
  href = Replace(href, "\", "/")
End If
If href = this Then href = ""
If href <> "" Then
  lvRept.ListItems.Add , "az", lRepImg, , lRepImg
  lvRept.ListItems("az").ListSubItems.Add 1, , FirstSlash(Replace(File, lbWeb.Caption, ""))
  lvRept.ListItems("az").ListSubItems.Add 2, , href
  lvRept.ListItems("az").ListSubItems(2).Tag = orgHref
'  lvRept.ListItems("az").EnsureVisible
  lvRept.ListItems("az").Key = ""
End If
n2:
Loop
SB.Panels(3).Text = "OK: " & LinksOK & " "
SB.Panels(4).Text = "Err: " & LinksErr & " "
SB.Panels(5).Text = "Ext: " & LinksUnknown & " "
End Sub

Sub DoReportAction(File As String, Action As Long)
On Error Resume Next
If Ext(File) = "wbackup" Then Exit Sub
Select Case Action
Case 1
AddFileInfo File, False
Case 2
ListLinksInDoc File
Case 3
AddFileInfo File, True
Case 4
AddMediaFile File
End Select
End Sub

Sub AddFileInfo(ThePath As String, bOnlyLargeFiles As Boolean)
On Error Resume Next
If GetFile(ThePath) = "files.inf" Then Exit Sub
If bOnlyLargeFiles Then
  If FileLen(ThePath) <= GetSizeLimit() Then Exit Sub
End If
lvRept.ListItems.Add 1, ThePath, FirstSlash(Replace(ThePath, lbWeb.Caption, "")), , FileIcon(ThePath)
lvRept.ListItems(ThePath).ListSubItems.Add 1, , Round((FileLen(ThePath) / 1024), 1) & " KB"
lvRept.ListItems(ThePath).ListSubItems.Add 2, , Format(FileDateTime(ThePath), "ddd, dd mmm yyyy")
lvRept.ListItems(ThePath).ListSubItems.Add 3, , IIf(bOnlyLargeFiles, Round((FileLen(ThePath) - GetSizeLimit) / 1024, 1) & " KB", GetTitle(ThePath))
'lvRept.ListItems(ThePath).EnsureVisible
SB.Panels(2).Text = ThePath
End Sub

Sub ArrangeLV(ForWhat As Long)
With lvRept
  Select Case ForWhat
    Case 1, 3
      .ColumnHeaders(1).Width = 2100
      .ColumnHeaders(1).Text = "File"
      .ColumnHeaders(2).Width = 1000
      .ColumnHeaders(2).Text = "Size"
      .ColumnHeaders(3).Width = 1600
      .ColumnHeaders(3).Text = "Modified"
      .ColumnHeaders(4).Text = IIf(ForWhat = 1, "Description", "Exceeds by")
      .ColumnHeaders(4).Width = 1800
      .Sorted = False
    Case 2
      .ColumnHeaders(1).Width = 1000
      .ColumnHeaders(1).Text = "Status"
      .ColumnHeaders(2).Width = 2025
      .ColumnHeaders(2).Text = "Page"
      .ColumnHeaders(3).Width = 3450
      .ColumnHeaders(3).Text = "Target"
      .ColumnHeaders(4).Width = 0
      .Sorted = True
      Case 4
      .ColumnHeaders(1).Width = 2100
      .ColumnHeaders(1).Text = "File"
      .ColumnHeaders(2).Width = 1000
      .ColumnHeaders(2).Text = "Size"
      .ColumnHeaders(3).Width = 1600
      .ColumnHeaders(3).Text = "Modified"
      .ColumnHeaders(4).Text = "Type"
      .ColumnHeaders(4).Width = 1800
      .Sorted = False
  End Select
End With
End Sub

Function GetFinalText(Operation As Long)
Select Case Operation
Case 1
GetFinalText = lvRept.ListItems.count & " files in the web."
Case 2
GetFinalText = "Found " & ParseInt(SB.Panels(3).Text) + ParseInt(SB.Panels(4).Text) + ParseInt(SB.Panels(5).Text) & " links."
Case 3
GetFinalText = lvRept.ListItems.count & " oversized files."
Case 4
GetFinalText = lvRept.ListItems.count & " media-related files."
End Select
End Function

Sub AddMediaFile(ThePath As String)
On Error Resume Next
If GetFile(ThePath) = "files.inf" Then Exit Sub
Dim st As String
st = ReadValue("Media", "", "Reports")
If Left(st, 1) <> " " Then st = " " & st
If Right(st, 1) <> " " Then st = st & " "
If st = "" Then st = " jpg gif bmp png jpeg mpeg exe zip sit hqx rar cab ocx "
If InStr(1, st, " " & Ext(ThePath) & " ") > 0 Then
  lvRept.ListItems.Add 1, ThePath, FirstSlash(Replace(ThePath, lbWeb.Caption, "")), , FileIcon(ThePath)
  lvRept.ListItems(ThePath).ListSubItems.Add 1, , Round((FileLen(ThePath) / 1024), 1) & " K"
  lvRept.ListItems(ThePath).ListSubItems.Add 2, , Format(FileDateTime(ThePath), "ddd, dd mmm yyyy")
  lvRept.ListItems(ThePath).ListSubItems.Add 3, , FileType(ThePath)
  'lvRept.ListItems(ThePath).EnsureVisible
  SB.Panels(2).Text = ThePath
End If
End Sub

Private Sub TB_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Static HoverIndex As Long
HoverIndex = (y - TB.Buttons(1).Height) \ TB.ButtonHeight
SB.Panels(1).Text = TB.Buttons(HoverIndex + 2).ToolTipText
End Sub
