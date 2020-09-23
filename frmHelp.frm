VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHelp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WonderHTML Help Contents"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   5
      Top             =   495
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Contents"
      TabPicture(0)   =   "frmHelp.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sTree"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "frmHelp.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "txFind"
      Tab(1).Control(3)=   "opFText"
      Tab(1).Control(4)=   "opFNode"
      Tab(1).Control(5)=   "chkMC"
      Tab(1).Control(6)=   "chWords"
      Tab(1).Control(7)=   "chSearch"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton chSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Begin &Search"
         Height          =   375
         Left            =   -73830
         MouseIcon       =   "frmHelp.frx":0182
         TabIndex        =   12
         Top             =   2250
         Width           =   1275
      End
      Begin VB.CheckBox chWords 
         Caption         =   "Whole &words only"
         Height          =   195
         Left            =   -74910
         TabIndex        =   11
         Top             =   1935
         Width           =   1770
      End
      Begin VB.CheckBox chkMC 
         Caption         =   "Match &Case"
         Height          =   195
         Left            =   -74910
         TabIndex        =   10
         Top             =   1710
         Width           =   1770
      End
      Begin VB.OptionButton opFNode 
         Caption         =   "Find in &Index"
         Height          =   240
         Left            =   -74910
         TabIndex        =   9
         Top             =   1365
         Width           =   2040
      End
      Begin VB.OptionButton opFText 
         Caption         =   "Find in &Text"
         Height          =   240
         Left            =   -74910
         TabIndex        =   8
         Top             =   1125
         Value           =   -1  'True
         Width           =   2040
      End
      Begin VB.TextBox txFind 
         Height          =   315
         Left            =   -74910
         TabIndex        =   7
         Top             =   630
         Width           =   2400
      End
      Begin MSComctlLib.TreeView sTree 
         Height          =   3465
         Left            =   90
         TabIndex        =   0
         Top             =   405
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   6112
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   317
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   1
         SingleSel       =   -1  'True
         ImageList       =   "imlHelp"
         Appearance      =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "WonderHTML™ Help System      © Sushant Pandurangi, 2001."
         Height          =   510
         Left            =   -74775
         TabIndex        =   13
         Top             =   3465
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Find what:"
         Height          =   195
         Left            =   -74910
         TabIndex        =   6
         Top             =   405
         Width           =   765
      End
   End
   Begin VB.CommandButton cmCopy 
      Cancel          =   -1  'True
      Caption         =   "C&opy All Text"
      Height          =   375
      Left            =   1485
      TabIndex        =   3
      Top             =   45
      Width           =   1365
   End
   Begin VB.CommandButton cmClose 
      Caption         =   "&Close Help"
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   1365
   End
   Begin MSComctlLib.ImageList imlHelp 
      Left            =   2745
      Top             =   2947
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":0A4C
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":0FE8
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":1584
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfHelp 
      Height          =   3930
      Left            =   2655
      TabIndex        =   1
      Top             =   495
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   6932
      _Version        =   393217
      BackColor       =   -2147483624
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmHelp.frx":1B20
      MouseIcon       =   "frmHelp.frx":1C2E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the Help System."
      Height          =   195
      Left            =   2970
      TabIndex        =   4
      Top             =   135
      UseMnemonic     =   0   'False
      Width           =   2115
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chSearch_Click()
Select Case True 'which one is true?
Case opFNode.Value
FindNodeString
Case opFText.Value
If rtfHelp.Find(txFind.Text, , , IIf(chkMC.Value = 1, rtfMatchCase, 0) + IIf(chWords.Value = 1, rtfWholeWord, 0)) = -1 Then lbInfo.Caption = "Cannot find instances of " & txFind.Text
End Select
End Sub

Private Sub cmClose_Click()
Unload Me
End Sub

Private Sub cmCopy_Click()
Clipboard.Clear
Clipboard.SetText rtfHelp.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
rtfHelp.LoadFile FullPath(App.Path, "Help\Main.rtf"), rtfRTF
LoadTheTreeView
End Sub

Sub LoadTheTreeView()
On Error Resume Next
Dim s As String, st() As String, i As Long
Dim thisone() As String
Open FullPath(App.Path, "HelpIndex.dat") For Input As #1
s = Input(LOF(1), 1)
Close #1
s = Replace(s, Chr(10), "")
st = Split(s, Chr(13))
For i = 0 To UBound(st)
If st(i) = "" Then GoTo n
thisone = Split(st(i), vbTab)
If thisone(0) = "" Then
sTree.Nodes.Add , , thisone(1), thisone(2), thisone(3)
Else
sTree.Nodes.Add thisone(0), tvwChild, thisone(1), thisone(2), thisone(3)
End If
If thisone(4) <> "" Then sTree.Nodes(thisone(1)).ExpandedImage = thisone(4)
Next i
n: 'next line
For i = 1 To sTree.Nodes.count
'sTree.Nodes(i).Expanded = True
Next i
sTree.SelectedItem = sTree.Nodes(1)
End Sub

Private Sub opFNode_Click()
chkMC.Enabled = False
chWords.Enabled = False
End Sub

Private Sub opFText_Click()
chkMC.Enabled = True
chWords.Enabled = True
End Sub

Private Sub rtfHelp_Click()
Dim l As Long, l2 As Long
If rtfHelp.SelText = "" Or IsLink(rtfHelp.Text) = False Then Exit Sub
l = rtfHelp.SelStart
l2 = rtfHelp.SelLength
ShellExecute hwnd, "open", rtfHelp.SelText, "", "", 10
rtfHelp.SelStart = l
rtfHelp.SelLength = l2
End Sub

Private Sub rtfHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Static s As String, l As Long
If s = RichWordOver(rtfHelp, X, y) Then Exit Sub
s = RichWordOver(rtfHelp, X, y, l)
If IsLink(s) Then
rtfHelp.MousePointer = 99
rtfHelp.SelStart = l - 1
rtfHelp.SelLength = Len(s)
Else
rtfHelp.MousePointer = 0
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then sTree.SetFocus Else txFind.SetFocus
End Sub

Private Sub sTree_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Image = "item" Or Node.Image = 3 Then ShowTopic Node.Key
End Sub

Sub ShowTopic(File As String)
On Error GoTo hell
rtfHelp.LoadFile FullPath(App.Path, "Help\" & File), rtfRTF
Exit Sub
hell:
lbInfo.Caption = File & ": The topic does not exist."
End Sub

Sub FindNodeString()
Dim i As Long
For i = 1 To sTree.Nodes.count
If InStr(1, sTree.Nodes(i).Text, txFind.Text, vbTextCompare) > 0 Then sTree.SelectedItem = sTree.Nodes(i): sTree.SetFocus: SSTab1.Tab = 0: Exit Sub
Next i
lbInfo.Caption = "Cannot find " & txFind.Text & " in the index."
End Sub
