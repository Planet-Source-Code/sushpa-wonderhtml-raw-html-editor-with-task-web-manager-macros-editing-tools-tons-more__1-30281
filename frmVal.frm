VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validation"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   Icon            =   "frmVal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmEdit 
      Caption         =   "&Edit tag"
      Height          =   375
      Left            =   68
      TabIndex        =   2
      Top             =   1680
      Width           =   1050
   End
   Begin VB.CommandButton cmOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2723
      TabIndex        =   0
      Top             =   1680
      Width           =   1050
   End
   Begin MSComctlLib.ImageList imlProbs 
      Left            =   113
      Top             =   1972
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVal.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVal.frx":06EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvProbs 
      Height          =   1635
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   2884
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlProbs"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Problem"
         Object.Width           =   4560
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tag"
         Object.Width           =   1641
      EndProperty
   End
End
Attribute VB_Name = "frmVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmEdit_Click()
On Error Resume Next
frmMain.ActiveForm.RTF1.SelStart = CLng(lvProbs.SelectedItem.Tag) - 1
frmMain.ActiveForm.RTF1.SetFocus
Me.Move 0, 0
End Sub

Private Sub cmOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetFont Me
SetWindowPos hWnd, -1, Left \ 15, Top \ 15, Width \ 15, Height \ 15, 0
End Sub

Sub ListLinksInDocWhileValidating(Text As String, File As String)
On Error Resume Next

Dim this As String, href As String

Dim l As Long, l2 As Long
Do
l = InStr(l2 + 1, Text, "<a ", vbTextCompare)
If l = 0 Then Exit Do
l2 = InStr(l + 1, Text, ">", vbTextCompare)
If l2 = 0 Then Exit Do
this = Mid$(Text, l, l2 - l + 1)
href = ReadAttrib("href", this)
If href = this Then GoTo nxt 'not found
href = Replace(href, "%20", " ")
If IsLink(href) Then GoTo nxt 'is an external link
If Left(href, 1) = "#" Then GoTo nxt 'anchors, bookmarks
ChDrive Left(File, 2)
ChDir Up1Level(File, "\")
ChDir Up1Level(href, "/")
href = FullPath(CurDir(Left(File, 2)), GetFile(href))
If Dir(href) <> "" Then GoTo nxt 'functional link
lvProbs.ListItems.Add , "tmp", "Broken Hyperlink", , 2
lvProbs.ListItems("tmp").ListSubItems.Add 1, , "<A>"
lvProbs.ListItems("tmp").Tag = l
lvProbs.ListItems("tmp").ListSubItems(1).Tag = "Points to " & href
lvProbs.ListItems("tmp").Key = ""
nxt:
Loop
l = 0: l2 = 0
'find IMG sources now
Do
l = InStr(l2 + 1, Text, "<img ", vbTextCompare)
If l = 0 Then Exit Do
l2 = InStr(l + 1, Text, ">", vbTextCompare)
If l2 = 0 Then Exit Do
this = Mid$(Text, l, l2 - l + 1)
href = ReadAttrib("src", this)
If href = this Then GoTo nxt2 'not found
href = Replace(href, "%20", " ")
If IsLink(href) Then GoTo nxt2 'is an external link
ChDrive Left(File, 2)
ChDir Up1Level(File, "\")
ChDir Up1Level(href, "/")
href = FullPath(CurDir(Left(File, 2)), GetFile(href))
If Dir(href) <> "" Then GoTo nxt2 'functional link
lvProbs.ListItems.Add , "tmp", "Broken Image source", , 2
lvProbs.ListItems("tmp").ListSubItems.Add 1, , "<IMG>"
lvProbs.ListItems("tmp").Tag = l
lvProbs.ListItems("tmp").ListSubItems(1).Tag = "Points to " & href
lvProbs.ListItems("tmp").Key = ""
nxt2:
Loop
End Sub

Private Sub lvProbs_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
frmMain.SB.Panels(1).Text = Item.ListSubItems(1).Tag
End Sub
