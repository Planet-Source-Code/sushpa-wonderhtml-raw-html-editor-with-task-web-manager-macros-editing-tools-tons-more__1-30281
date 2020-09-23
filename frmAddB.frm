VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddB 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Browser"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmAddB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageList imlImg 
      Left            =   1170
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddB.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmCanc 
      Cancel          =   -1  'True
      Caption         =   "'so that Esc works"
      Height          =   420
      Left            =   1800
      TabIndex        =   7
      Top             =   5000
      Width           =   1770
   End
   Begin VB.PictureBox pp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   180
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   1980
      Width           =   480
   End
   Begin VB.CommandButton cmClose 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   3870
      TabIndex        =   5
      Top             =   2160
      Width           =   1230
   End
   Begin VB.CommandButton cmBrowse 
      Caption         =   "Sel&ect..."
      Height          =   375
      Left            =   3870
      TabIndex        =   4
      Top             =   1350
      Width           =   1230
   End
   Begin VB.CommandButton cmEdit 
      Caption         =   "E&dit path"
      Height          =   375
      Left            =   3870
      TabIndex        =   3
      Top             =   945
      Width           =   1230
   End
   Begin VB.CommandButton cmRem 
      Caption         =   "R&emove..."
      Height          =   375
      Left            =   3870
      TabIndex        =   2
      Top             =   540
      Width           =   1230
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "A&dd New..."
      Height          =   375
      Left            =   3870
      TabIndex        =   1
      Top             =   135
      Width           =   1230
   End
   Begin MSComctlLib.ListView lvB 
      Height          =   1740
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3069
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlImg"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   6085
      EndProperty
   End
End
Attribute VB_Name = "frmAddB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmAdd_Click()
On Error Resume Next
lvB.ListItems.Add , "tmp", , , 1
lvB.SelectedItem = lvB.ListItems("tmp")
lvB.ListItems("tmp").Key = ""
lvB.SetFocus
lvB.StartLabelEdit
End Sub

Private Sub cmBrowse_Click()
On Error GoTo hell
Dim s As String
s = frmMain.CD.Filter
frmMain.CD.Filter = "Executable applications (*.exe)|*.exe"
frmMain.CD.ShowOpen
lvB.ListItems.Add , , frmMain.CD.Filename, , 1
hell:
End Sub

Private Sub cmCanc_Click()
Unload Me
End Sub

Private Sub cmClose_Click()
On Error Resume Next
Dim i As Integer
For i = 1 To lvB.ListItems.count
If Trim(lvB.ListItems(i).Text) = "" Then lvB.ListItems.Remove i
Next i
SaveValue "count", lvB.ListItems.count, "Browsers", FullPath(App.Path, "editors.inf")
For i = 1 To lvB.ListItems.count
SaveValue "item" & i, lvB.ListItems(i).Text, "Browsers", FullPath(App.Path, "editors.inf")
Next i
Unload Me
frmMain.LoadBrowserList
End Sub

Private Sub cmEdit_Click()
On Error Resume Next
lvB.SetFocus
lvB.StartLabelEdit
End Sub

Private Sub cmRem_Click()
On Error Resume Next
lvB.ListItems.Remove lvB.SelectedItem.Index
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
Load_Browser_List
PaintIcon lvB.SelectedItem.Text, pp
End Sub

Sub Load_Browser_List()
On Error Resume Next
Dim i As Integer
Dim s As Integer
s = ReadValue("count", , "Browsers", FullPath(App.Path, "editors.inf"))
For i = 1 To s
lvB.ListItems.Add , "tmp", ReadValue("item" & i, , "Browsers", FullPath(App.Path, "editors.inf")), , 1
If lvB.ListItems("tmp").Text = "" Then lvB.ListItems.Remove "tmp"
lvB.ListItems("tmp").Key = ""
Next i
End Sub

Private Sub lvB_ItemClick(ByVal Item As MSComctlLib.ListItem)
pp.Cls
PaintIcon Item.Text, pp
End Sub
