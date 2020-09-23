VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTemp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Use Template"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   Icon            =   "frmTemp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   2550
      TabIndex        =   3
      Top             =   2730
      Width           =   1275
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Continue"
      Height          =   375
      Left            =   1215
      TabIndex        =   2
      Top             =   2730
      Width           =   1275
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   3923
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSComctlLib.ImageList imIcons 
      Left            =   3653
      Top             =   585
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
            Picture         =   "frmTemp.frx":1042
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvTemp 
      Height          =   2580
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   4551
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Template"
         Object.Width           =   6085
      EndProperty
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmNo_Click()
Unload Me

End Sub

Private Sub cmOK_Click()
On Error Resume Next
If lvTemp.SelectedItem Is Nothing Then lvTemp.SelectedItem = lvTemp.ListItems(1)
Dim lpF As New frmChild
Load lpF
lpF.RTF1.LoadFile lvTemp.SelectedItem.Key, rtfText
lpF.RTF1.SetFocus
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
Dim i As Integer
File1.Path = FullPath(App.Path, "Templates")
For i = 0 To File1.ListCount - 1
lvTemp.ListItems.Add , FullPath(File1.Path, File1.list(i)), NoExt(File1.list(i)), , 1
Next i
End Sub
