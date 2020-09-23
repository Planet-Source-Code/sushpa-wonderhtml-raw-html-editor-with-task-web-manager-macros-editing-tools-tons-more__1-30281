VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmButt 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Buttons"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "frmButt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmOK 
      Caption         =   "Cr&eate"
      Height          =   375
      Left            =   1350
      TabIndex        =   9
      Top             =   2745
      Width           =   1005
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2430
      TabIndex        =   7
      Top             =   2745
      Width           =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2580
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   4551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Create"
      TabPicture(0)   =   "frmButt.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imIMages"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txCur"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txTTI"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "pp"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Remove"
      TabPicture(1)   =   "frmButt.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvBtn"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox pp 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   2340
         Picture         =   "frmButt.frx":0044
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Top             =   900
         Width           =   480
      End
      Begin MSComctlLib.ListView lvBtn 
         Height          =   2085
         Left            =   -74910
         TabIndex        =   8
         Top             =   405
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   3678
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imlTB2"
         SmallIcons      =   "imlTB2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Button"
            Object.Width           =   2593
         EndProperty
      End
      Begin VB.TextBox txTTI 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   2040
         Width           =   2985
      End
      Begin VB.TextBox txCur 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   1410
         Width           =   1545
      End
      Begin MSComctlLib.ImageCombo imIMages 
         Height          =   330
         Left            =   180
         TabIndex        =   3
         Top             =   780
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "00"
         ImageList       =   "imlTB2"
      End
      Begin MSComctlLib.ImageList imlTB2 
         Left            =   45
         Top             =   -135
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
               Picture         =   "frmButt.frx":090E
               Key             =   "custom"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Image:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   555
         Width           =   1590
      End
      Begin VB.Label Label3 
         Caption         =   "Text to Insert:"
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   1815
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Place cursor at:"
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   1185
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1350
      TabIndex        =   10
      Top             =   2745
      Width           =   1005
   End
End
Attribute VB_Name = "frmButt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmDel_Click()
On Error Resume Next
If lvBtn.SelectedItem Is Nothing Then Exit Sub
SaveValue "Btn" & (19 + lvBtn.SelectedItem.Index), "", "Buttons"
SaveValue "Btn" & (19 + lvBtn.SelectedItem.Index) & "Sel", 0, "Buttons"
SaveValue "Btn" & (19 + lvBtn.SelectedItem.Index) & "Img", 0, "Buttons"
Me.Hide
frmMain.LoadToolBar
Unload Me
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
SaveValue "BtnCount", frmMain.TB2.Buttons.count + 1, "Buttons"
SaveValue "Btn" & frmMain.TB2.Buttons.count + 1, txTTI.Text, "Buttons"
SaveValue "Btn" & frmMain.TB2.Buttons.count + 1 & "Img", imIMages.Text, "Buttons"
SaveValue "Btn" & frmMain.TB2.Buttons.count + 1 & "Sel", txCur.Text, "Buttons"
Me.Hide
frmMain.LoadToolBar
Unload Me
End Sub

Private Sub Form_Load()
SetFont Me
Dim i As Integer, s As String
For i = 1 To imlTB2.ListImages.count
imIMages.ComboItems.Add , imlTB2.ListImages(i).Key, imlTB2.ListImages(i).Key, i
'String(2 - Len(CStr(i)), "0") & i
Next i
If frmMain.TB2.Buttons.count <= 19 Then Exit Sub
For i = 20 To frmMain.TB2.Buttons.count
s = ReadValue("Btn" & i, "", "Buttons")
If s = "" Then GoTo n
lvBtn.ListItems.Add , , s, , ReadValue("Btn" & i & "Img", 0, "Buttons")
n: 'next item
Next i
End Sub

Private Sub imIMages_GotFocus()
frmMain.SB.Panels(1).Text = "Select the icon to display on the button."
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 1 Then cmOK.SetFocus Else cmDel.SetFocus
End Sub

Private Sub txCur_LostFocus()
If Not IsNumeric(txCur.Text) Then txCur.Text = 0
If Val(txCur.Text) > Len(txTTI.Text) + 1 Then txCur.Text = 0
End Sub

Private Sub txTTI_GotFocus()
frmMain.SB.Panels(1).Text = "Text to insert when button is pressed. Use | for indicating a new line."
End Sub


Private Sub txCur_GotFocus()
frmMain.SB.Panels(1).Text = "The position to move the cursor to (counted from start)"
End Sub

