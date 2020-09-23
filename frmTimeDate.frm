VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Date and Time"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   Icon            =   "frmTimeDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txPrev 
      Height          =   285
      Left            =   2070
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   855
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeDate.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeDate.frx":0CEC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Insert"
      Default         =   -1  'True
      Height          =   375
      Left            =   2475
      TabIndex        =   3
      Top             =   1260
      Width           =   915
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3420
      TabIndex        =   4
      Top             =   1260
      Width           =   915
   End
   Begin VB.ComboBox lstTimes 
      Height          =   315
      Left            =   727
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   855
      Width           =   3480
   End
   Begin VB.ComboBox lstDates 
      Height          =   315
      Left            =   727
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   3480
   End
   Begin VB.PictureBox PP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   187
      Picture         =   "frmTimeDate.frx":15C8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   405
      Width           =   480
   End
   Begin VB.CheckBox chUpd 
      Caption         =   "U&pdate automatically"
      Height          =   240
      Left            =   135
      TabIndex        =   2
      Top             =   1305
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Time Format:"
      Height          =   195
      Left            =   727
      TabIndex        =   7
      Top             =   630
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Date Format:"
      Height          =   195
      Left            =   727
      TabIndex        =   6
      Top             =   45
      Width           =   1050
   End
End
Attribute VB_Name = "frmDate"
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
Dim s As String
s = IIf(chUpd.Value = 1, GetUpdateCode(), txPrev.Text)
With frmMain.ActiveForm.RTF1
.SelText = s
End With
Unload Me
End Sub

Private Sub Form_Load()
SetFont Me
lstDates.AddItem ""
lstDates.AddItem Format(Date, "mmmm dd, yyyy")
lstDates.AddItem Format(Date, "dddd, mmmm dd, yyyy")
lstDates.AddItem Format(Date, "dd mmm, yyyy")
lstDates.AddItem Format(Date, "dddd, dd mmm yyyy")
lstDates.AddItem Format(Date, "dd-mm-yyyy")
lstDates.AddItem Format(Date, "ddd dd-mm-yyyy")
lstDates.AddItem Format(Date, "yyyy-mm-dd")
lstDates.AddItem Format(Date, "mmmm yyyy")
lstDates.ListIndex = 0

lstTimes.AddItem ""
lstTimes.AddItem Format(Time, "hh:nn AMPM")
lstTimes.AddItem Format(Time, "hh:nn:ss AMPM")
lstTimes.AddItem Format(Time, "hh:nn")
lstTimes.AddItem Format(Time, "hh:nn:ss")
lstTimes.ListIndex = 0
End Sub

Private Sub lstDates_Click()
Preview
End Sub

Private Sub lstDates_GotFocus()
PP.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub lstTimes_Click()
Preview
End Sub

Sub Preview()
txPrev.Text = lstDates.Text & " " & lstTimes.Text
If txPrev.Text = " " Then txPrev.Text = ""
txPrev.Text = Trim(txPrev.Text)
End Sub

Function GetUpdateCode() As String
Dim s As String
s = txPrev.Text
s = "<!-- WHTMLDATE " & FormatsD(lstDates.ListIndex) & " " & FormatsT(lstTimes.ListIndex) & "//-->" & txPrev.Text & "<!-- ENDWHTMLDATE //-->"
GetUpdateCode = s
End Function

Function FormatsD(nIndex As Long) As String
Select Case nIndex
Case 0
FormatsD = ""
Case 1
FormatsD = "mmmm dd, yyyy"
Case 2
FormatsD = "dddd, mmmm dd, yyyy"
Case 3
FormatsD = "dd mmm, yyyy"
Case 4
FormatsD = "dddd dd mmm, yyyy"
Case 5
FormatsD = "dd-mm-yyyy"
Case 6
FormatsD = "ddd dd-mm-yy"
Case 7
FormatsD = "yyy, mm dd"
Case 8
FormatsD = "mmmm yyyy"
End Select
End Function

Function FormatsT(nIndex As Long) As String
Select Case nIndex
Case 0
FormatsT = ""
Case 1
FormatsT = "hh:nn AMPM"
Case 2
FormatsT = "hh:nn:ss AMPM"
Case 3
FormatsT = "hh:nn"
Case 4
FormatsT = "hh:nn:ss"
End Select
End Function

Private Sub lstTimes_GotFocus()
PP.Picture = ImageList1.ListImages(1).Picture
End Sub
