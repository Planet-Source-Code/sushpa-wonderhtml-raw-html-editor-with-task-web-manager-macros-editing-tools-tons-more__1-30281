VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFile 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Information"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy &to clipboard"
      Height          =   375
      Left            =   60
      TabIndex        =   22
      Top             =   4770
      Width           =   1680
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   4770
      UseMaskColor    =   -1  'True
      Width           =   870
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3465
      TabIndex        =   7
      Top             =   4770
      UseMaskColor    =   -1  'True
      Width           =   870
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4650
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   8202
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Properties"
      TabPicture(0)   =   "frmFileInfo.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Image2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "frFrame"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txLoc"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txCre"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txTitle"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txSize"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txType"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txMod"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txAcc"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txMIME"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txShort"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txName"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "iml"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Tasks"
      TabPicture(1)   =   "frmFileInfo.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txComments"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chTask"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txTask"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frFile"
      Tab(1).Control(4)=   "cmGetTaskInfo"
      Tab(1).Control(5)=   "cmSaveTask"
      Tab(1).Control(6)=   "cmDelTask"
      Tab(1).Control(7)=   "Line1(1)"
      Tab(1).Control(8)=   "lbFile"
      Tab(1).Control(9)=   "Label7"
      Tab(1).Control(10)=   "lbTask"
      Tab(1).Control(11)=   "Image1"
      Tab(1).Control(12)=   "lbTaskStatus"
      Tab(1).ControlCount=   13
      Begin MSComctlLib.ImageList iml 
         Left            =   810
         Top             =   1710
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFileInfo.frx":0044
               Key             =   "unknown"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFileInfo.frx":03F8
               Key             =   "folder"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFileInfo.frx":04F8
               Key             =   "html"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFileInfo.frx":0628
               Key             =   "image"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFileInfo.frx":077C
               Key             =   "text"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   630
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "File not found."
         Top             =   525
         Width           =   3525
      End
      Begin VB.TextBox txComments 
         Height          =   315
         Left            =   -74820
         MaxLength       =   255
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1005
         Width           =   3900
      End
      Begin VB.CheckBox chTask 
         Caption         =   "Mark this file with a task."
         Height          =   195
         Left            =   -74820
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2850
         Width           =   3500
      End
      Begin VB.TextBox txTask 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74820
         MaxLength       =   255
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3390
         Width           =   3900
      End
      Begin VB.Frame frFile 
         Caption         =   "Untitled"
         Height          =   125
         Left            =   -74820
         TabIndex        =   26
         Top             =   405
         Width           =   3900
      End
      Begin VB.CommandButton cmGetTaskInfo 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   -74805
         TabIndex        =   25
         Top             =   3930
         Width           =   1095
      End
      Begin VB.CommandButton cmSaveTask 
         Caption         =   "&Update"
         Height          =   375
         Left            =   -73650
         TabIndex        =   24
         Top             =   3930
         Width           =   1005
      End
      Begin VB.CommandButton cmDelTask 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   -71790
         TabIndex        =   23
         Top             =   3930
         Width           =   870
      End
      Begin VB.TextBox txShort 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1260
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Unable to determine DOS Name."
         Top             =   2130
         Width           =   2800
      End
      Begin VB.TextBox txMIME 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1260
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "0.0 Seconds @ 0 KB/s (0 K)"
         Top             =   3225
         Width           =   2800
      End
      Begin VB.TextBox txAcc 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1260
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Mon, 1 Jan 1901"
         Top             =   1410
         Width           =   2800
      End
      Begin VB.TextBox txMod 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1260
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Mon, 1 Jan 1901, 12:00 AM"
         Top             =   1185
         Width           =   2800
      End
      Begin VB.TextBox txType 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1260
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Unknown file type"
         Top             =   2760
         Width           =   2800
      End
      Begin VB.TextBox txSize 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1260
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0 KB (0 Bytes)"
         Top             =   2535
         Width           =   2800
      End
      Begin VB.TextBox txTitle 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1260
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   2985
         Width           =   2800
      End
      Begin VB.TextBox txCre 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1260
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Mon, 1 Jan 1901, 12:00 AM"
         Top             =   960
         Width           =   2800
      End
      Begin VB.TextBox txLoc 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1260
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Unknown, shared or non-existent"
         Top             =   1905
         Width           =   2800
      End
      Begin VB.Frame frFrame 
         Caption         =   "File attributes"
         Height          =   870
         Left            =   180
         TabIndex        =   34
         Top             =   3600
         Width           =   3885
         Begin VB.CheckBox chFol 
            Caption         =   "Directory"
            Height          =   195
            Left            =   2295
            TabIndex        =   35
            Top             =   495
            Width           =   1050
         End
         Begin VB.CheckBox chHD 
            Caption         =   "Hidden"
            Height          =   195
            Left            =   1170
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   495
            Width           =   915
         End
         Begin VB.CheckBox chRO 
            Caption         =   "Read-Only"
            Height          =   195
            Left            =   1170
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   270
            Width           =   1590
         End
         Begin VB.CheckBox chAr 
            Caption         =   "Archive"
            Height          =   195
            Left            =   180
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   270
            Width           =   1320
         End
         Begin VB.CheckBox chSys 
            Caption         =   "System"
            Height          =   195
            Left            =   180
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   495
            Width           =   1320
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   -74640
         X2              =   -71085
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Image Image2 
         Height          =   330
         Left            =   225
         Picture         =   "frmFileInfo.frx":0880
         Top             =   450
         Width           =   300
      End
      Begin VB.Label lbFile 
         AutoSize        =   -1  'True
         Caption         =   "Untitled"
         Height          =   195
         Left            =   -74685
         TabIndex        =   33
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Comments:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   32
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lbTask 
         AutoSize        =   -1  'True
         Caption         =   "Task Description:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -74820
         TabIndex        =   31
         Top             =   3165
         Width           =   1230
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74775
         Picture         =   "frmFileInfo.frx":0971
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label lbTaskStatus 
         Caption         =   "There are no tasks related to this file."
         Height          =   645
         Left            =   -74235
         TabIndex        =   30
         Top             =   2160
         Width           =   3000
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "DOS Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   2130
         Width           =   825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Load time:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   3225
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Accessed:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Modified:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1185
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   2760
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File Size:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   2535
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Page Title:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   2985
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Created:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1905
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Guess what - this form showed Filename, Filesize and Date
'only, before I recklessly tweaked it. Lot of progress, nah?
Option Explicit

Private Sub chAr_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
chAr.Value = 1 - chAr.Value
End Sub

Private Sub chFol_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
chFol.Value = 1 - chFol.Value
End Sub

Private Sub chHD_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
chHD.Value = 1 - chHD.Value
End Sub

Private Sub chRO_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
chRO.Value = 1 - chRO.Value
End Sub

Private Sub chSys_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
chSys.Value = 1 - chSys.Value
End Sub

Private Sub chTask_Click()
On Error Resume Next
txTask.Enabled = CBool(chTask.Value)
lbTask.Enabled = txTask.Enabled
txTask.SetFocus
End Sub

Private Sub cmCopy_Click()
Dim s As String
s = s & String(15 + Len(txName.Text), "-") & vbCrLf
s = s & "Properties for " & txName.Text & vbCrLf
s = s & String(15 + Len(txName.Text), "-") & vbCrLf & vbCrLf
s = s & "Created:    " & txCre.Text & vbCrLf
s = s & "Modified:   " & txMod.Text & vbCrLf
s = s & "Accessed:   " & txAcc.Text & vbCrLf & vbCrLf
s = s & "Location:   " & txLoc.Text & vbCrLf
s = s & "DOS name:   " & txShort.Text & vbCrLf & vbCrLf
s = s & "File size:  " & txSize.Text & vbCrLf
s = s & "File type:  " & txType.Text & vbCrLf
s = s & "Page Title: " & txTitle.Text & vbCrLf
s = s & "Load time:  " & txMIME.Text & vbCrLf & vbCrLf
s = s & "Attributes: " & vbCrLf
s = s & " Archive:   " & IIf(chAr.Value = 1, "Yes", "No") & vbCrLf
s = s & " Readonly:  " & IIf(chRO.Value = 1, "Yes", "No") & vbCrLf
s = s & " Hidden:    " & IIf(chHD.Value = 1, "Yes", "No") & vbCrLf
s = s & " System:    " & IIf(chSys.Value = 1, "Yes", "No") & vbCrLf
Clipboard.Clear
Clipboard.SetText s
End Sub

Private Sub cmDelTask_Click()
chTask.Value = 0
txTask.Text = ""
SaveComment txName.Text
End Sub

Private Sub cmGetTaskInfo_Click()
GetComments txName.Text
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
SaveComment txName.Text
Unload Me
End Sub

Private Sub cmSaveTask_Click()
SaveComment txName.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
Screen.MousePointer = vbHourglass
End Sub

Sub Proceed()
On Error Resume Next
Dim tmp As String * 255, lenShort As Long
'we needed to wait a while for the form to load, then
'do the stuff with the FileDateTime and so on.
txName.Text = InitCap(Tag)
If Dir(txName.Text, GetAttr(txName.Text)) = "" Then
Dim lpF As Control
For Each lpF In Me.Controls
lpF.Enabled = False
Next lpF
cmNo.Enabled = True
SSTab1.Enabled = True
cmDelTask.Enabled = True
txType.Text = "No information available."
cmNo.Caption = "Cl&ose"
Screen.MousePointer = 0: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
Exit Sub
End If
txTitle.Text = GetTitle(txName.Text)
PaintGeneralizedIcon
Caption = "File Information: " & GetFile(Tag)
txSize.Text = Round(FileLen(Tag) / 1024, 2) & " KB (" & Format(FileLen(Tag), "00,00") & " Bytes)"
If GetAttr(txName.Text) And vbDirectory Then txSize.Text = "-": chFol.Value = 1
GetFileInfo txName.Text, txMod, txCre, txAcc
txLoc.Text = GetPath(txName.Text)
GetFileType
GetComments txName.Text
GetAttributes
txMIME.Text = GetSpeedInfo(FileLen(txName.Text))
frFile.Caption = txName.Text
lbFile.Caption = txName.Text
lenShort = GetShortPathName(txName.Text, tmp, 255)
txShort.Text = Left$(tmp, lenShort)
Screen.MousePointer = 0: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
End Sub

Sub GetFileType()
Dim s As String
If txName.Text = Ext(txName.Text) Then txType.Text = "Directory (Folder)":  Exit Sub
s = GetValue(HKEY_CLASSES_ROOT, "." & Ext(txName.Text), "", "")
If s = "" Then txType.Text = "Unknown file: *." & (Ext(txName.Text)): Exit Sub
s = GetValue(HKEY_CLASSES_ROOT, s, "", "")
If s = "" Then txType.Text = "Unknown file: *." & (Ext(txName.Text)): Exit Sub
txType.Text = s
End Sub

Sub GetAttributes()
Dim l As VbFileAttribute
l = GetAttr(txName.Text)
If (l And vbArchive) Then chAr.Value = 1
If (l And vbReadOnly) Then chRO.Value = 1
If (l And vbSystem) Then chSys.Value = 1
If (l And vbHidden) Then chHD.Value = 1
End Sub

Sub GetComments(Filename As String)
Dim s As String, sTime As String
s = ReadValue("Task", "", Filename, FullPath(frmMain.CurrentWeb, "files.inf"))
txComments.Text = ReadValue("Comment", "", Filename, FullPath(frmMain.CurrentWeb, "files.inf"))
chTask.Value = CBinary(s <> "")
txTask.Text = s
sTime = ReadValue("TaskTime", "(unknown)", Filename, FullPath(frmMain.CurrentWeb, "files.inf"))
If chTask.Value = 1 Then lbTaskStatus.Caption = "This file is marked with a pending task." & vbCrLf & "Date: " & sTime
End Sub

Sub SaveComment(Filename As String)
Dim s As String, sTime As String
Dim File As String

File = FullPath(frmMain.CurrentWeb, "files.inf")

s = ReadValue("Comment", "", Filename, File)
If txComments.Text <> s Then SaveValue "Comment", txComments.Text, Filename, File

s = ReadValue("Task", "", Filename, File)
If s <> txTask.Text Then SaveValue "Task", txTask.Text, Filename, File Else Exit Sub

sTime = ReadValue("TaskTime", "(unknown)", Filename, FullPath(frmMain.CurrentWeb, "files.inf"))
SaveValue "TaskTime", Format(Date, "dd mmm yyyy, at ") & Format(Time, "h:mm AMPM."), Filename, File
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 1 Then txComments.SetFocus Else cmOK.SetFocus
End Sub

Private Sub txTask_GotFocus()
txTask.SelStart = 0
txTask.SelLength = Len(txTask.Text)
End Sub

Private Sub txTask_LostFocus()
txTask.Text = Replace(txTask.Text, "[", "(")
txTask.Text = Replace(txTask.Text, "]", ")")
End Sub

Sub PaintGeneralizedIcon()
Dim s As String
Select Case Ext(txName.Text)
Case "htm", "html", "asp", "shtml", "php"
s = "html"
Case "txt", "cgi", "js", "vbs"
s = "text"
Case "jpg", "gif", "bmp", "jpe", "png"
s = "image"
Case Else
s = "unknown"
End Select
If GetAttr(txName.Text) And vbDirectory Then s = "folder"
Image2.Picture = iml.ListImages(s).Picture
End Sub
