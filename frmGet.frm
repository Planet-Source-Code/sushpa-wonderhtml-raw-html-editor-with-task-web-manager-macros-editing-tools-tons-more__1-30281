VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load Remote File"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmGet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   3930
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8599
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "frmGet.frx":000C
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmStop 
      Caption         =   "Abo&rt"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1260
      TabIndex        =   8
      Top             =   3510
      Width           =   870
   End
   Begin VB.CommandButton cmStart 
      Caption         =   "&Begin"
      Default         =   -1  'True
      Height          =   375
      Left            =   3225
      TabIndex        =   2
      Top             =   3517
      Width           =   870
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   4170
      TabIndex        =   3
      Top             =   3517
      Width           =   870
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   165
      TabIndex        =   1
      Top             =   3517
      Width           =   1005
   End
   Begin VB.ComboBox cbAddr 
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   367
      Width           =   2595
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   270
      Left            =   2460
      TabIndex        =   5
      Top             =   3112
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin SHDocVwCtl.WebBrowser IEMain 
      Height          =   2175
      Left            =   2460
      TabIndex        =   4
      Top             =   817
      Width           =   2595
      ExtentX         =   4577
      ExtentY         =   3836
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3270
      Left            =   165
      Picture         =   "frmGet.frx":05A8
      ScaleHeight     =   3210
      ScaleWidth      =   2040
      TabIndex        =   6
      Top             =   142
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Document address:"
      Height          =   195
      Left            =   2460
      TabIndex        =   7
      Top             =   142
      Width           =   1395
   End
End
Attribute VB_Name = "frmGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmStart_Click()
On Error Resume Next
IEMain.Navigate cbAddr.Text
IEMain.SetFocus
If Not IsContained(cbAddr.Text, cbAddr) And cbAddr.Text <> "" Then cbAddr.AddItem cbAddr.Text, 0
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
IEMain.Navigate "about:blank"
GetAddrs
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetAddrs
End Sub

Private Sub IEMain_DownloadBegin()
cmStop.Enabled = True
End Sub

Private Sub IEMain_DownloadComplete()
On Error Resume Next
Dim sText As String
cmStop.Enabled = False
sText = IEMain.Document.documentelement.innerhtml
If InStr(1, sText, "Cannot find server", vbTextCompare) = 14 Or _
InStr(1, sText, "No page to display", vbTextCompare) = 14 Then _
SB.Panels(1).Text = "Error occured while loading page.": Exit Sub
End Sub

Private Sub IEMain_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
PB.Max = ProgressMax
PB.Value = Progress
End Sub

Private Sub IEMain_SetSecureLockIcon(ByVal SecureLockIcon As Long)
SB.Panels(2).Visible = CBool(SecureLockIcon)
End Sub

Private Sub IEMain_StatusTextChange(ByVal Text As String)
SB.Panels(1).Text = Text
End Sub

Private Sub IEMain_TitleChange(ByVal Text As String)
Caption = Text
End Sub

Private Sub IEMain_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
Cancel = True
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
On Error Resume Next
Dim lpF As New frmChild
Load lpF
lpF.Icon = lpF.p1.Picture
lpF.RTF1.Text = IEMain.Document.documentelement.innerhtml
lpF.SetFocus
lpF.RTF1.SetFocus
Unload Me
End Sub

Private Sub mnuHelpTopics_Click()
ShowHelp "GetRemote.rtf"
End Sub

Sub GetAddrs()
On Error Resume Next
Dim i As Integer, s As String
For i = 0 To ReadValue("URLCount", , "Get") - 1
s = ReadValue("URL" & i, , "Get")
If s <> "" Then cbAddr.AddItem s
Next i
End Sub

Sub SetAddrs()
On Error Resume Next
Dim i As Integer
SaveValue "URLCount", IIf(cbAddr.ListCount > 10, 10, cbAddr.ListCount), "Get"
For i = 0 To 9
SaveValue "URL" & i, cbAddr.list(i), "Get"
Next i
End Sub
