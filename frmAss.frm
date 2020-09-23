VERSION 5.00
Begin VB.Form frmAss 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Associate Editor"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frmAss.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txDesc 
      Height          =   315
      Left            =   1215
      TabIndex        =   7
      Top             =   315
      Width           =   3480
   End
   Begin VB.CommandButton cmOK 
      Cancel          =   -1  'True
      Caption         =   "O&K"
      Default         =   -1  'True
      Height          =   375
      Left            =   2655
      TabIndex        =   6
      Top             =   1440
      Width           =   1005
   End
   Begin VB.CommandButton cmNo 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3735
      TabIndex        =   5
      Top             =   1440
      Width           =   1005
   End
   Begin VB.CommandButton cmPick 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   3710
      TabIndex        =   4
      Top             =   870
      Width           =   1000
   End
   Begin VB.TextBox txApp 
      Height          =   315
      Left            =   135
      TabIndex        =   3
      Top             =   900
      Width           =   3480
   End
   Begin VB.TextBox txExt 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   315
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   1215
      TabIndex        =   8
      Top             =   90
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Open with:"
      Height          =   240
      Left            =   135
      TabIndex        =   2
      Top             =   675
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Extension:"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   765
   End
End
Attribute VB_Name = "frmAss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmNo_Click()
assoc_text = ""
Unload Me
End Sub

Private Sub cmOK_Click()
On Error Resume Next
If txExt.Text = "" Then txExt.SetFocus: Exit Sub
If txApp.Text = "" Then txApp.SetFocus: Exit Sub
assoc_text = txExt.Text & ";" & txApp.Text & ";" & txDesc.Text
Unload Me
End Sub

Private Sub cmPick_Click()
On Error GoTo hell
Dim s As String
s = frmMain.CD.Filter
frmMain.CD.Filter = "Executable applications (*.exe, *.dll)|*.exe;*.dll"
frmMain.CD.Filename = txApp.Text
frmMain.CD.ShowOpen
txApp.Text = frmMain.CD.Filename
txApp.SetFocus
txExt.SetFocus
hell:
End Sub

Private Sub Form_Load()
SetFont Me
End Sub

Private Sub txApp_GotFocus()
frmMain.SB.Panels(1).Text = "Enter the program you want to open " & Chr(147) & FileType(txExt.Text) & Chr(148) & " type of files with."
End Sub

Private Sub txApp_LostFocus()
On Error Resume Next
If Dir(txApp.Text) = "" Then
MsgBox "The file cannot be found.", vbExclamation
txApp.SelStart = 0
txApp.SelLength = Len(txApp.Text)
txApp.SetFocus
End If
End Sub

Private Sub txExt_GotFocus()
frmMain.SB.Panels(1).Text = "Enter the extension of the file type you want to set up."
End Sub

Private Sub txExt_LostFocus()
On Error Resume Next
If Left(txExt.Text, 1) = "." Then txExt.Text = Right(txExt.Text, Len(txExt.Text) - 1)
If txDesc.Text = "" Then txDesc.Text = UCase(txExt.Text) & " File Handler"
End Sub
