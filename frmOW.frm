VERSION 5.00
Begin VB.Form frmOW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open with..."
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "frmOW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox pp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3150
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   480
   End
   Begin VB.ListBox cbApps 
      Height          =   1230
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   2670
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2970
      TabIndex        =   3
      Top             =   495
      Width           =   870
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Open"
      Default         =   -1  'True
      Height          =   375
      Left            =   2970
      TabIndex        =   2
      Top             =   90
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "Program to use:"
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   1410
   End
End
Attribute VB_Name = "frmOW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colApps As New Collection

Private Sub cbApps_Click()
pp.Cls
PaintIcon colApps(cbApps.ListIndex + 1), pp
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
On Error Resume Next
Dim sParam As String * 260
GetShortPathName Caption, sParam, 260
ShellExecute hwnd, "open", colApps(cbApps.ListIndex + 1), sParam, "", 10
Unload Me
End Sub

Private Sub Form_Load()
SetFont Me
GetAssoc
End Sub

Sub GetAssoc()
On Error Resume Next
Dim i As Integer, assoc As String
Dim ss() As String, vals() As String

assoc = ReadValue("Config", "", "Editors", FullPath(App.Path, "editors.inf"))
ss = Split(assoc, ",")

For i = 0 To UBound(ss)
assoc = ReadValue(ss(i), "", "Editors", FullPath(App.Path, "editors.inf"))
If assoc = "" Then GoTo n
vals = Split(assoc, "|")
cbApps.AddItem GetFile(vals(1)) & " (" & vals(0) & ")"
colApps.Add vals(1)
n:
Next i

End Sub
