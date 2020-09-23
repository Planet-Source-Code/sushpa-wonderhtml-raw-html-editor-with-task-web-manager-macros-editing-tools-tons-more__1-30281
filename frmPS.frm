VERSION 5.00
Begin VB.Form frmPS 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paste Special"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmPS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3225
      TabIndex        =   6
      Top             =   675
      Width           =   915
   End
   Begin VB.CommandButton cmHelp 
      Caption         =   "&Help"
      Height          =   390
      Left            =   3225
      TabIndex        =   7
      Top             =   1125
      Width           =   915
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Paste"
      Default         =   -1  'True
      Height          =   390
      Left            =   3225
      TabIndex        =   5
      Top             =   225
      Width           =   915
   End
   Begin VB.OptionButton opType 
      Caption         =   "Convert symbols to entities"
      Height          =   285
      Index           =   4
      Left            =   135
      TabIndex        =   4
      Top             =   1320
      Width           =   2940
   End
   Begin VB.OptionButton opType 
      Caption         =   "Treat breaks as paragraphs"
      Height          =   285
      Index           =   3
      Left            =   135
      TabIndex        =   3
      Top             =   1035
      Value           =   -1  'True
      Width           =   2940
   End
   Begin VB.OptionButton opType 
      Caption         =   "Treat as preformatted"
      Height          =   285
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   750
      Width           =   2940
   End
   Begin VB.OptionButton opType 
      Caption         =   "Convert to numbered list"
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   465
      Width           =   2940
   End
   Begin VB.OptionButton opType 
      Caption         =   "Convert to bulleted list"
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   2940
   End
   Begin VB.TextBox txPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1935
      Width           =   4110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clipboard contents:"
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   1710
      Width           =   1410
   End
End
Attribute VB_Name = "frmPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheChoice As Integer

Private Sub cmHelp_Click()
ShowHelp "PasteSpecial.rtf"
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
On Error Resume Next
Me.Hide
With frmMain.ActiveForm
Select Case TheChoice
Case 0
.mnuPasteSpecialUnordered_Click
Case 1
.mnuPasteSpecialOrdered_Click
Case 2
.mnuPasteSpecialPRE_Click
Case 3
.mnupasteSpecialBRasP_Click
Case 4
.mnuPasteSpecialTagsEntities_Click
End Select
End With
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
If Clipboard.GetFormat(vbCFText) = False Then
MsgBox "The clipboard does not contain text.", vbExclamation, "Error"
opType(0).Enabled = False
opType(1).Enabled = False
opType(2).Enabled = False
opType(3).Enabled = False
opType(4).Enabled = False
cmOk.Enabled = False
Else
opType_Click 3
txPrev.Text = Clipboard.GetText
Label1.Caption = Label1.Caption & " (" & Len(txPrev.Text) & " bytes)"
End If
TheChoice = 3
End Sub

Private Sub opType_Click(Index As Integer)
TheChoice = Index
End Sub
