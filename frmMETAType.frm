VERSION 5.00
Begin VB.Form frmMType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Tag"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   Icon            =   "frmMETAType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txContent 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   1395
      Width           =   2715
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1845
      TabIndex        =   4
      Top             =   1800
      Width           =   1005
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   945
      TabIndex        =   3
      Top             =   1800
      Width           =   825
   End
   Begin VB.OptionButton opt 
      Caption         =   "http-equiv"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   900
      Width           =   1545
   End
   Begin VB.OptionButton opt 
      Caption         =   "name"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   675
      Value           =   -1  'True
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Content:"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   1170
      Width           =   645
   End
   Begin VB.Label lbID 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1418
      TabIndex        =   6
      Top             =   90
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "META type:"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   405
      Width           =   840
   End
End
Attribute VB_Name = "frmMType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmNo_Click()
returnedMTYPE = ""
Unload Me
End Sub

Private Sub cmOK_Click()
returnedMTYPE = IIf(opt(0).Value, opt(0).Caption, opt(1).Caption) & "ÿþýüûú" & txContent.Text 'ÿþýüûú junk is a delimiter
Unload Me
End Sub

Private Sub Form_Load()
SetFont Me
End Sub

Private Sub txContent_GotFocus()
txContent.SelStart = 0
txContent.SelLength = Len(txContent.Text)
End Sub
