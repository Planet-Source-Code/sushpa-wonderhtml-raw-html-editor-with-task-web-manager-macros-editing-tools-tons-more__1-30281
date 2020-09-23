VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStats 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Statistics"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1125
      TabIndex        =   0
      Top             =   1440
      Width           =   960
   End
   Begin RichTextLib.RichTextBox RTF1 
      Height          =   1950
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   3440
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmStats.frx":058A
   End
   Begin VB.Label lbChars 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2475
      TabIndex        =   10
      Top             =   1035
      Width           =   45
   End
   Begin VB.Label lbParagraphs 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2475
      TabIndex        =   9
      Top             =   810
      Width           =   45
   End
   Begin VB.Label lbLines 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2475
      TabIndex        =   8
      Top             =   585
      Width           =   45
   End
   Begin VB.Label lbSentences 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2475
      TabIndex        =   7
      Top             =   360
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Characters:"
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   6
      Top             =   1035
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Paragraphs:"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   5
      Top             =   810
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lines:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   4
      Top             =   585
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sentences:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   360
      Width           =   810
   End
   Begin VB.Label lbWords 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2475
      TabIndex        =   2
      Top             =   135
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Words:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   525
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ShowStats(T As String)
Screen.MousePointer = 11: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(2).Picture
RTF1.Text = T
CleanUp RTF1, False
ReplaceStuff RTF1
ConvEntities RTF1
lbWords.Caption = StrCount(RTF1.Text, " ") + StrCount(RTF1.Text, "/") + StrCount(RTF1.Text, "-") + 1
lbSentences.Caption = StrCount(RTF1.Text, ". ") + StrCount(RTF1.Text, "." & vbNewLine)
lbChars.Caption = Len(RTF1.Text)
lbLines.Caption = GetTotalLines(RTF1)
lbParagraphs.Caption = StrCount(RTF1.Text, vbNewLine & vbNewLine)
Screen.MousePointer = 0: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
End Sub

Private Sub cmOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetFont Me
End Sub

Private Sub RTF1_SelChange()
frmMain.SB.Panels(1).Text = "Calculating... " & Round((RTF1.SelStart * 100) / Len(RTF1.Text), 0) & "% done"
End Sub
