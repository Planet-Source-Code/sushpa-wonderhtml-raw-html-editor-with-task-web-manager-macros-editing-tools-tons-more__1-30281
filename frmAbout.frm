VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About WonderHTML 0.90"
   ClientHeight    =   4350
   ClientLeft      =   2325
   ClientTop       =   1995
   ClientWidth     =   5550
   Icon            =   "frmAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox pBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4350
      Left            =   0
      ScaleHeight     =   4290
      ScaleWidth      =   5490
      TabIndex        =   0
      Top             =   0
      Width           =   5550
      Begin VB.CommandButton cmClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   375
         Left            =   4275
         TabIndex        =   1
         Top             =   3690
         Width           =   1005
      End
      Begin VB.PictureBox pIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   810
         Left            =   240
         Picture         =   "frmAbout.frx":1042
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   2
         Top             =   855
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wonder"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "HTML"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   405
         Width           =   930
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000016&
         BorderWidth     =   2
         X1              =   0
         X2              =   5535
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000016&
         BorderWidth     =   2
         X1              =   615
         X2              =   615
         Y1              =   4365
         Y2              =   900
      End
      Begin VB.Label lbCredits 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":1A19
         Height          =   1455
         Left            =   720
         TabIndex        =   6
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   4965
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   630
         X2              =   630
         Y1              =   900
         Y2              =   4365
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         BorderWidth     =   2
         X1              =   0
         X2              =   5535
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Label lbLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://sushantshome.tripod.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   720
         MouseIcon       =   "frmAbout.frx":1B4C
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   3690
         Width           =   2805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "© 2001, Sushant Pandurangi."
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   3915
         Width           =   2160
      End
      Begin VB.Label lbInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":2416
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   1440
         TabIndex        =   3
         Top             =   225
         Width           =   4020
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
lbInfo.Caption = Replace(lbInfo.Caption, "\n", vbCrLf)
lbCredits.Caption = Replace(lbCredits.Caption, "\n", vbCrLf)
Label2(0).FontName = "Times New Roman"
Label2(1).FontName = "Times New Roman"
End Sub

Private Sub lbLink_Click()
ShellExecute hwnd, "open", lbLink.Caption, "", "", 10
End Sub

