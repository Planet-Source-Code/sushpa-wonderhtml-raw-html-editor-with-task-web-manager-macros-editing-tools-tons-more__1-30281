VERSION 5.00
Begin VB.Form frmDel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete file"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmDel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3870
      TabIndex        =   1
      Top             =   1170
      Width           =   960
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Delete"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   1170
      Width           =   1095
   End
   Begin VB.PictureBox PD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   180
      Picture         =   "frmDel.frx":000C
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   810
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Do you want to send these file(s) to the recycle bin?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   990
      TabIndex        =   2
      Top             =   315
      Width           =   3840
   End
End
Attribute VB_Name = "frmDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Delete confirmation
Option Explicit

Private Sub cmNo_Click()
IfDeleteFile = False
Unload Me
End Sub

Private Sub cmOK_Click()
IfDeleteFile = True
Unload Me
End Sub

Private Sub Form_Load()
SetFont Me
End Sub
