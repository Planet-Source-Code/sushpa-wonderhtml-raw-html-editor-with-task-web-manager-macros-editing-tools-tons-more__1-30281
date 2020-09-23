VERSION 5.00
Begin VB.Form frmReg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FREEWARE Registration"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cbHome 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmReg.frx":058A
      Left            =   1045
      List            =   "frmReg.frx":0651
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "United States of America"
      Top             =   1215
      Width           =   3570
   End
   Begin VB.TextBox txHomepage 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1045
      TabIndex        =   5
      Text            =   "http://"
      Top             =   1935
      Width           =   3570
   End
   Begin VB.TextBox txEmail 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1045
      TabIndex        =   4
      Top             =   1575
      Width           =   3570
   End
   Begin VB.TextBox txAge 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1045
      TabIndex        =   2
      Top             =   855
      Width           =   3570
   End
   Begin VB.TextBox txOccup 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1045
      TabIndex        =   1
      Top             =   495
      Width           =   3570
   End
   Begin VB.TextBox txName 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1045
      TabIndex        =   0
      Top             =   135
      Width           =   3570
   End
   Begin VB.CommandButton cmNo 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   3195
      Width           =   1005
   End
   Begin VB.CommandButton cmOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Register!"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3195
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Residence:"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   1275
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmReg.frx":08F6
      Height          =   1005
      Left            =   135
      TabIndex        =   13
      Top             =   2340
      Width           =   4425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Homepage:"
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   12
      Top             =   1995
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail ID:"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   1635
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age / Sex:"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   915
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   555
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name:"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   195
      Width           =   840
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetFont Me
Label3.Caption = "----PRIVACY POLICY----" & vbCrLf & Label3.Caption
End Sub

