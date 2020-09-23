VERSION 5.00
Begin VB.Form frmCPick 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colour picker"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   HasDC           =   0   'False
   Icon            =   "frmCPick.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox pTmp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2160
      Picture         =   "frmCPick.frx":058A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   270
      Width           =   480
   End
   Begin VB.PictureBox pStan 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   75
      MousePointer    =   2  'Cross
      Picture         =   "frmCPick.frx":22B4
      ScaleHeight     =   420
      ScaleWidth      =   1680
      TabIndex        =   11
      Top             =   315
      Width           =   1740
      Begin VB.Shape shpTmp 
         BorderStyle     =   3  'Dot
         Height          =   210
         Left            =   0
         Top             =   0
         Width           =   210
      End
      Begin VB.Shape shpMov 
         BorderWidth     =   2
         Height          =   210
         Left            =   0
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.PictureBox pHolder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2865
      Left            =   75
      MousePointer    =   2  'Cross
      Picture         =   "frmCPick.frx":2956
      ScaleHeight     =   2805
      ScaleWidth      =   2625
      TabIndex        =   10
      Top             =   1125
      Width           =   2685
      Begin VB.Line lY 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3005
      End
      Begin VB.Line lX 
         X1              =   0
         X2              =   2700
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lY2 
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3005
      End
      Begin VB.Line lX2 
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         X1              =   0
         X2              =   2700
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox pPrev2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3015
      ScaleHeight     =   480
      ScaleWidth      =   840
      TabIndex        =   5
      Top             =   3330
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2910
      TabIndex        =   4
      Top             =   645
      Width           =   1005
   End
   Begin VB.PictureBox pPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3015
      ScaleHeight     =   480
      ScaleWidth      =   840
      TabIndex        =   3
      Top             =   2475
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Copy"
      Default         =   -1  'True
      Height          =   375
      Left            =   2910
      TabIndex        =   2
      Top             =   195
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Selected:"
      Height          =   195
      Index           =   1
      Left            =   3060
      TabIndex        =   9
      Top             =   2250
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current:"
      Height          =   195
      Index           =   0
      Left            =   3060
      TabIndex        =   8
      Top             =   3105
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Standard Colours:"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   90
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Custom Colour Palette:"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   930
      Width           =   1665
   End
   Begin VB.Label lbHex 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "#FFFFFF"
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
      Left            =   3015
      TabIndex        =   1
      Top             =   1905
      Width           =   780
   End
   Begin VB.Label lbRGB 
      Caption         =   "Red:              Green:         Blue:"
      Height          =   690
      Left            =   2880
      TabIndex        =   0
      Top             =   1170
      Width           =   1050
   End
End
Attribute VB_Name = "frmCPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lbHex.Caption
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
End Sub

Private Sub pHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim cl As Long
If Not lX2.Visible Then lX2.Visible = True
If Not lY2.Visible Then lY2.Visible = True
If shpTmp.Visible Then shpTmp.Visible = False
cl = GetPixel(pHolder.hdc, X \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
pPrev2.BackColor = cl
lY2.X1 = X
lY2.X2 = X
lX2.Y1 = y
lX2.Y2 = y
If Button = 1 Then
lbRGB.Caption = GetClrRGBVal(cl)
lbHex.Caption = GetHexVal(cl)
End If
End Sub

Private Sub pHolder_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim cl As Long
cl = GetPixel(pHolder.hdc, X \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
lbRGB.Caption = GetClrRGBVal(cl)
lbHex.Caption = GetHexVal(cl)
pPrev.BackColor = cl
lY.X1 = X
lY.X2 = X
lX.Y1 = y
lX.Y2 = y
End Sub

Private Sub pStan_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
shpTmp.Left = (X \ 210) * 210 'rounded
shpTmp.Top = (y \ 210) * 210 'rounded
pPrev2.BackColor = GetPixel(pStan.hdc, X / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
If lX2.Visible Then lX2.Visible = False
If lY2.Visible Then lY2.Visible = False
If Not shpTmp.Visible Then shpTmp.Visible = True
End Sub

Private Sub pStan_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
shpMov.Left = (X \ 210) * 210 'rounded
shpMov.Top = (y \ 210) * 210 'rounded
pStan.Refresh
Dim Color As Long
Color = GetPixel(pStan.hdc, X / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
lbHex.Caption = GetHexVal(Color)
lbRGB.Caption = GetClrRGBVal(Color)
pPrev.BackColor = Color
End Sub

