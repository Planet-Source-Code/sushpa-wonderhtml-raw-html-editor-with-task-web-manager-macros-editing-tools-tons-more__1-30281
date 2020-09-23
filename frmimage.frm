VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImage 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmimage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox pHV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   900
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   7
      Top             =   2655
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox pV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   1170
      MousePointer    =   9  'Size W E
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   1530
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox pH 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   360
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   45
   End
   Begin MSComctlLib.ImageList imlTB 
      Left            =   540
      Top             =   3285
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmimage.frx":0E42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlTB"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grayscale"
            Object.ToolTipText     =   "Convert to GrayScale"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pD 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1665
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   3420
      Width           =   240
   End
   Begin VB.HScrollBar HS 
      Height          =   240
      LargeChange     =   15
      Left            =   2205
      SmallChange     =   3
      TabIndex        =   2
      Top             =   4410
      Width           =   2265
   End
   Begin VB.VScrollBar VS 
      Height          =   690
      LargeChange     =   15
      Left            =   4950
      SmallChange     =   3
      TabIndex        =   1
      Top             =   1980
      Width           =   240
   End
   Begin VB.PictureBox pB 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1860
      Left            =   315
      ScaleHeight     =   1860
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   720
      Width           =   825
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'######################################
'WonderHTML 1.2 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
Option Explicit

Private Sub Form_GotFocus()
On Error Resume Next
frmMain.tvW.SelectedItem = frmMain.tvW.Nodes(Caption)
DisableBar
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
TB1.Style = frmMain.TB.Style
Form_Resize
End Sub

Private Sub Form_LostFocus()
If FormsLeft >= 1 Then EnableBar
End Sub

Sub Form_Resize()
On Error Resume Next
PB.Move 0, 0
HS.Move 0, ScaleHeight - HS.Height, ScaleWidth - 15, 15
VS.Move ScaleWidth - VS.Width, 0, 15, ScaleHeight - 15
HS.Enabled = (PB.Width > ScaleWidth)
VS.Enabled = (PB.Height > ScaleHeight - 21)
HS.Max = PB.Width - ScaleWidth
VS.Max = PB.Height - ScaleHeight
PD.Move ScaleWidth - 15, ScaleHeight - 15
pH.Move (PB.Width \ 2) - 1, PB.Height - 1
pV.Move PB.Width - 1, (PB.Height \ 2) - 1
pHV.Move pV.Left, pH.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If FormsLeft >= 1 Then EnableBar
frmMain.ActiveForm.SetFocus
frmMain.ActiveForm.RTF1.SetFocus
End Sub

Private Sub HS_Change()
On Error Resume Next
If HS.Value = 0 Or HS.Value = HS.Max Then Exit Sub
PB.Left = -HS.Value
PB.SetFocus
End Sub

Private Sub HS_Scroll()
HS_Change
End Sub

Private Sub pB_GotFocus()
Form_GotFocus
Form_Resize
End Sub

Private Sub pB_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
Case vbKeyUp
VS.Value = VS.Value - 1
Case vbKeyDown
VS.Value = VS.Value + 1
Case vbKeyLeft
HS.Value = HS.Value - 1
Case vbKeyRight
HS.Value = HS.Value + 1
Case vbKeyPageDown
VS.Value = VS.Value + VS.LargeChange
Case vbKeyPageUp
VS.Value = VS.Value - VS.LargeChange
End Select
End Sub

Private Sub pB_LostFocus()
Form_LostFocus
End Sub

Private Sub TB1_ButtonClick(ByVal Button As MSComctlLib.Button)
ConvGrayScale
End Sub

Private Sub VS_Change()
On Error Resume Next
If VS.Value = 0 Or VS.Value = VS.Max Then Exit Sub
PB.Top = -VS.Value
PB.SetFocus
End Sub

Private Sub VS_Scroll()
VS_Change
End Sub

Function GetPixels(X As Single, y As Single) As String
GetPixels = (X / Screen.TwipsPerPixelX) & ", " & (y / Screen.TwipsPerPixelY)
End Function

Sub EnableBar()
Dim i As Integer
For i = 1 To frmMain.TB.Buttons.count
frmMain.TB.Buttons(i).Enabled = True
Next i
For i = 1 To frmMain.TB2.Buttons.count
frmMain.TB2.Buttons(i).Enabled = True
Next i
For i = 1 To frmMain.tbEdit.Buttons.count
frmMain.tbEdit.Buttons(i).Enabled = True
Next i
For i = 1 To frmMain.tbMore.Buttons.count
frmMain.tbMore.Buttons(i).Enabled = True
Next i
frmMain.cbFonts.Enabled = True
frmMain.cbSizes.Enabled = True
frmMain.cbClrs.Enabled = True
End Sub

Sub DisableBar()
Dim i As Integer
For i = 1 To frmMain.TB.Buttons.count
frmMain.TB.Buttons(i).Enabled = False
Next i
frmMain.TB.Buttons(1).Enabled = True
frmMain.TB.Buttons(2).Enabled = True
frmMain.TB.Buttons(10).Enabled = True
For i = 1 To frmMain.TB2.Buttons.count
frmMain.TB2.Buttons(i).Enabled = False
Next i
For i = 1 To frmMain.tbEdit.Buttons.count
frmMain.tbEdit.Buttons(i).Enabled = False
Next i
For i = 1 To frmMain.tbMore.Buttons.count
frmMain.tbMore.Buttons(i).Enabled = False
Next i
frmMain.cbFonts.Enabled = False
frmMain.cbSizes.Enabled = False
frmMain.cbClrs.Enabled = False
End Sub

Sub ConvGrayScale()
On Error Resume Next
Dim Color As Long, R As Long, G As Long, b As Long
Dim i As Long, ii As Long
For i = 0 To PB.Height
frmMain.SB.Panels(1).Text = "Rendering line " & i & " of " & PB.Height
  For ii = 0 To PB.Width
    Color = GetPixel(PB.hdc, ii, i)
    R = Color Mod 256
    G = (Color \ 256) Mod 256
    b = Color \ 65536
    Color = (R + G + b) \ 3 'average, rounded
    SetPixel PB.hdc, ii, i, IIf(Color > 127, vbWhite, vbBlack)
  Next ii
Next i
End Sub
