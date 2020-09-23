VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCFade 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colour Fade"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "frmCFade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   405
      TabIndex        =   5
      Top             =   2835
      Width           =   1095
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Fade text"
      Default         =   -1  'True
      Height          =   375
      Left            =   405
      TabIndex        =   4
      Top             =   1935
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RTF1 
      Height          =   3300
      Left            =   1935
      TabIndex        =   6
      Top             =   45
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   5821
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmCFade.frx":058A
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1635
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   1725
      Begin VB.CommandButton cmBrowse2 
         Caption         =   "..."
         Height          =   330
         Left            =   1260
         TabIndex        =   3
         Top             =   1170
         Width           =   330
      End
      Begin VB.TextBox txFadeTo 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Text            =   "#FFFFFF"
         Top             =   1170
         Width           =   1095
      End
      Begin VB.CommandButton cmBrowse1 
         Caption         =   "..."
         Height          =   330
         Left            =   1260
         TabIndex        =   1
         Top             =   495
         Width           =   330
      End
      Begin VB.TextBox txFadeFrom 
         Height          =   315
         Left            =   135
         TabIndex        =   0
         Text            =   "#000000"
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fade to:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   945
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fade from:"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.CommandButton cmCopyCode 
      Caption         =   "&Copy code"
      Height          =   375
      Left            =   405
      TabIndex        =   10
      Top             =   2385
      Width           =   1095
   End
End
Attribute VB_Name = "frmCFade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TypeRGB        'This user type is used to convert long decimals to Bytes
    R As Byte
    G As Byte
    b As Byte
End Type

Private Type TypeLong       'This user type is used to convert long decimals to Bytes
    RGBColor As Long
End Type

Private Sub cmBrowse1_Click()
BrowseClr txFadeFrom
End Sub

Private Sub cmBrowse2_Click()
BrowseClr txFadeTo
End Sub

Private Sub cmCopyCode_Click()
Clipboard.Clear
Clipboard.SetText RTF1.Text
Unload Me
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
RTF1.Text = GenerateCode(RTF1.Text, HTML2RGB(txFadeFrom.Text), HTML2RGB(txFadeTo.Text))
End Sub

Private Sub Form_Load()
SetFont Me
End Sub

Function LegalHex(ValH As String) As Boolean
LegalHex = False
Dim sTest As Boolean
sTest = True
Dim i As Integer
If Left(ValH, 1) = "#" And Len(ValH) = 7 Then
  For i = 2 To 7
    If InStr(" 48 49 50 51 52 53 54 55 56 57 58 65 66 67 68 69 70 ", CStr(Asc(UCase(Mid$(ValH, i, 1))))) = 0 Then sTest = False
  Next i
LegalHex = sTest
End If
End Function

Public Function GenerateCode(ByVal Text As String, Color1 As Long, Color2 As Long) As String
Dim count As Long, i As Long
count = Len(Text) + 1: If count = 1 Then Exit Function

Dim r1 As Integer, r2 As Integer
Dim g1 As Integer, g2 As Integer
Dim b1 As Integer, b2 As Integer

Dim Col1 As TypeRGB, Col2 As TypeRGB

Col1 = RGB2typeRGB(Color1)
Col2 = RGB2typeRGB(Color2)

r1 = Col1.R: g1 = Col1.G: b1 = Col1.b
r2 = Col2.R: g2 = Col2.G: b2 = Col2.b

Dim rD As Integer, gD As Integer, bD As Integer
Dim rF As Integer, gF As Integer, bF As Integer

Dim s As String, Code As String, out As String

rD = r1 - r2
gD = g1 - g2
bD = b1 - b2

    For i = 1 To count
        s = Mid(Text, i, 1)
        
        'rf = Int(IIf(rd < 0, r1 + Int(Abs(rd / Ccount * i)), r1 - Int(Abs(rd / Ccount * i)))) Mod 255
        'gf = Int(IIf(gd < 0, r1 + Int(Abs(gd / Ccount * i)), r1 - Int(Abs(gd / Ccount * i)))) Mod 255
        'bf = Int(IIf(bd < 0, r1 + Int(Abs(bd / Ccount * i)), r1 - Int(Abs(bd / Ccount * i)))) Mod 255
        
        rF = Int(r1 - (rD / count * i)) 'Mod 255
        gF = Int(g1 - (gD / count * i)) 'Mod 255
        bF = Int(b1 - (bD / count * i)) 'Mod 255

        If rD = 0 Then rF = r1
        If gD = 0 Then gF = g1
        If bD = 0 Then bF = b1
        
        On Error Resume Next
    Code = "<FONT color=""" & RGB2HTML(RGB(rF, gF, bF)) & """>" & s & "</FONT>"
    
    out = out & Code
    
    Next
    
    GenerateCode = out

End Function

Function RGB2HTML(ByVal RGBColor As Long) As String

Dim tmpCol As TypeRGB

With tmpCol
.R = RGBColor Mod 256
.G = (RGBColor \ 256) Mod 256
.b = RGBColor \ 65536
End With

Dim RString As String, GString As String, BString As String

RString = Trim(Hex$(tmpCol.R))
GString = Trim(Hex$(tmpCol.G))
BString = Trim(Hex$(tmpCol.b))

If Len(RString) = 1 Then RString = "0" & RString
If Len(GString) = 1 Then GString = "0" & GString
If Len(BString) = 1 Then BString = "0" & BString

RGB2HTML = "#" & RString & GString & BString

End Function

Function HTML2RGB(ByVal sHexColor As String) As Long
    Dim lCol As Long, i, n
    If Left(sHexColor, 1) = "#" Then sHexColor = Mid(sHexColor, 2)
    sHexColor = UCase(sHexColor)
    
    For i = 1 To Len(sHexColor) Step 2
        lCol = lCol + Dec(Mid(sHexColor, i, 2)) * 256 ^ n
        n = n + 1
    Next i
    HTML2RGB = lCol
End Function

Private Function RGB2typeRGB(ByVal RGBColor As Long) As TypeRGB
With RGB2typeRGB
.R = RGBColor Mod 256
.G = (RGBColor \ 256) Mod 256
.b = RGBColor \ 65536
End With
End Function

Private Function HTML2typeRGB(ByVal HTMLColor As String) As TypeRGB
Dim temp As String
temp = HTML2RGB(HTMLColor)
Dim out As TypeRGB
out = RGB2typeRGB(temp)
With HTML2typeRGB
.R = out.R
.G = out.G
.b = out.b
End With
End Function

Private Function Dec(ByVal sHex As String) As Long 'Converts Hex to Decimal
    Const HVal = "0123456789ABCDEF"
    Dim iPos As Byte, i As Integer, lDec As Long
    Dim l As Integer, X As Byte
    l = Len(sHex)
    If l > 255 Then Exit Function
    lDec = 0
    For i = l To 1 Step -1
        X = InStr(1, HVal, Mid(sHex, i, 1), vbTextCompare)
        If X = 0 Then Exit Function Else X = X - 1
        lDec = lDec + X * 16 ^ (l - i)
    Next i
    Dec = lDec
End Function

