VERSION 5.00
Begin VB.Form frmSymbols 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Symbol"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmsymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtInsert 
      Height          =   315
      Left            =   675
      TabIndex        =   5
      Top             =   2115
      Width           =   1725
   End
   Begin VB.PictureBox picHolder 
      BackColor       =   &H00FFFFFF&
      Height          =   1910
      Left            =   90
      ScaleHeight     =   1845
      ScaleWidth      =   6090
      TabIndex        =   0
      Top             =   90
      Width           =   6150
      Begin VB.PictureBox pRight 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   5865
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         Top             =   1610
         Width           =   240
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   240
         Left            =   0
         TabIndex        =   9
         Top             =   1605
         Width           =   5870
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1610
         Left            =   5855
         TabIndex        =   6
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblBigDisplay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   2205
         TabIndex        =   8
         Top             =   585
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblSymbols 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.CommandButton cmOK 
      Cancel          =   -1  'True
      Caption         =   "&Insert"
      Default         =   -1  'True
      Height          =   375
      Left            =   4590
      TabIndex        =   1
      Top             =   2115
      Width           =   870
   End
   Begin VB.CommandButton cmNo 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   5490
      TabIndex        =   2
      Top             =   2115
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Insert:"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   1965
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'######################################
'WonderHTML 0.90 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
'this remains the same
Option Explicit
Private CurrentLabel As Integer
Private noperline As Integer
Private numberoflines As Integer
Private linesout As Integer
Private gignore As Boolean
Private minuschars As Integer
Private fntFont As String
Private blnLoadedFonts As Boolean
Private Const BorderWidth As Integer = 0

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyUp
        If Shift = 0 Then
            picHolder_KeyDown KeyCode, Shift
        End If
        KeyCode = 0
    End Select
End Sub

Private Sub Form_Load()

    On Error Resume Next
    SetFont Me
    lblSymbols(0).FontName = "Arial"
    lblBigDisplay.FontName = "Arial"
    lblBigDisplay.FontSize = 14
    blnLoadedFonts = False

    noperline = 0
    
    FillSymbols (0)
    gignore = True
    VScroll1.Max = linesout
    VScroll1.Min = 0
    gignore = False
    ' Set the currently selected label to 0
    CurrentLabel = 0
    
End Sub

Sub FillSymbols(ByVal startnumber As Integer)
On Error Resume Next
Dim i As Integer, currentchar As Integer
Dim NewLeftPos As Long, NewTop As Long
    gignore = False
    ' use minus chars to take away left co-or
    minuschars = 1
    ' number of lines
    numberoflines = 1
    ' hide the first symbol
    lblSymbols(0).Left = -5000
    ' number of lines off screen
    linesout = 0
    ' number of symbols per line
    'noperline = 0
    ' Hide the picture box
    picHolder.Visible = False
    For i = 1 To 223
        ' Load the new symbol label
        On Error Resume Next
        Load lblSymbols(i)
        On Error GoTo 0
        ' change the current char - miss out
        ' the first 32 chars
        currentchar = i + startnumber + 32
        If currentchar > 255 Then Exit For
        ' Set caption to char
        lblSymbols(i).Caption = Chr(currentchar)
        ' New left position
        ' (i - 1) [to allow left to start at 0
        ' - minuschars [to take away the previous
        ' symbols from prev. lines
        ' * (lblsymbols(i).Width - 12)
        ' [To move number from left plus
        ' line width
        NewLeftPos = BorderWidth + ((i) - minuschars) * (lblSymbols(i).Width - 20)
        ' If the new left pos is bigger than
        ' the container width - new symbol
        ' then start a new line
        If NewLeftPos > picHolder.ScaleWidth - lblSymbols(i).Width - VScroll1.Width Then
            ' Add the number of current symbols
            ' minus the one just created
            minuschars = lblSymbols.count - 1
            ' Set the number per line (excluding
            ' current symbol, if it is not set
            ' -1 for currentsymbol
            ' -1 for first label which is not shown
            If noperline = 0 Then noperline = lblSymbols.count - 2
            ' increment the number of lines
            numberoflines = numberoflines + 1
            ' new top position (new line)
            ' lines - 1 [allow for top =0
            ' (lblsymbols(i).Height - 12)
            ' [number of lines - thick line
            NewTop = (numberoflines) * (lblSymbols(i).Height - 20)
            ' If the new top pos is greater than
            ' picHolder bottom line then increment
            ' lines out of screen
            If NewTop + lblSymbols(i).Height > picHolder.Height Then
                linesout = linesout + 1
            End If
            ' Set the new left to include the new
            ' minuschar value
            'NewLeftPos = ((i) - minuschars) * (lblsymbols(i).Width - 12)
            NewLeftPos = BorderWidth + (i - minuschars) * (lblSymbols(i).Width - 20)
        End If
        ' Refresh pic1
        'picHolder.Refresh
        ' set top pos of symbol
        lblSymbols(i).Top = (numberoflines - 1) * (lblSymbols(i).Height - 20)
        ' set new left
        lblSymbols(i).Left = NewLeftPos
        ' make is visible
        lblSymbols(i).Visible = True
    Next
    ' Show the picture again
    picHolder.Visible = True
End Sub

Private Sub lblBigDisplay_DblClick()
    txtInsert.Text = txtInsert.Text & lblSymbols(CurrentLabel).Caption
End Sub

Private Sub lblsymbols_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error GoTo errHandler
    Dim fRed As String
    lblBigDisplay.Left = lblSymbols(Index).Left - ((lblBigDisplay.Width - lblSymbols(Index).Width) / 2)
    lblBigDisplay.Top = lblSymbols(Index).Top - ((lblBigDisplay.Height - lblSymbols(Index).Height) / 2)
    lblBigDisplay.Caption = lblSymbols(Index).Caption
    lblBigDisplay.Visible = True
    CurrentLabel = Index
    fRed = lblSymbols(Index).Caption
    lblAsc.Caption = "Keyboard: Alt+0" & Asc(fRed)
errHandler:
End Sub

Private Sub picHolder_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not Shift = 0 Then Exit Sub
    If KeyCode = vbKeyLeft And Not CurrentLabel = 1 Then
        lblsymbols_MouseDown CurrentLabel - 1, 0, 0, 0, 0
    ElseIf KeyCode = vbKeyRight And Not CurrentLabel = lblSymbols.count - 2 Then
        lblsymbols_MouseDown CurrentLabel + 1, 0, 0, 0, 0
    ElseIf KeyCode = vbKeyUp And CurrentLabel > noperline Then
        lblsymbols_MouseDown CurrentLabel - noperline, 0, 0, 0, 0
    ElseIf KeyCode = vbKeyDown And CurrentLabel < (lblSymbols.count - 2 + noperline) Then
        lblsymbols_MouseDown CurrentLabel + noperline, 0, 0, 0, 0
    End If
End Sub

Private Sub VScroll1_Change()
On Error Resume Next
Dim Label As Label, charstart As Long
    If Not gignore Then
        MousePointer = vbHourglass
        For Each Label In lblSymbols
            If Not Label.Index = 0 Then
                Unload Label
            End If
        Next
        charstart = VScroll1.Value * noperline
        FillSymbols (charstart)
        MousePointer = vbDefault
        lblsymbols_MouseDown CurrentLabel, 0, 0, 0, 0
    End If
    picHolder.SetFocus
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
On Error Resume Next
Dim i As Integer, s As String
For i = 1 To Len(txtInsert.Text)
s = s & "&#" & Asc(Mid(txtInsert.Text, i, 1)) & ";"
Next i
frmMain.ActiveForm.RTF1.SelText = s
frmMain.ActiveForm.RTF1.SelStart = frmMain.ActiveForm.RTF1.SelStart - Len(s)
frmMain.ActiveForm.RTF1.SelLength = s
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmMain.ActiveForm.RTF1.SetFocus
End Sub
