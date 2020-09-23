VERSION 5.00
Begin VB.Form frmImp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Files"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmImp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "<< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3465
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox p2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   2475
      Picture         =   "frmImp.frx":000C
      ScaleHeight     =   630
      ScaleWidth      =   3150
      TabIndex        =   11
      Top             =   135
      Width           =   3180
   End
   Begin VB.ComboBox txWhatToImp 
      Height          =   315
      ItemData        =   "frmImp.frx":0881
      Left            =   2475
      List            =   "frmImp.frx":08E5
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "*.*"
      Top             =   1305
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3270
      Left            =   90
      Picture         =   "frmImp.frx":09C9
      ScaleHeight     =   3210
      ScaleWidth      =   2040
      TabIndex        =   8
      Top             =   90
      Width           =   2100
   End
   Begin VB.CheckBox chSub 
      Caption         =   "Include &Subfolders"
      Height          =   285
      Left            =   2475
      TabIndex        =   7
      Top             =   2295
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Next >>"
      Default         =   -1  'True
      Height          =   375
      Left            =   4635
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   225
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmBrowse 
      Caption         =   "Br&owse..."
      Height          =   375
      Left            =   4230
      TabIndex        =   4
      Top             =   2745
      Width           =   1095
   End
   Begin VB.TextBox txLoc 
      Height          =   315
      Left            =   2475
      TabIndex        =   3
      Top             =   1980
      Width           =   2850
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   2070
      TabIndex        =   9
      Top             =   4590
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.FileListBox Fil1 
      Height          =   1260
      Left            =   2790
      TabIndex        =   10
      Top             =   4635
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Location to Import from:"
      Height          =   195
      Left            =   2475
      TabIndex        =   2
      Top             =   1770
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Files of type:"
      Height          =   195
      Left            =   2475
      TabIndex        =   1
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label SB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter file types, optionally separated by a semi-colon. Click '&Next'."
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   97
      TabIndex        =   13
      Top             =   4185
      Width           =   5700
   End
End
Attribute VB_Name = "frmImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmBrowse_Click()
Dim s As String
s = SelectDir(True)
If s <> "" Then txLoc.Text = s
End Sub

Private Sub cmNo_Click()
If SearchFlag = True Then
SearchFlag = False
ExitFlag = True
Else
Unload Me
End If
End Sub

Private Sub cmOK_Click()
If cmOK.Caption = "&Next >>" Then
  If Right(txLoc.Text, 1) = "\" Then txLoc.Text = Left(txLoc.Text, Len(txLoc.Text) - 1)
  If Dir(txLoc.Text, vbDirectory) = "" Then MsgBox "Can't find location.", vbExclamation: Exit Sub
  If txLoc.Text = "" Then Exit Sub
  Fil1.Pattern = txWhatToImp.Text
  Dir1.Path = txLoc.Text
  SearchFlag = True
  ExitFlag = False
  CopyFiles Dir1.Path, CBool(chSub.Value), Dir1.Path
  ExitFlag = True
  SearchFlag = False
  SB.Caption = "Completed importing files."
  frmMain.mnuWebRefresh_Click
  cmOK.Cancel = "Fini&sh"
Else
  Unload Me
End If
End Sub


Private Sub Dir1_Change()
Fil1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
Caption = frmMain.tvW.Nodes(1).Key
If frmMain.tvW.Nodes.count = 0 Then Caption = App.Path  'that'll spoil the fun
End Sub

Sub MakePath(ByVal Path As String)
On Error Resume Next
If Dir(Up1Level(Path), vbDirectory) <> "" Then MkDir Path: Exit Sub
Dim MinPath As String 'minimum path which exists
Dim MainPath As String
MainPath = Path
Do
Path = Up1Level(Path)
If Dir(Path, vbDirectory) <> "" Then MinPath = Path: Exit Do
Loop
MainPath = Replace(MainPath, MinPath, "", , , vbTextCompare)
MainPath = FirstSlash(MainPath)
Dim st() As String, i As Integer
st = Split(MainPath, "\")
For i = 0 To UBound(st)
MkDir MinPath & "\" & st(i)
MinPath = MinPath & "\" & st(i)
Next i
End Sub


Function CopyFiles(NewPath As String, bSub As Boolean, BackUp As String) As Integer
On Error Resume Next
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String

Dim retval As Integer

    If ExitFlag = True Then
        CopyFiles = True
        Exit Function
    End If
    
    SearchFlag = True           ' Set flag so the user can interrupt.
    CopyFiles = False            ' Set to True if there is an error.
    
    If SearchFlag = False Then
        CopyFiles = True
        Exit Function
    End If
    
    
        If bSub Then DirsToPeek = Dir1.ListCount Else DirsToPeek = 0               ' How many directories below this?
    
    Do While DirsToPeek > 0 And SearchFlag = True
    
        OldPath = Dir1.Path                      ' Save old path for next recursion.
        Dir1.Path = NewPath
        If Dir1.ListCount > 0 Then
            ' Get to the node bottom.
            Dir1.Path = Dir1.list(DirsToPeek - 1)
            retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
            AbandonSearch = CopyFiles(NewPath, bSub, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If Fil1.ListCount Then
        If Len(Dir1.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = Dir1.Path                  ' If at root level, leave as is...
        Else
            ThePath = Dir1.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        Dim s As String
        For ind = 0 To Fil1.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath & Fil1.list(ind)
            s = Replace(entry, txLoc.Text, Caption, , , vbTextCompare)
            If Dir(Up1Level(s, "\"), vbDirectory) = "" Then MakePath Up1Level(s)
            FileCopy entry, FullPath(Up1Level(s, "\"), GetFile(entry))
            SB.Caption = "Copied " & entry
        Next ind
n:
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        Dir1.Path = BackUp
    End If
    Exit Function
End Function

