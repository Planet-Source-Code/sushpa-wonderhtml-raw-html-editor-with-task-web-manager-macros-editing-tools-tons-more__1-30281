VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   2430
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1710
      Left            =   0
      Picture         =   "frmTip.frx":1042
      ScaleHeight     =   1710
      ScaleWidth      =   975
      TabIndex        =   6
      Top             =   135
      Width           =   975
   End
   Begin VB.CommandButton cmdNextTip 
      BackColor       =   &H80000004&
      Caption         =   "N&ext"
      Height          =   375
      Left            =   3150
      MousePointer    =   99  'Custom
      Picture         =   "frmTip.frx":1BAC
      TabIndex        =   0
      Top             =   1980
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1125
      ScaleHeight     =   1755
      ScaleWidth      =   3540
      TabIndex        =   3
      Top             =   90
      Width           =   3600
      Begin VB.Timer tmUnload 
         Enabled         =   0   'False
         Interval        =   7500
         Left            =   225
         Top             =   765
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   1750
         Left            =   3315
         TabIndex        =   4
         Top             =   0
         Width           =   225
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H80000018&
         Height          =   1680
         Left            =   45
         TabIndex        =   5
         Top             =   45
         Width           =   3225
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000004&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3960
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1980
      Width           =   780
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "Auto&Show Tips at Startup."
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   2070
      Width           =   2370
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'thanks to VB for the tip-of-day template

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Select a tip at random.
   CurrentTip = Int((Tips.count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

  '  CurrentTip = CurrentTip + 1
 '  If Tips.Count < CurrentTip Then
 '       CurrentTip = 1
  '  End If
  
  If Tips.Item(CurrentTip) = "" Then DoNextTip
    
    ' Show it.
    DisplayCurrentTip
    
End Sub

Private Sub chkLoadTipsAtStartup_Click()
    ' save whether or not this form should be displayed at startup
    SaveValue "StartupTips", chkLoadTipsAtStartup.Value
'    SendMessage chkLoadTipsAtStartup.hwnd, WM_KILLFOCUS, 0, 0&
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
SetFont Me
    
    tmUnload.Enabled = (ReadValue("StartupTips") = 1)
    
    ' Seed Rnd
    Randomize
    
    ' Read in the tips file and display a tip at random.
    If Dir(FullPath(App.Path, TIP_FILE)) = "" Then
    MsgBox "Can't find the tips file.", vbExclamation
    Unload Me
    End If
    
    If LoadTips(FullPath(App.Path, "whtml.tip")) = False Then
    lblTipText.Caption = "That this installation might have been messed up? The tips file could not be found!"
    cmdNextTip.Enabled = False
    chkLoadTipsAtStartup.Enabled = False
    Exit Sub
    End If
    
    DoNextTip

End Sub

Public Sub DisplayCurrentTip()
    If Tips.count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub

Private Sub tmUnload_Timer()
On Error Resume Next
Unload Me
End Sub
