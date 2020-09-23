VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOpts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WonderHTML: User Preferences"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmopts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageList imlEd 
      Left            =   1350
      Top             =   3375
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
            Picture         =   "frmopts.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3525
      Left            =   90
      TabIndex        =   38
      Top             =   90
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   6218
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Settings"
      TabPicture(0)   =   "frmopts.frx":1060
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txDefW"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbSizes"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chTB"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "opVM(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "opVM(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "opVM(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "imFonts"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chDocTB"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chSB"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chFL"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmINI"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   " Misc"
      TabPicture(1)   =   "frmopts.frx":15FA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chAS"
      Tab(1).Control(1)=   "chFileTree"
      Tab(1).Control(2)=   "chTaskM"
      Tab(1).Control(3)=   "chScV"
      Tab(1).Control(4)=   "chSF"
      Tab(1).Control(5)=   "chAC"
      Tab(1).Control(6)=   "txComments"
      Tab(1).Control(7)=   "txBAt"
      Tab(1).Control(8)=   "chDoc"
      Tab(1).Control(9)=   "chImage"
      Tab(1).Control(10)=   "txUser"
      Tab(1).Control(11)=   "chAutoIndent"
      Tab(1).Control(12)=   "Label20"
      Tab(1).Control(13)=   "Label16"
      Tab(1).Control(14)=   "Label15"
      Tab(1).Control(15)=   "Line2"
      Tab(1).Control(16)=   "Image2"
      Tab(1).Control(17)=   "Label5"
      Tab(1).Control(18)=   "Label4"
      Tab(1).Control(19)=   "Label8"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   " Macros"
      TabPicture(2)   =   "frmopts.frx":1994
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txPrev"
      Tab(2).Control(1)=   "cmInsert"
      Tab(2).Control(2)=   "txSelLen"
      Tab(2).Control(3)=   "txCurPos"
      Tab(2).Control(4)=   "txTTI"
      Tab(2).Control(5)=   "lsKeys"
      Tab(2).Control(6)=   "cbAccel"
      Tab(2).Control(7)=   "Label13"
      Tab(2).Control(8)=   "Label12"
      Tab(2).Control(9)=   "Label11"
      Tab(2).Control(10)=   "Label10"
      Tab(2).Control(11)=   "Label9"
      Tab(2).Control(12)=   "Label6"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   " Editors "
      TabPicture(3)   =   "frmopts.frx":1F2E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pp"
      Tab(3).Control(1)=   "cmRemAss"
      Tab(3).Control(2)=   "cmAddAss"
      Tab(3).Control(3)=   "lsEditors"
      Tab(3).Control(4)=   "Label14"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   " Reports"
      TabPicture(4)   =   "frmopts.frx":22C8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chAllLinks"
      Tab(4).Control(1)=   "txMedia"
      Tab(4).Control(2)=   "cbSpeeds"
      Tab(4).Control(3)=   "txSizeLimit"
      Tab(4).Control(4)=   "chGrid"
      Tab(4).Control(5)=   "Label17"
      Tab(4).Control(6)=   "Image4"
      Tab(4).Control(7)=   "Label19"
      Tab(4).Control(8)=   "Image5"
      Tab(4).Control(9)=   "Image3"
      Tab(4).Control(10)=   "Label18(0)"
      Tab(4).ControlCount=   11
      Begin VB.CheckBox chAS 
         Caption         =   "&AutoSave files"
         Height          =   195
         Left            =   -71040
         TabIndex        =   64
         Top             =   2160
         Width           =   1905
      End
      Begin VB.CheckBox chAllLinks 
         Caption         =   "S&how only broken or unknown hyperlinks"
         Height          =   195
         Left            =   -74865
         TabIndex        =   34
         Top             =   2970
         Width           =   5595
      End
      Begin VB.TextBox txMedia 
         Height          =   315
         Left            =   -74010
         TabIndex        =   33
         Text            =   " gif jpg bmp psd wma mp3 zip ra ram "
         ToolTipText     =   "Use spaces to separate entries. Add spaces at start and end."
         Top             =   2295
         Width           =   2400
      End
      Begin VB.ComboBox cbSpeeds 
         Height          =   315
         ItemData        =   "frmopts.frx":2862
         Left            =   -74010
         List            =   "frmopts.frx":2881
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1575
         Width           =   1230
      End
      Begin VB.TextBox txSizeLimit 
         Height          =   315
         Left            =   -73965
         TabIndex        =   31
         Text            =   "75"
         Top             =   810
         Width           =   465
      End
      Begin VB.CheckBox chGrid 
         Caption         =   "Display &Gridlines while viewing reports"
         Height          =   195
         Left            =   -74865
         TabIndex        =   35
         Top             =   3195
         Width           =   5595
      End
      Begin VB.TextBox txPrev 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         HideSelection   =   0   'False
         Left            =   -71805
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   63
         Top             =   1485
         Width           =   2535
      End
      Begin VB.CheckBox chFileTree 
         Caption         =   "Main FileTree"
         Height          =   195
         Left            =   -71040
         TabIndex        =   17
         Top             =   1440
         Width           =   1905
      End
      Begin VB.CheckBox chTaskM 
         Caption         =   "Task Manager"
         Height          =   195
         Left            =   -71040
         TabIndex        =   16
         Top             =   1215
         Width           =   1905
      End
      Begin VB.CheckBox chScV 
         Caption         =   "Script Outline"
         Height          =   195
         Left            =   -71040
         TabIndex        =   15
         Top             =   990
         Width           =   1905
      End
      Begin VB.CommandButton cmINI 
         Height          =   375
         Left            =   4635
         MaskColor       =   &H000000FF&
         Picture         =   "frmopts.frx":28D9
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2970
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chFL 
         Caption         =   "&Flat Buttons"
         Height          =   195
         Left            =   1980
         TabIndex        =   5
         Top             =   1530
         Width           =   1635
      End
      Begin VB.CheckBox chSB 
         Caption         =   "Show &StatusBar"
         Height          =   195
         Left            =   1980
         TabIndex        =   3
         Top             =   1305
         Width           =   1680
      End
      Begin VB.CheckBox chDocTB 
         Caption         =   "Show H&TMLBar"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   1530
         Width           =   2000
      End
      Begin VB.PictureBox pp 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   510
         Left            =   -74820
         ScaleHeight     =   510
         ScaleWidth      =   510
         TabIndex        =   55
         Top             =   2700
         Width           =   510
      End
      Begin VB.CommandButton cmRemAss 
         Caption         =   "R&emove"
         Height          =   375
         Left            =   -70275
         TabIndex        =   30
         Top             =   2745
         Width           =   1050
      End
      Begin VB.CommandButton cmAddAss 
         Caption         =   "&Add..."
         Height          =   375
         Left            =   -71355
         TabIndex        =   29
         Top             =   2745
         Width           =   1005
      End
      Begin MSComctlLib.ListView lsEditors 
         Height          =   2220
         Left            =   -74865
         TabIndex        =   28
         Top             =   450
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   3916
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlEd"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "ext"
            Text            =   "Extension"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "path"
            Text            =   "Program Path"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.ComboBox imFonts 
         Height          =   315
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   765
         Width           =   2310
      End
      Begin VB.CommandButton cmInsert 
         Caption         =   "&Update Macro"
         Height          =   375
         Left            =   -73290
         TabIndex        =   27
         Top             =   2700
         Width           =   1365
      End
      Begin VB.TextBox txSelLen 
         Height          =   315
         Left            =   -73290
         TabIndex        =   26
         Text            =   "0"
         Top             =   2295
         Width           =   1365
      End
      Begin VB.TextBox txCurPos 
         Height          =   315
         Left            =   -73290
         TabIndex        =   25
         Text            =   "0"
         Top             =   1665
         Width           =   1365
      End
      Begin VB.TextBox txTTI 
         Height          =   315
         Left            =   -73290
         TabIndex        =   24
         Top             =   1035
         Width           =   4020
      End
      Begin VB.ListBox lsKeys 
         Height          =   2010
         ItemData        =   "frmopts.frx":3D6B
         Left            =   -74865
         List            =   "frmopts.frx":3DC0
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cbAccel 
         Height          =   315
         ItemData        =   "frmopts.frx":3EA0
         Left            =   -73560
         List            =   "frmopts.frx":3EA7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   405
         Width           =   1455
      End
      Begin VB.CheckBox chSF 
         Caption         =   "&Select definition"
         Height          =   195
         Left            =   -71040
         TabIndex        =   21
         Top             =   3150
         Width           =   1905
      End
      Begin VB.OptionButton opVM 
         Caption         =   "P&rinter DC"
         Height          =   195
         Index           =   2
         Left            =   4005
         TabIndex        =   9
         Top             =   2025
         Width           =   1815
      End
      Begin VB.OptionButton opVM 
         Caption         =   "Wo&rd wrap"
         Height          =   195
         Index           =   1
         Left            =   4005
         TabIndex        =   8
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton opVM 
         Caption         =   "Continuous"
         Height          =   195
         Index           =   0
         Left            =   4005
         TabIndex        =   7
         Top             =   1575
         Width           =   1815
      End
      Begin VB.CheckBox chAC 
         Caption         =   "Use &AutoLink"
         Height          =   195
         Left            =   -71040
         TabIndex        =   19
         Top             =   2700
         Width           =   1905
      End
      Begin VB.TextBox txComments 
         Height          =   315
         Left            =   -74820
         TabIndex        =   13
         Top             =   2325
         Width           =   3570
      End
      Begin VB.TextBox txBAt 
         Height          =   315
         Left            =   -74820
         TabIndex        =   12
         Text            =   "alink=""#FF0000"" vlink=""#000080"" link=""#0000FF"""
         Top             =   1560
         Width           =   3570
      End
      Begin VB.CheckBox chDoc 
         Caption         =   "HTML Outline"
         Height          =   195
         Left            =   -71040
         TabIndex        =   14
         Top             =   765
         Width           =   1905
      End
      Begin VB.CheckBox chImage 
         Caption         =   "Image Viewer"
         Height          =   195
         Left            =   -71040
         TabIndex        =   18
         Top             =   1935
         Width           =   1905
      End
      Begin VB.TextBox txUser 
         Height          =   315
         Left            =   -74325
         TabIndex        =   11
         Top             =   825
         Width           =   3075
      End
      Begin VB.CheckBox chTB 
         Caption         =   "Show &ToolBar"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   1305
         Width           =   2000
      End
      Begin VB.ComboBox cbSizes 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   765
         Width           =   600
      End
      Begin VB.TextBox txDefW 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   2160
         Width           =   2850
      End
      Begin VB.CheckBox chAutoIndent 
         Caption         =   "Use Auto&Indent"
         Height          =   195
         Left            =   -71040
         TabIndex        =   20
         Top             =   2925
         Width           =   1905
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "More wonders"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71040
         TabIndex        =   62
         Top             =   1710
         Width           =   1200
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   """Media"" files are files of type:"
         Height          =   195
         Left            =   -74055
         TabIndex        =   61
         Top             =   2070
         Width           =   2115
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -74595
         Picture         =   "frmopts.frx":3EB7
         Top             =   2115
         Width           =   480
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Assume connection speed of:"
         Height          =   195
         Left            =   -74010
         TabIndex        =   60
         Top             =   1350
         Width           =   2115
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   -74640
         Picture         =   "frmopts.frx":4B81
         Top             =   1395
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -74640
         Picture         =   "frmopts.frx":65FB
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   """Large"" files are files over (KB):"
         Height          =   195
         Index           =   0
         Left            =   -74010
         TabIndex        =   59
         Top             =   585
         Width           =   2250
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "As you type..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71040
         TabIndex        =   57
         Top             =   2430
         Width           =   1140
      End
      Begin VB.Label Label15 
         Caption         =   "WonderPane"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71040
         TabIndex        =   56
         Top             =   495
         Width           =   1680
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   -74685
         X2              =   -71490
         Y1              =   3015
         Y2              =   3015
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74865
         Picture         =   "frmopts.frx":72C5
         Top             =   675
         Width           =   480
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "The selected file's path is automatically passed as a parameter."
         Height          =   195
         Left            =   -74250
         TabIndex        =   54
         Top             =   3240
         Width           =   4575
      End
      Begin VB.Label Label13 
         Caption         =   "Selection length:"
         Height          =   195
         Left            =   -73290
         TabIndex        =   53
         Top             =   2070
         Width           =   1320
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cursor position:"
         Height          =   195
         Left            =   -73290
         TabIndex        =   52
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Use \n to signify carriage return characters."
         Height          =   195
         Left            =   -74460
         TabIndex        =   51
         Top             =   3195
         Width           =   3165
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "When key is presed, insert this text:"
         Height          =   195
         Left            =   -73290
         TabIndex        =   50
         Top             =   810
         Width           =   2625
      End
      Begin VB.Label Label9 
         Caption         =   "KeyCode:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   49
         Top             =   855
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Accelerator key:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   48
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "View &Mode:"
         Height          =   195
         Left            =   4005
         TabIndex        =   47
         Top             =   1305
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Insert this as a comment in all documents:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   46
         Top             =   2130
         Width           =   3015
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Insert these attributes in the BODY tag:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   45
         Top             =   1365
         Width           =   2880
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "A&uthor name to insert in META tags:"
         Height          =   195
         Left            =   -74325
         TabIndex        =   44
         Top             =   600
         Width           =   2610
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Editor F&ont:"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Specify a web to be loaded automatically on startup. Leave this box blank to avoid using this feature."
         Height          =   645
         Left            =   180
         TabIndex        =   42
         Top             =   2655
         Width           =   3345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Def&ault web:"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   1935
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   3735
         X2              =   3735
         Y1              =   3195
         Y2              =   540
      End
   End
   Begin VB.CommandButton cmOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3975
      MaskColor       =   &H000000FF&
      Picture         =   "frmopts.frx":7B8F
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3660
      UseMaskColor    =   -1  'True
      Width           =   1005
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5055
      TabIndex        =   41
      Top             =   3660
      Width           =   960
   End
   Begin VB.ComboBox cbSort 
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   37
      Top             =   5100
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lbInfo 
      Caption         =   "© 2001 Boomerang Software Corp."
      Height          =   240
      Left            =   135
      TabIndex        =   58
      Top             =   3735
      Width           =   3525
   End
   Begin VB.Label Test 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   1545
      TabIndex        =   36
      Top             =   5055
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "frmOpts"
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
'put in more options, bigger dialog: 15 Sep 2001
Option Explicit

Private Sub cbAccel_GotFocus()
frmMain.SB.Panels(1).Text = "Selected Modifier (accelerator) Key: Ctrl+Shift"
'SSTab1.Tab = 2
End Sub

Private Sub cbSizes_GotFocus()
frmMain.SB.Panels(1).Text = "Select the font size to use, to display HTML Documents."
End Sub

Private Sub chAC_Click()
frmMain.SB.Panels(1).Text = "Specify whether or not to convert plain URLs to links while typing"
End Sub

Private Sub chDocTB_Click()
frmMain.SB.Panels(1).Text = "Shows or hides the HTML toolbar."
End Sub

Private Sub chFileTree_Click()
chDoc.Enabled = CBool(chFileTree.Value)
chScV.Enabled = CBool(chFileTree.Value)
chTaskM.Enabled = CBool(chFileTree.Value)
End Sub

Private Sub chFL_Click()
frmMain.SB.Panels(1).Text = "Toggle the style of the main tool bar."
End Sub

Private Sub chSB_Click()
frmMain.SB.Panels(1).Text = "Toggle the visibility of the status bar."
End Sub

Private Sub chTB_Click()
frmMain.SB.Panels(1).Text = "Toggle the visibility of the main tool bar."
End Sub

Private Sub cmAddAss_Click()
On Error GoTo hell
Dim ass As String
ass = AddAssociation()
If ass = "" Or ass = ";" Then Exit Sub
Dim asses() As String
'don't laugh
asses = Split(ass, ";")
lsEditors.ListItems.Add , asses(0), asses(0), , 1
lsEditors.ListItems(asses(0)).ListSubItems.Add 1, asses(1), GetFile(asses(1)), , asses(1)
lsEditors.ListItems(asses(0)).ListSubItems.Add 2, , asses(2)
Exit Sub
hell:
On Error Resume Next
lsEditors.SelectedItem = lsEditors.ListItems(asses(0))
lsEditors.SetFocus
lsEditors_ItemClick lsEditors.SelectedItem
End Sub

Private Sub cmINI_Click()
ShellExecute hwnd, "open", "settings.ini", "", App.Path, 10
End Sub

Private Sub cmInsert_Click()
On Error Resume Next
SaveValue "+" & Left(lsKeys.list(lsKeys.ListIndex), 2) & "Text", txTTI.Text, "Macros"
SaveValue "+" & Left(lsKeys.list(lsKeys.ListIndex), 2) & "Sel", txCurPos.Text & ";" & txSelLen.Text, "Macros"
lsKeys.SetFocus
End Sub

Private Sub cmNo_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
'On Error Resume Next
MousePointer = 11
lbInfo.Caption = "Saving selected settings..."
If imFonts.Text = "" Then imFonts.ListIndex = 0
SaveValue "DocumentTree", CBool(chDoc.Value)
SaveValue "AutoSave", CBool(chAS.Value)
SaveValue "HTMLBar", CBool(chDocTB.Value)
SaveValue "FontName", imFonts.Text, "Documents"
SaveValue "Media", txMedia.Text, "Reports"
SaveValue "BrokenOnly", CBool(chAllLinks.Value), "Reports"
SaveValue "AutoIndent", CBool(chAutoIndent.Value), "Documents"
SaveValue "FontSize", cbSizes.Text, "Documents"
SaveValue "FlatBar", chFL.Value
SaveValue "WebDefault", txDefW.Text
SaveValue "ToolBar", chTB.Value
SaveValue "StatusBar", chSB.Value
SaveValue "SizeLimit", txSizeLimit.Text, "Reports"
SaveValue "ConnectionSpeed", cbSpeeds.Text, "Reports"
SaveValue "Gridlines", CBool(chGrid.Value), "Reports"
SaveValue "BodyAttrib", txBAt.Text, "Documents"
SaveValue "SelectFind", CBool(chSF.Value), "Documents"
SaveValue "Comments", txComments.Text, "Documents"
SaveValue "Author", txUser.Text, "Documents"
SaveValue "ImageViewer", chImage.Value
SaveValue "ViewMode", GetVM(), "Documents"
SaveValue "AutoLink", CBool(chAC.Value), "Documents"
SaveValue "+" & Left(lsKeys.list(lsKeys.ListIndex), 2) & "Text", txTTI.Text, "Macros"
SaveValue "+" & Left(lsKeys.list(lsKeys.ListIndex), 2) & "Sel", txCurPos.Text & ";" & txSelLen.Text, "Macros"
SaveAssoc
MousePointer = 0
Unload Me
frmMain.SetPrefs
frmMain.GetPrefs
End Sub

Private Sub cmRemAss_Click()
On Error Resume Next
If MsgBox("Remove the association for ." & lsEditors.SelectedItem.Text & " files?", vbYesNo + vbExclamation, "Remove") = vbNo Then Exit Sub
lsEditors.ListItems.Remove lsEditors.SelectedItem.Index
End Sub

Private Sub Form_Load()
'On Error Resume Next
Call AddFonts
SetFont Me
GetAssoc
chAC.Value = CBinary(ReadValue("AutoLink", False, "Documents"))
chAS.Value = CBinary(ReadValue("AutoSave", True))
chDoc.Value = CBinary(ReadValue("DocumentTree", frmMain.SSTab1.TabVisible(1)))
chImage.Value = ReadValue("ImageViewer", 1)
chDocTB.Value = CBinary(ReadValue("HTMLBar", frmMain.CBR.Bands(4).Visible))
cbSizes.Text = CInt(ReadValue("FontSize", , "Documents"))
txDefW.Text = ReadValue("WebDefault")
chAllLinks.Value = CBinary(ReadValue("BrokenOnly", False, "Reports"))
chGrid.Value = CBinary(ReadValue("Gridlines", False, "Reports"))
txMedia.Text = ReadValue("Media", "", "Reports")
txSizeLimit.Text = ReadValue("SizeLimit", , "Reports")
cbSpeeds.Text = ReadValue("ConnectionSpeed", , "Reports")
chFL.Value = CBinary(ReadValue("FlatBar", frmMain.TB.Style = tbrFlat))
chTB.Value = CBinary(ReadValue("ToolBar", frmMain.CBR.Bands(1).Visible))
chSF.Value = CBinary(ReadValue("SelectFind", , "Documents"))
chSB.Value = CBinary(ReadValue("StatusBar", frmMain.SB.Visible))
chScV.Value = CBinary(ReadValue("ScriptView", frmMain.SSTab1.TabVisible(2)))
chTaskM.Value = CBinary(ReadValue("TaskView", frmMain.SSTab1.TabVisible(3)))
chFileTree.Value = CBinary(ReadValue("FileTree", frmMain.SSTab1.TabVisible(0)))
chAutoIndent.Value = CBinary(ReadValue("AutoIndent", , "Documents"))
txBAt.Text = ReadValue("BodyAttrib", , "Documents")
txComments.Text = ReadValue("Comments", "<!-- Created with WonderHTML 2001 //-->", "Documents")
txUser.Text = ReadValue("Author", , "Documents")
opVM(CInt(ReadValue("ViewMode", , "Documents"))).Value = True
cbAccel.ListIndex = 0
lsKeys.ListIndex = 0
End Sub

Sub AddFonts()
On Error Resume Next
Dim i As Integer, IC As Integer
For i = 0 To Screen.FontCount - 1
imFonts.AddItem Screen.Fonts(i)
Next i
For i = 8 To 72 Step 2
cbSizes.AddItem i
Next i
imFonts.Text = ReadValue("FontName", , "Documents")
End Sub

Function GetVM() As Integer
Dim i As Integer
For i = 0 To 2
If opVM(i).Value Then GetVM = i: Exit Function
Next i
End Function


Private Sub imFonts_GotFocus()
frmMain.SB.Panels(1).Text = "Select the font face to use, to display HTML Documents."
'SSTab1.Tab = 0
End Sub

Private Sub lsEditors_DblClick()
On Error Resume Next
Dim asses() As String
Load frmAss
frmAss.txApp.Text = lsEditors.SelectedItem.ListSubItems(1).Key
frmAss.txDesc.Text = lsEditors.SelectedItem.ListSubItems(2).Text
frmAss.txExt.Text = lsEditors.SelectedItem.Text
frmAss.txApp.SetFocus
frmAss.txExt.SetFocus
frmAss.txExt.Enabled = False
frmAss.txExt.BackColor = Me.BackColor
frmAss.Show vbModal
If assoc_text = "" Then Exit Sub
asses = Split(assoc_text, ";")
lsEditors.SelectedItem.Text = asses(0)
lsEditors.SelectedItem.ListSubItems(1).Key = asses(1)
lsEditors.SelectedItem.ListSubItems(1).Text = GetFile(asses(1))
lsEditors.SelectedItem.ListSubItems(2).Text = asses(2)
hell:
End Sub

Private Sub lsEditors_GotFocus()
'SSTab1.Tab = 3
End Sub

Private Sub lsEditors_ItemClick(ByVal Item As MSComctlLib.ListItem)
pp.Cls
PaintIcon Item.ListSubItems(1).Key, pp
frmMain.SB.Panels(1).Text = Chr(147) & FileType(Item.Text) & Chr(148) & " types open with " & GetFile(Item.ListSubItems(1).Key) & "."
End Sub

Private Sub lsKeys_Click()
On Error Resume Next
Dim sel() As String, s As String, l As Long
l = CLng(Left(lsKeys.list(lsKeys.ListIndex), 2))
s = ReadValue("+" & l & "Text", "", "Macros")
txTTI.Text = s
s = ReadValue("+" & l & "Sel", ";", "Macros")
sel = Split(s, ";")
txCurPos.Text = sel(0)
If txCurPos.Text = "" Then txCurPos.Text = 0
txSelLen.Text = sel(1)
If txSelLen.Text = "" Then txSelLen.Text = 0
txPrev.Text = Replace(txTTI.Text, "\n", vbCrLf)
txPrev.SelStart = Len(txPrev.Text) - CInt(txCurPos.Text)
txPrev.SelLength = CInt(txSelLen.Text)
txPrev.SetFocus
End Sub

Private Sub lsKeys_GotFocus()
frmMain.SB.Panels(1).Text = "Select the key to be used with the Ctrl+Shift combination."
End Sub

Private Sub opVM_Click(Index As Integer)
frmMain.SB.Panels(1).Text = "Set the view mode accordingly, and specify how text should wrap."
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Select Case SSTab1.Tab
Case 0
imFonts.SetFocus
Case 1
txUser.SetFocus
Case 2
lsKeys.SetFocus
Case 3
lsEditors.SetFocus
Case 4
txSizeLimit.SetFocus
End Select
End Sub

Private Sub txCurPos_GotFocus()
frmMain.SB.Panels(1).Text = "Enter the number of positions to move the cursor back, after the text has been inserted."
End Sub

Private Sub txCurPos_LostFocus()
On Error Resume Next
If IsNumeric(txCurPos.Text) = False Or txCurPos.Text = "" Or CLng(txCurPos.Text) > Len(txTTI.Text) Then txCurPos.Text = 0
End Sub

Private Sub txDefW_GotFocus()
frmMain.SB.Panels(1).Text = "Specify the path of the web to load on startup."
End Sub

Private Sub txSelLen_GotFocus()
frmMain.SB.Panels(1).Text = "Enter the number of characters to select, after the text has been inserted."
End Sub

Private Sub txSelLen_LostFocus()
On Error Resume Next
If IsNumeric(txSelLen.Text) = False Or txSelLen.Text = "" Or CLng(txSelLen.Text) > Len(txTTI.Text) Then txSelLen.Text = 0
End Sub

Sub GetAssoc()
On Error Resume Next
Dim i As Integer, assoc As String
Dim ss() As String, vals() As String

assoc = ReadValue("Config", "", "Editors", FullPath(App.Path, "editors.inf"))
ss = Split(assoc, ",")

For i = 0 To UBound(ss)
assoc = ReadValue(ss(i), "", "Editors", FullPath(App.Path, "editors.inf"))
If assoc = "" Then GoTo n
vals = Split(assoc, "|")
lsEditors.ListItems.Add , ss(i), ss(i), , 1
lsEditors.ListItems(ss(i)).ListSubItems.Add 1, vals(1), GetFile(vals(1)), , vals(1)
lsEditors.ListItems(ss(i)).ListSubItems.Add 2, , vals(0)
n:
Next i

lsEditors_ItemClick lsEditors.ListItems(1)

End Sub

Sub SaveAssoc()
On Error Resume Next
Dim i As Integer, assoc As String
Dim ss() As String

For i = 1 To lsEditors.ListItems.count - 1
assoc = assoc & lsEditors.ListItems(i).Text & ","
Next i
assoc = assoc & lsEditors.ListItems(lsEditors.ListItems.count).Text

SaveValue "Config", assoc, "Editors", FullPath(App.Path, "editors.inf")

ss = Split(assoc, ",")
For i = 0 To UBound(ss)
SaveValue ss(i), lsEditors.ListItems(i + 1).ListSubItems(2).Text & "|" & lsEditors.ListItems(i + 1).ListSubItems(1).Key, "Editors", FullPath(App.Path, "editors.inf")
Next i

End Sub

Private Sub txTTI_GotFocus()
On Error Resume Next
frmMain.SB.Panels(1).Text = "Enter the text to insert upon pressing Ctrl+Shift+“" & Chr(CLng(Left(lsKeys.list(lsKeys.ListIndex), 2))) & Chr(148)
End Sub

Private Sub txUser_GotFocus()
'SSTab1.Tab = 1
End Sub
