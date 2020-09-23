VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "WonderHTML"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "frmMain.frx":1042
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl3.CoolBar CBR 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1323
      BandCount       =   4
      _CBWidth        =   11880
      _CBHeight       =   750
      _Version        =   "6.0.8169"
      Child1          =   "TB"
      MinWidth1       =   3990
      MinHeight1      =   330
      Width1          =   3990
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "tbEdit"
      MinWidth2       =   4305
      MinHeight2      =   330
      Width2          =   4305
      NewRow2         =   0   'False
      Child3          =   "tbMore"
      MinWidth3       =   2805
      MinHeight3      =   330
      Width3          =   1095
      NewRow3         =   0   'False
      Child4          =   "TB2"
      MinWidth4       =   9240
      MinHeight4      =   330
      Width4          =   9240
      NewRow4         =   -1  'True
      Begin MSComctlLib.Toolbar tbMore 
         Height          =   330
         Left            =   8985
         TabIndex        =   16
         Top             =   30
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlTB"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "lastpos"
               Object.ToolTipText     =   "Last Position (F9)"
               ImageKey        =   "refresh"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "define"
               Object.ToolTipText     =   "Definition (F2)"
               ImageKey        =   "defEx"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Find text in document (Ctrl+F)"
               ImageKey        =   "find"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "findnext"
               Object.ToolTipText     =   "Find next occurence (F3)"
               ImageKey        =   "findnext"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "ascii"
               Object.ToolTipText     =   "Convert HTML to text (F8)"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TB2 
         Height          =   330
         Left            =   165
         TabIndex        =   15
         Top             =   390
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "imlTB2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "12"
               Style           =   4
               Object.Width           =   3930
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "3"
               Object.ToolTipText     =   "Make selected text bold"
               Object.Tag             =   "<B></B>"
               ImageKey        =   "bold"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "4"
               Object.ToolTipText     =   "Make selected text Italic"
               Object.Tag             =   "<EM></EM>"
               ImageKey        =   "italic"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "3"
               Object.ToolTipText     =   "Underline selected text"
               Object.Tag             =   "<U></U>"
               ImageKey        =   "underline"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "14"
               Object.ToolTipText     =   "Create a bulleted list"
               Object.Tag             =   "<UL>|    <LI></LI>|</UL>"
               ImageKey        =   "bullets"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "14"
               Object.ToolTipText     =   "Create a numbered list"
               Object.Tag             =   "<OL>|    <LI></LI>|</OL>"
               ImageKey        =   "numbers"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "6"
               Object.ToolTipText     =   "Create a hyperlink"
               Object.Tag             =   "<A href=""""></A>"
               ImageKey        =   "link"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "14"
               Object.ToolTipText     =   "Insert a Java applet"
               Object.Tag             =   "<APPLET code="""" codebase="""" width=""128"" height=""128"">|Your browser is not java-enabled.|</APPLET>"
               ImageKey        =   "applet"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Insert default date and time"
               Object.Tag             =   "!"
               ImageKey        =   "time"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "72"
               Object.ToolTipText     =   "Insert a script block"
               Object.Tag             =   "<SCRIPT language=""Javascript"" type=""text/javascript"">|<!-- hide code||//end code -->|</SCRIPT>"
               ImageKey        =   "script"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "22"
               Object.ToolTipText     =   "Increase the indentation level of text"
               Object.Tag             =   "    <BLOCKQUOTE>|    |    </BLOCKQUOTE>"
               ImageKey        =   "indent"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.ComboBox cbClrs 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":29FC
            Left            =   2880
            List            =   "frmMain.frx":2A33
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   15
            Width           =   1050
         End
         Begin VB.ComboBox cbSizes 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":2AB0
            Left            =   2070
            List            =   "frmMain.frx":2B02
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   15
            Width           =   780
         End
         Begin VB.ComboBox cbFonts 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":2B80
            Left            =   0
            List            =   "frmMain.frx":2B82
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   15
            Width           =   2040
         End
      End
      Begin MSComctlLib.Toolbar tbEdit 
         Height          =   330
         Left            =   4380
         TabIndex        =   14
         Top             =   30
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlTB"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "undo"
               Object.ToolTipText     =   "Undo the last action (Ctrl+Z)"
               ImageKey        =   "undo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "redo"
               Object.ToolTipText     =   "Redo the previously undone action (Ctrl+Y)"
               ImageKey        =   "redo"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "cut"
               Object.ToolTipText     =   "Cut and place text on the clipboard (Ctrl+X)"
               ImageKey        =   "cut"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "copy"
               Object.ToolTipText     =   "Copy and place text on the clipboard (Ctrl+C)"
               ImageKey        =   "copy"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "paste"
               Object.ToolTipText     =   "Paste text from the clipboard (Ctrl+V)"
               ImageKey        =   "paste"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "edit"
               Object.ToolTipText     =   "Edit the document's properties (Alt+Enter)"
               ImageKey        =   "html"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "symbol"
               Object.ToolTipText     =   "Insert symbols"
               ImageKey        =   "symbol"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "date"
               Object.ToolTipText     =   "Insert date/time"
               ImageIndex      =   27
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "cpick"
               Object.ToolTipText     =   "Pick and choose colours"
               ImageIndex      =   28
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "cfade"
               Object.ToolTipText     =   "Generate faded colour text"
               ImageKey        =   "clr"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "opensrc"
               Object.ToolTipText     =   "Open SRC document"
               ImageKey        =   "open"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TB 
         Height          =   330
         Left            =   165
         TabIndex        =   13
         Tag             =   "The main toolbar contains quick features for text editing."
         Top             =   30
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlTB"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "New file"
               Object.ToolTipText     =   "New document (Ctrl+N)"
               Object.Tag             =   "New"
               ImageKey        =   "new"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Blank"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Open file"
               Object.ToolTipText     =   "Open document (Ctrl+O)"
               Object.Tag             =   "Open"
               ImageKey        =   "open"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   9
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "browse"
                     Text            =   "Local file..."
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "remote"
                     Text            =   "Remote file..."
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "sep"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "Save file"
               Object.ToolTipText     =   "Save the document (Ctrl+S)"
               Object.Tag             =   "Save"
               ImageKey        =   "save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "Print file"
               Object.ToolTipText     =   "Print the document (Ctrl+P)"
               Object.Tag             =   "Print"
               ImageKey        =   "print"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "See a print preview"
               Object.Tag             =   "Test"
               ImageKey        =   "pprev"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "Close file"
               Object.ToolTipText     =   "Close (Ctrl+F4)"
               Object.Tag             =   "Close"
               ImageKey        =   "close"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "&Close"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Close &all"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Web manager"
               Object.ToolTipText     =   "Web tools"
               Object.Tag             =   "Web"
               ImageKey        =   "web"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Description     =   "Preview"
               Object.ToolTipText     =   "Preview (F5)"
               Object.Tag             =   "View"
               ImageKey        =   "prev"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Browsers..."
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         MousePointer    =   99
      End
   End
   Begin MSComctlLib.ImageList imlTasks 
      Left            =   4320
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B84
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4550
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5708
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":675C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7438
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlBusy 
      Left            =   5445
      Top             =   3915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A144
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A6E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5175
      Top             =   1395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "html"
      DialogTitle     =   "WonderHTML"
      Filter          =   $"frmMain.frx":AC7C
   End
   Begin VB.PictureBox pLeft 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   0
      ScaleHeight     =   7260
      ScaleWidth      =   3270
      TabIndex        =   2
      Top             =   750
      Width           =   3270
      Begin VB.PictureBox pS 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5370
         Left            =   3150
         MouseIcon       =   "frmMain.frx":AD17
         MousePointer    =   99  'Custom
         ScaleHeight     =   5370
         ScaleWidth      =   45
         TabIndex        =   7
         Top             =   -90
         Width           =   45
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6270
         Left            =   0
         TabIndex        =   3
         Tag             =   "Files   DocumentScripts Tasks   "
         Top             =   0
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   11060
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   556
         TabMaxWidth     =   882
         ShowFocusRect   =   0   'False
         MouseIcon       =   "frmMain.frx":AE69
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   " Files"
         TabPicture(0)   =   "frmMain.frx":AE85
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "tvW"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Fldr"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Fil"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txHL"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   " "
         TabPicture(1)   =   "frmMain.frx":B21F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "tvD"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " "
         TabPicture(2)   =   "frmMain.frx":B5B9
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "tvS"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   " "
         TabPicture(3)   =   "frmMain.frx":B953
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lvTasks"
         Tab(3).Control(1)=   "tbView"
         Tab(3).ControlCount=   2
         Begin VB.TextBox txHL 
            BorderStyle     =   0  'None
            Height          =   600
            HideSelection   =   0   'False
            Left            =   720
            MultiLine       =   -1  'True
            TabIndex        =   11
            Text            =   "frmMain.frx":BEED
            Top             =   4995
            Visible         =   0   'False
            Width           =   1365
         End
         Begin MSComctlLib.TreeView tvS 
            Height          =   4560
            Left            =   -74910
            TabIndex        =   10
            Tag             =   "The ScriptView displays the script functions and variables in your documents, in a hierachial manner."
            Top             =   405
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   8043
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   326
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imlJS"
            Appearance      =   1
         End
         Begin MSComctlLib.Toolbar tbView 
            Height          =   570
            Left            =   -74910
            TabIndex        =   9
            Top             =   405
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   1005
            ButtonWidth     =   1032
            ButtonHeight    =   1005
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "imlTasks"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.ToolTipText     =   "Create or manage tasks for this file"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.ToolTipText     =   "Open this file for editing"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Refresh the task list"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "View a report for the web"
                  ImageIndex      =   6
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvTasks 
            Height          =   3795
            Left            =   -74910
            TabIndex        =   8
            Tag             =   "The TaskView scans all files and lists those, which are commented or have pending tasks."
            Top             =   1035
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   6694
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            PictureAlignment=   2
            _Version        =   393217
            Icons           =   "iml32"
            SmallIcons      =   "imlTV"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "File"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Details"
               Object.Width           =   2540
            EndProperty
            Picture         =   "frmMain.frx":BF19
         End
         Begin VB.FileListBox Fil 
            BackColor       =   &H80000018&
            Height          =   1260
            Hidden          =   -1  'True
            Left            =   270
            System          =   -1  'True
            TabIndex        =   6
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.DirListBox Fldr 
            Height          =   990
            Left            =   495
            TabIndex        =   5
            Top             =   750
            Visible         =   0   'False
            Width           =   1590
         End
         Begin MSComctlLib.TreeView tvW 
            Height          =   4560
            Left            =   90
            TabIndex        =   0
            Tag             =   $"frmMain.frx":2F06F
            Top             =   405
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   8043
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   335
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "imlTV"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView tvD 
            Height          =   4635
            Left            =   -74910
            TabIndex        =   4
            Tag             =   $"frmMain.frx":2F0FA
            Top             =   405
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   8176
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   335
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imlHTML"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList imlJS 
      Left            =   5805
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F1AA
            Key             =   "function"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F746
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FCE2
            Key             =   "var"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3027E
            Key             =   "obj"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tC 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4320
      Top             =   1935
   End
   Begin MSComctlLib.ImageList imlTV 
      Left            =   3735
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3081A
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30BB6
            Key             =   "otherfile"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31152
            Key             =   "open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32FD6
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34E5A
            Key             =   "audio"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36CDE
            Key             =   "program"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37D32
            Key             =   "shellscript"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38D86
            Key             =   "script"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39122
            Key             =   "winword"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A176
            Key             =   "image"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AFCA
            Key             =   "psd"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C01E
            Key             =   "pdf"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D072
            Key             =   "archive"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D40E
            Key             =   "css"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTB 
      Left            =   5040
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D9AA
            Key             =   "asc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F82E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FDCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40AA6
            Key             =   "close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":414DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":421BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42A4E
            Key             =   "symbol"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44576
            Key             =   "html"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":453CA
            Key             =   "new"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4641E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":467BA
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4760E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48462
            Key             =   "find"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":492B6
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A10A
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AF5E
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BDB2
            Key             =   "web"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D77E
            Key             =   "open"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DD1A
            Key             =   "save"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ED6E
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FDC2
            Key             =   "print"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":508FA
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":50E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51432
            Key             =   "time"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":519CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51F6A
            Key             =   "pprev"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5250A
            Key             =   "clr"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52AA6
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53042
            Key             =   "findnext"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53E96
            Key             =   "pageview"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54232
            Key             =   "viewpage"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":547CE
            Key             =   "viewcode"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54D6A
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55106
            Key             =   "lastpos"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5526E
            Key             =   "def"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5580A
            Key             =   "defEx"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Tag             =   "The status bar displays current status and informs you of errors."
      Top             =   8010
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13996
            Picture         =   "frmMain.frx":5596E
            Text            =   "Press F6 for quick help or F1 for contents."
            TextSave        =   "Press F6 for quick help or F1 for contents."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2672
            MinWidth        =   1412
            Text            =   " Line 0, Col 0, Sel 0 "
            TextSave        =   " Line 0, Col 0, Sel 0 "
            Object.ToolTipText     =   "Click to go to a particular line."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1412
            Text            =   " 0 Lines "
            TextSave        =   " 0 Lines "
            Object.ToolTipText     =   "Shows total lines in the document."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   494
            MinWidth        =   494
            Object.ToolTipText     =   "Indicates when WonderHTML is busy."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1766
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTB2 
      Left            =   5805
      Top             =   3060
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55F0A
            Key             =   "strike"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56066
            Key             =   "center"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":561C2
            Key             =   "big"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5631E
            Key             =   "small"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5647A
            Key             =   "left"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":565D6
            Key             =   "bullets"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56732
            Key             =   "numbers"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5688E
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":569EA
            Key             =   "time"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":583B6
            Key             =   "font"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58952
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58AAE
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58C0A
            Key             =   "applet"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59A5E
            Key             =   "right"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59BBA
            Key             =   "link"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59D16
            Key             =   "image"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59E72
            Key             =   "script"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A40E
            Key             =   "indent"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A9B6
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AB12
            Key             =   "purple_doc"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B0AE
            Key             =   "diag_doc"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B3CA
            Key             =   "smileys"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B966
            Key             =   "props"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BAC2
            Key             =   "colors"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C356
            Key             =   "custom"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D1AA
            Key             =   "colour"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D746
            Key             =   "person"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DCE2
            Key             =   "nodes"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlHTML 
      Left            =   5040
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DE3E
            Key             =   "root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F80A
            Key             =   "category"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5FDA6
            Key             =   "Objects"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60DFA
            Key             =   "Plugins"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61E4E
            Key             =   "declare"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":623EA
            Key             =   "images"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6323E
            Key             =   "links"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":637E6
            Key             =   "bookmarks"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":63D8E
            Key             =   "applets"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":64336
            Key             =   "layers"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":64652
            Key             =   "comments"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":64BF6
            Key             =   "forms"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6519E
            Key             =   "styles"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":661F2
            Key             =   "lists"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6634E
            Key             =   "headings"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":668EA
            Key             =   "scripts"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66C86
            Key             =   "divisions"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":67222
            Key             =   "tables"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6737E
            Key             =   "others"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewWhat 
         Caption         =   "&New"
         Begin VB.Menu mnuFileNew 
            Caption         =   "&Blank"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuUsingTemplate 
            Caption         =   "&Using..."
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpenRemote 
         Caption         =   "&Load..."
      End
      Begin VB.Menu mnuFTPSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPublish 
         Caption         =   "&Publish..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrinterSetup 
         Caption         =   "Print s&etup..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Enabled         =   0   'False
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu sepbarMRU 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTree 
      Caption         =   "&File list"
      Visible         =   0   'False
      Begin VB.Menu mnuTreeOpen 
         Caption         =   "&Open file"
      End
      Begin VB.Menu mnuOpenWith 
         Caption         =   "Open &with..."
      End
      Begin VB.Menu mnuTreeLinkFile 
         Caption         =   "Create &link"
      End
      Begin VB.Menu asdasd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenFileList 
         Caption         =   "Enlist files..."
      End
      Begin VB.Menu mnuFilelistSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuTreeRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuTreeSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeCopy 
         Caption         =   "&Copy to..."
      End
      Begin VB.Menu mnuTreeMove 
         Caption         =   "&Move to..."
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "File I&nfo..."
      End
   End
   Begin VB.Menu mnuWeb 
      Caption         =   "&Web"
      Begin VB.Menu mnuNewWeb 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuOpenWeb 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuWebSepZ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebRefresh 
         Caption         =   "R&efresh"
      End
      Begin VB.Menu mnuWebSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebDefault 
         Caption         =   "D&efault"
      End
      Begin VB.Menu mnuWebSepX 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebToolsReporter 
         Caption         =   "&Report..."
      End
      Begin VB.Menu mnuWebSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebMRU 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuBatch 
      Caption         =   "&Batch"
      Begin VB.Menu mnuConvert 
         Caption         =   "Con&vert..."
      End
      Begin VB.Menu mnuImp 
         Caption         =   "&Import..."
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFindFIles 
         Caption         =   "Fin&d..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolBar 
         Caption         =   "&Tool Bar"
      End
      Begin VB.Menu mnuTB2 
         Caption         =   "B&uttons"
         Visible         =   0   'False
         Begin VB.Menu mnuAddBtn 
            Caption         =   "&Add Button..."
         End
         Begin VB.Menu mnuRemBtn 
            Caption         =   "&Remove..."
         End
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "&Status Bar"
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFileTree 
         Caption         =   "File &Tree"
      End
      Begin VB.Menu mnuViewDocuments 
         Caption         =   "&Document"
      End
      Begin VB.Menu mnuViewScripts 
         Caption         =   "&ScriptView"
      End
      Begin VB.Menu mnuViewTask 
         Caption         =   "&TaskView"
      End
      Begin VB.Menu mnuPopTask 
         Caption         =   "T&asks"
         Visible         =   0   'False
         Begin VB.Menu mnuEditTask 
            Caption         =   "&Edit task..."
         End
         Begin VB.Menu mnuEditPage 
            Caption         =   "E&dit page..."
         End
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpHomepage 
         Caption         =   "&Homepage"
      End
      Begin VB.Menu mnuRegisterApp 
         Caption         =   "&Registration"
      End
      Begin VB.Menu mnuHSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTodayTip 
         Caption         =   "&Today's tip"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "A&bout..."
      End
   End
End
Attribute VB_Name = "frmMDI"
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

Option Explicit
Public WebMRU As Collection
Public FileMRU As Collection
Public fileList As String

Private Sub cbClrs_Click()
On Error Resume Next
Clipboard.Clear
Dim sClrText As String
If cbClrs.Text <> "Custom..." Then sClrText = cbClrs.Text Else sClrText = InsCustomColors
If sClrText = "" Then Exit Sub
ActiveForm.RTF1.Visible = False
Dim sTag As String, s As String, sWhole As String
Dim pos As Long, pos2 As Long, lOrig As Long
Dim modeLen As Long, selLenDiff As Long
If IsInQuotes(ActiveForm.RTF1.SelStart) Then GoOutsideQuotes (ActiveForm.RTF1.SelStart)
lOrig = ActiveForm.RTF1.SelStart
s = Mid$(ActiveForm.RTF1.Text, 1, ActiveForm.RTF1.SelStart)
pos2 = InStrRev(s, ">")
pos = InStrRev(s, "<")
If pos = 0 Or pos2 = 0 Then GoTo absolute
pos = pos + 1
sTag = Mid$(ActiveForm.RTF1.Text, pos, pos2 - pos)
sWhole = sTag
If InStr(sTag, " ") > 0 Then sTag = Left(sTag, InStr(sTag, " ") - 1)
If LCase(sTag) = "font" Then
  pos = InStr(pos + 1, s, " color=")
  If pos = 0 Then GoTo addattrib
  pos = pos + Len(" color=")
  pos2 = InStr(pos + 1, s, " ")
  modeLen = 2
  pos2 = pos2 + 1
  If Mid$(s, pos, 1) = Chr(34) Then pos2 = InStr(pos + 1, s, Chr(34)): modeLen = 0
  If Mid$(s, pos, 1) = "'" Then pos2 = InStr(pos + 1, s, "'"): modeLen = 0
  If pos2 = 1 Then pos2 = InStrRev(s, ">") + 1
  ActiveForm.RTF1.SelStart = pos - 1
  ActiveForm.RTF1.SelLength = pos2 - pos + 1 - modeLen
  selLenDiff = ActiveForm.RTF1.SelLength - Len(sClrText) + modeLen - 2 '2 since we're adding quotes
  ActiveForm.RTF1.SelText = Chr(34) & sClrText & Chr(34)
Else
  GoTo absolute
End If
ActiveForm.RTF1.SelStart = lOrig - selLenDiff
ActiveForm.RTF1.Visible = True
ActiveForm.RTF1.SetFocus
Exit Sub
absolute:
ActiveForm.RTF1.SelText = "<FONT color=" & Chr(34) & sClrText & Chr(34) & "></FONT>"
ActiveForm.RTF1.SelStart = ActiveForm.RTF1.SelStart - 7
ActiveForm.RTF1.Visible = True
ActiveForm.RTF1.SetFocus
Exit Sub
addattrib:
pos = InStrRev(s, ">")
If pos = 0 Then GoTo skip
lOrig = ActiveForm.RTF1.SelStart
ActiveForm.RTF1.SelStart = pos - 1
ActiveForm.RTF1.SelText = " color=" & Chr(34) & sClrText & Chr(34)
ActiveForm.RTF1.SelStart = lOrig + 9 + Len(cbClrs.Text)
skip:
ActiveForm.RTF1.Visible = True
ActiveForm.RTF1.SetFocus
End Sub

Private Sub cbFonts_Click()
On Error Resume Next
If cbFonts.Text = "" Then Exit Sub
ActiveForm.RTF1.Visible = False
Dim sTag As String, s As String, sWhole As String
Dim pos As Long, pos2 As Long, lOrig As Long
Dim modeLen As Long, selLenDiff As Long
If IsInQuotes(ActiveForm.RTF1.SelStart) Then GoOutsideQuotes (ActiveForm.RTF1.SelStart)
lOrig = ActiveForm.RTF1.SelStart
s = Mid$(ActiveForm.RTF1.Text, 1, ActiveForm.RTF1.SelStart)
pos2 = InStrRev(s, ">")
pos = InStrRev(s, "<")
If pos = 0 Or pos2 = 0 Then GoTo absolute
pos = pos + 1
sTag = Mid$(ActiveForm.RTF1.Text, pos, pos2 - pos)
sWhole = sTag
If InStr(sTag, " ") > 0 Then sTag = Left(sTag, InStr(sTag, " ") - 1)
If LCase(sTag) = "font" Then
  pos = InStr(pos + 1, s, " face=")
  If pos = 0 Then GoTo addattrib
  pos = pos + Len(" face=")
  pos2 = InStr(pos + 1, s, " ")
  modeLen = 2
  pos2 = pos2 + 1
  If Mid$(s, pos, 1) = Chr(34) Then pos2 = InStr(pos + 1, s, Chr(34)): modeLen = 0
  If Mid$(s, pos, 1) = "'" Then pos2 = InStr(pos + 1, s, "'"): modeLen = 0
  If pos2 = 1 Then pos2 = InStrRev(s, ">") + 1
  ActiveForm.RTF1.SelStart = pos - 1
  ActiveForm.RTF1.SelLength = pos2 - pos + 1 - modeLen
  selLenDiff = ActiveForm.RTF1.SelLength - Len(cbFonts.Text) + modeLen - 2 '2 since we're adding quotes
  ActiveForm.RTF1.SelText = Chr(34) & cbFonts.Text & Chr(34)
Else
  GoTo absolute
End If
ActiveForm.RTF1.SelStart = lOrig - selLenDiff
ActiveForm.RTF1.Visible = True
ActiveForm.RTF1.SetFocus
Exit Sub
absolute:
ActiveForm.RTF1.SelText = "<FONT face=" & Chr(34) & cbFonts.Text & Chr(34) & "></FONT>"
ActiveForm.RTF1.SelStart = ActiveForm.RTF1.SelStart - 7
ActiveForm.RTF1.Visible = True
ActiveForm.RTF1.SetFocus
Exit Sub
addattrib:
pos = InStrRev(s, ">")
If pos = 0 Then GoTo skip
lOrig = ActiveForm.RTF1.SelStart
ActiveForm.RTF1.SelStart = pos - 1
ActiveForm.RTF1.SelText = " face=" & Chr(34) & cbFonts.Text & Chr(34)
ActiveForm.RTF1.SelStart = lOrig + 8 + Len(cbFonts.Text)
skip:
ActiveForm.RTF1.Visible = True
ActiveForm.RTF1.SetFocus
End Sub

Private Sub cbSizes_KeyDown(KeyCode As Integer, Shift As Integer)
cbSizes.SetFocus
cbSizes.SelLength = 0
cbSizes.SelStart = Len(cbSizes.Text)
End Sub

Private Sub cbSizes_Click()
On Error Resume Next
If cbSizes.Text = "" Then Exit Sub
ActiveForm.RTF1.Visible = False
Dim sTag As String, s As String, sWhole As String
Dim pos As Long, pos2 As Long, lOrig As Long
Dim modeLen As Long, selLenDiff As Long
If IsInQuotes(ActiveForm.RTF1.SelStart) Then GoOutsideQuotes (ActiveForm.RTF1.SelStart)
lOrig = ActiveForm.RTF1.SelStart
s = Mid$(ActiveForm.RTF1.Text, 1, ActiveForm.RTF1.SelStart)
pos2 = InStrRev(s, ">")
pos = InStrRev(s, "<")
If pos = 0 Or pos2 = 0 Then GoTo absolute
pos = pos + 1
sTag = Mid$(ActiveForm.RTF1.Text, pos, pos2 - pos)
sWhole = sTag
If InStr(sTag, " ") > 0 Then sTag = Left(sTag, InStr(sTag, " ") - 1)
If LCase(sTag) = "font" Then
  pos = InStr(pos + 1, s, " size=")
  If pos = 0 Then GoTo addattrib
  pos = pos + Len(" size=")
  pos2 = InStr(pos + 1, s, " ")
  modeLen = 2
  pos2 = pos2 + 1
  If Mid$(s, pos, 1) = Chr(34) Then pos2 = InStr(pos + 1, s, Chr(34)): modeLen = 0
  If Mid$(s, pos, 1) = "'" Then pos2 = InStr(pos + 1, s, "'"): modeLen = 0
  If pos2 = 1 Then pos2 = InStrRev(s, ">") + 1
  ActiveForm.RTF1.SelStart = pos - 1
  ActiveForm.RTF1.SelLength = pos2 - pos + 1 - modeLen
  selLenDiff = ActiveForm.RTF1.SelLength - Len(cbSizes.Text) + modeLen - 2 '2 since we're adding quotes
  ActiveForm.RTF1.SelText = Chr(34) & cbSizes.Text & Chr(34)
Else
  GoTo absolute
End If
ActiveForm.RTF1.SelStart = lOrig - selLenDiff
ActiveForm.RTF1.Visible = True
ActiveForm.RTF1.SetFocus
Exit Sub
absolute:
ActiveForm.RTF1.SelText = "<FONT size=" & Chr(34) & cbSizes.Text & Chr(34) & "></FONT>"
ActiveForm.RTF1.SelStart = ActiveForm.RTF1.SelStart - 7
ActiveForm.RTF1.Visible = True
ActiveForm.RTF1.SetFocus
Exit Sub
addattrib:
pos = InStrRev(s, ">")
If pos = 0 Then GoTo skip
lOrig = ActiveForm.RTF1.SelStart
ActiveForm.RTF1.SelStart = pos - 1
ActiveForm.RTF1.SelText = " size=" & Chr(34) & cbSizes.Text & Chr(34)
ActiveForm.RTF1.SelStart = lOrig + 8 + Len(cbSizes.Text)
skip:
ActiveForm.RTF1.Visible = True
ActiveForm.RTF1.SetFocus
End Sub

Private Sub Fldr_Change()
On Error Resume Next
Fil.Path = Fldr.Path
End Sub

Private Sub lvTasks_DblClick()
If lvTasks.ListItems.count = 0 Then Exit Sub
tbView_ButtonClick tbView.Buttons(2)
End Sub

Private Sub lvTasks_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
SB.Panels(1).Text = Item.Key & ": " & Chr(34) & Item.ListSubItems(1).Text & Chr(34) & " (" & IIf(Item.SmallIcon = 13, "Comment", "Task") & ")"
End Sub

Private Sub lvTasks_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If lvTasks.ListItems.count > 0 And Button = 2 Then PopupMenu mnuPopTask
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
GetPrefs
LoadToolBar
LoadFonts
SetFont Me
GetFileMRU
GetWebMRU
AddFlags
ChDir App.Path
LoadCMDLine
LoadMenus
ResizeBar
If ReadValue("WebDefault") <> "" Then LoadWeb ReadValue("WebDefault")
AddTemplates
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
mnuView_Click
If Button = 2 And Shift = 0 Then PopupMenu mnuView
If Button = 2 And Shift = 1 Then PopupMenu mnuFile
If Button = 2 And Shift = 2 Then PopupMenu mnuWeb
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If SSTab1.Height <> pLeft.ScaleHeight - 15 Then MDIForm_Resize
End Sub

Sub MDIForm_Resize()
On Error Resume Next
SSTab1.Height = pLeft.ScaleHeight - 15
tvD.Height = SSTab1.Height - 525
tvW.Height = ReadValue("TreeHeight", tvD.Height)
tvS.Height = tvD.Height
lvTasks.Width = tvW.Width
lvTasks.Height = SSTab1.Height - lvTasks.Top - 105
lvTasks.ColumnHeaders(2).Width = lvTasks.Width - 1440 - 300  '300 for scrollbar
ResizeBar
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
SetPrefs
End
End Sub

Private Sub mnuAddBtn_Click()
frmButt.Show vbModal
End Sub

Private Sub mnuConvert_Click()
Load frmConv
frmConv.SSTab1.Tab = 1
frmConv.lbFile.Enabled = False
frmConv.Show vbModal
End Sub

Private Sub mnuEditPage_Click()
tbView_ButtonClick tbView.Buttons(2)
End Sub

Private Sub mnuEditTask_Click()
tbView_ButtonClick tbView.Buttons(1)
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
On Error Resume Next
Dim lpF As New frmChild
Load lpF
lpF.LoadHTMLFile mnuFileMRU(Index).Tag
End Sub

Private Sub mnuFileNew_Click()
NewDocument
End Sub

Private Sub mnuFileOpen_Click()
OpenDocument
End Sub

Private Sub mnuFileProperties_Click()
FileInfo tvW.SelectedItem.Key
End Sub

Private Sub mnuFindFiles_Click()
Load frmFind
With frmFind
.txFindFiles.TabIndex = 0
.SSTab1.Tab = 2
.Label1(1).Enabled = False
.Label1(2).Enabled = False
.Label1(3).Enabled = False
.txF.Enabled = False
.txR.Enabled = False
.chCase.Enabled = False
.chWhole.Enabled = False
.cmFind.Enabled = False
.cmRepAll.Enabled = False
.cmRepThis.Enabled = False
.lbG.Enabled = False
.lbDetPos.Enabled = False
.txG.Enabled = False
.cbAB.Enabled = False
.cbAbsRel.Enabled = False
.chCloseGo.Enabled = False
.cmG.Enabled = False
.opG(0).Enabled = False
.opG(1).Enabled = False
.opG(2).Enabled = False
.Label6.Enabled = False
.SSTab1_Click 0
frmFind.Show vbModal
End With
End Sub

Private Sub mnuGenFileList_Click()
On Error Resume Next
Fldr.Path = PathIsLegal(tvW.SelectedItem.Index)
SB.Panels(4).Picture = imlBusy.ListImages(2).Picture
GenFileList Fldr.Path
SB.Panels(1).Text = "Finished generating file list."
NewDocument
ActiveForm.RTF1.SelText = "<TABLE width=100% border=0 cellspacing=0 cellpadding=1>" & vbNewLine & "<!-- Begin WonderHTML FileList //-->" & vbNewLine & fileList & "</TABLE>" & vbNewLine
ActiveForm.RTF1.SelText = "<!-- End WonderHTML FileList //-->"
SB.Panels(4).Picture = imlBusy.ListImages(1).Picture
fileList = ""
ActiveForm.RTF1.SetFocus
End Sub

Private Sub mnuHelpAbout_Click()
Load frmSplash
frmSplash.tUnload.Interval = 4000
frmSplash.tUnload.Enabled = True
frmSplash.Show vbModal
frmAbout.Show vbModal
End Sub

Private Sub mnuHelpContents_Click()
frmHelp.Show vbModal
'ShellExecute 0, "open", App.Path & "\help\index.html", "", "", 10
End Sub

Private Sub mnuHelpHomepage_Click()
ShellExecute 0, "open", "http://sushantshome.tripod.com/vb/wonder.html", "", "", 10
End Sub

Private Sub mnuImp_Click()
If IsWebOpen Then frmImp.Show vbModal Else MsgBox "There is no open web.", vbExclamation
End Sub

Private Sub mnuNewWeb_Click()
On Error GoTo hell
Dim lpLoc As String
lpLoc = SelectDir()
If lpLoc = "" Then Exit Sub
If IsWebOpen Then CloseWeb
MkDir lpLoc
LoadWeb lpLoc
Exit Sub
hell:
MsgBox Error, vbExclamation
End Sub

Private Sub mnuOpenRemote_Click()
frmGet.Show
End Sub

Private Sub mnuOpenWeb_Click()
On Error GoTo hell
Dim lpLoc As String
lpLoc = SelectDir(True, 3465)
If lpLoc = "" Then Exit Sub
LoadWeb lpLoc
Exit Sub
hell:
MsgBox Error, vbExclamation
End Sub

Private Sub mnuOpenWith_Click()
Load frmOW
frmOW.Caption = tvW.SelectedItem.Key
frmOW.Show vbModal
End Sub

Private Sub mnuPublish_Click()
If tvW.Nodes.count < 2 Then MsgBox "No web is open or there are no files in the web.", vbCritical: Exit Sub
If Dir(FullPath(App.Path, "wpublish.exe")) = "" Then MsgBox "wpublish.exe cannot be found.", vbCritical: Exit Sub
ShellExecute hWnd, "open", FullPath(App.Path, "wpublish.exe"), tvW.Nodes(1).Key, "", 10
End Sub

Private Sub mnuRegisterApp_Click()
frmReg.Show vbModal
End Sub

Private Sub mnuRemBtn_Click()
Load frmButt
frmButt.SSTab1.Tab = 1
frmButt.Show vbModal
End Sub

Private Sub mnuTodayTip_Click()
On Error Resume Next
frmTip.Show vbModal
End Sub

Private Sub mnuTreeCopy_Click()
On Error Resume Next
Dim loc As String

loc = SelectDir(True, 3465)
If loc = "" Then Exit Sub

If Right(loc, 1) <> "\" Then loc = loc & "\"

MousePointer = 11: SB.Panels(4).Picture = imlBusy.ListImages(2).Picture
DoEvents
If CopyFile(tvW.SelectedItem.Key, loc & tvW.SelectedItem.Text) Then
tvW.Nodes.Add Left(loc, Len(loc) - 1), tvwChild, loc & tvW.SelectedItem.Text, tvW.SelectedItem.Text, FileIcon(tvW.SelectedItem.Text)
End If
DoEvents
MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture
End Sub

Private Sub mnuTreeDelete_Click()
On Error Resume Next
MousePointer = 11: SB.Panels(4).Picture = imlBusy.ListImages(2).Picture
DoEvents
If DeleteFile(tvW.SelectedItem.Key) Then tvW.Nodes.Remove tvW.SelectedItem.Index
DoEvents
MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture
End Sub

Private Sub mnuTreeLinkFile_Click()
On Error Resume Next
Dim Path As String
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Caption = "Untitled" Then Path = tvW.SelectedItem.Key: GoTo n

Path = Replace(tvW.SelectedItem.Key, Up1Level(ActiveForm.Caption), "")
Path = Replace(Path, "\", "/")
If Left(Path, 1) = "/" Then Path = Right(Path, Len(Path) - 1)
n: 'next
ActiveForm.RTF1.SelText = "<A href=" & Chr(34) & Path & Chr(34) & ">" & Path & "</A>"
ActiveForm.RTF1.SelStart = ActiveForm.RTF1.SelStart - Len(Path & "</A>")
ActiveForm.RTF1.SelLength = Len(Path)
ActiveForm.RTF1.SetFocus
End Sub

Private Sub mnuTreeMove_Click()
On Error Resume Next
Dim loc As String

loc = SelectDir(True, 3465)
If loc = "" Then Exit Sub

If Right(loc, 1) <> "\" Then loc = loc & "\"

MousePointer = 11: SB.Panels(4).Picture = imlBusy.ListImages(2).Picture
DoEvents
If MoveFile(tvW.SelectedItem.Key, loc & tvW.SelectedItem.Text) Then
tvW.Nodes.Add Left(loc, Len(loc) - 1), tvwChild, loc & tvW.SelectedItem.Text, tvW.SelectedItem.Text, FileIcon(tvW.SelectedItem.Text)
tvW.Nodes.Remove tvW.SelectedItem.Index
End If
DoEvents
MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture


End Sub

Private Sub mnuTreeOpen_Click()
OnNodeClick tvW.SelectedItem
End Sub

Private Sub mnuTreeRename_Click()
tvW.StartLabelEdit
End Sub

Private Sub mnuUsingTemplate_Click()
frmTemp.Show vbModal
End Sub

Private Sub mnuViewDocuments_Click()
mnuViewDocuments.Checked = Not mnuViewDocuments.Checked
SSTab1.TabVisible(1) = mnuViewDocuments.Checked
SaveValue "DocumentTree", mnuViewDocuments.Checked
End Sub

Private Sub mnuViewOptions_Click()
frmOpts.Show vbModal
End Sub

Private Sub mnuViewScripts_Click()
mnuViewScripts.Checked = Not mnuViewScripts.Checked
SSTab1.TabVisible(2) = mnuViewScripts.Checked
SaveValue "ScriptView", mnuViewScripts.Checked
End Sub

Private Sub mnuViewTask_Click()
mnuViewTask.Checked = Not mnuViewTask.Checked
frmMain.SSTab1.TabVisible(3) = mnuViewTask.Checked
SaveValue "TaskView", frmMain.SSTab1.TabVisible(3)
End Sub

Private Sub mnuWeb_Click()
On Error GoTo hell
mnuWebDefault.Enabled = (tvW.Nodes.count > 0)
mnuWebRefresh.Enabled = mnuWebDefault.Enabled
mnuWebClose.Enabled = (tvW.Nodes.count > 0)
mnuWebDefault.Checked = (ReadValue("WebDefault") = frmMain.tvW.Nodes(1).Key)
Exit Sub
hell:
mnuWebDefault.Enabled = False
Resume Next
End Sub

Private Sub mnuWebClose_Click()
CloseWeb
End Sub

Private Sub mnuWebDefault_Click()
mnuWebDefault.Checked = Not mnuWebDefault.Checked
If mnuWebDefault.Checked Then
SaveValue "WebDefault", tvW.Nodes(1).Key 'root
Else
SaveValue "WebDefault", "" 'nothing
End If
End Sub

Private Sub mnuWebMRU_Click(Index As Integer)
LoadWeb mnuWebMRU(Index).Tag
End Sub

Sub mnuWebRefresh_Click()
LoadWeb frmMain.tvW.Nodes(1).Key
End Sub

Private Sub mnuWebToolsReporter_Click()
On Error GoTo hell
ShowReport tvW.Nodes(1).Text
Exit Sub
hell:
MsgBox "There is no open web.", vbExclamation
End Sub

Private Sub pLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
SSTab1_MouseMove 0, 0, 0, 0 'dummies
End Sub

Private Sub pS_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
tC.Enabled = True
End Sub

Private Sub pS_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
tC.Enabled = False
ResizeBar
SaveValue "TreeWidth", pLeft.Width
End Sub

Private Sub SB_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If X < SB.Panels(1).Width Or Button <> 1 Then Exit Sub
If X > SB.Panels(1).Width And X < SB.Panels(1).Width + SB.Panels(2).Width Then SB.Panels(2).Bevel = sbrInset
End Sub

Private Sub SB_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If X < SB.Panels(1).Width Then Exit Sub
If X > SB.Panels(1).Width And X < SB.Panels(1).Width + SB.Panels(2).Width Then SB.Panels(2).Bevel = sbrRaised: PanelClick 2
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
SSTab1.TabCaption(PreviousTab) = " "
SSTab1.Caption = " " & Trim(Mid$(SSTab1.Tag, SSTab1.Tab * 8 + 1, 8))
If (SSTab1.Tab = 1 Or SSTab1.Tab = 2) And FormsLeft = 0 Then MsgBox "There are no documents open.", vbExclamation: SSTab1.Tab = PreviousTab
If SSTab1.Tab = 3 And tvW.Nodes.count = 0 Then MsgBox "There is no web open.", vbExclamation: SSTab1.Tab = PreviousTab
ActiveForm.RTF1.SetFocus
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If SSTab1.Height <> pLeft.ScaleHeight - 15 Then MDIForm_Resize

End Sub

Private Sub PanelClick(PanelIndex As Long)
Select Case PanelIndex
Case 2
If FormsLeft = 0 Then Exit Sub
ActiveForm.GotoLineProc
Case 5
Load frmOpts
frmOpts.SSTab1.Tab = 4
frmOpts.cbSpeeds.TabIndex = 0
frmOpts.Show vbModal
End Select
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lWidth As Long
On Error GoTo h
Select Case Button.Index
Case 1 'new
NewDocument
Case 2 'open
OpenDocument
Case 3 'save
ActiveForm.mnuFileSave_Click
Case 5 'print
ActiveForm.mnuFilePrint_Click
Case 6 'pprev
ActiveForm.mnuFilePrintPreview_Click
Case 8 'close
Unload ActiveForm
Case 10 'web
lWidth = (TB.ButtonWidth * 6) + (3 * TB.Buttons(4).Width) + 3 * 195 + 140 '4 seps, 6 buttons before 11, 195 is width of drop-down as well as coolbar left
TB.Buttons(10).Value = tbrPressed
mnuWeb_Click 'enable/disable
PopupMenu mnuWeb, , lWidth, TB.ButtonHeight + 45, mnuOpenWeb
TB.Buttons(10).Value = tbrUnpressed
Case 12 'test
ActiveForm.mnuPreview_Click
End Select
h:
If Err.Number = 91 Then SB.Panels(1).Text = "No documents are currently open."
End Sub

Private Sub TB_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
If ButtonMenu.parent.Index = 1 Then 'new button
  If ButtonMenu.Index = 1 Then
  mnuFileNew_Click
  Else
  Dim lpF As New frmChild
  Load lpF
  lpF.RTF1.LoadFile ButtonMenu.Key, rtfText
  lpF.Icon = lpF.p1.Picture
  lpF.RTF1.SetFocus
  End If
ElseIf ButtonMenu.parent.Index = 2 Then
  If ButtonMenu.Index > 3 Then 'open button
  mnuFileMRU_Click ButtonMenu.Index - 3
  Else
  If ButtonMenu.Key = "browse" Then TB_ButtonClick TB.Buttons(2) Else mnuOpenRemote_Click
  End If
ElseIf ButtonMenu.parent.Index = 12 Then
  If ButtonMenu.Index = 1 Then frmAddB.Show vbModal
  If ButtonMenu.Index > 2 Then
    If ActiveForm.Caption = "Untitled" Then ActiveForm.mnuFileSave_Click
    If ActiveForm.Caption = "Untitled" Then SB.Panels(1).Text = "The file must be saved in order to preview it.": Exit Sub
    ShellExecute hWnd, "open", ButtonMenu.Tag, ActiveForm.Caption, "", 10
  End If
ElseIf ButtonMenu.parent.Index = 8 Then
  If ButtonMenu.Index = 1 Then Unload ActiveForm
  If ButtonMenu.Index = 2 Then ActiveForm.mnuWindowUnloadAll_Click
End If
End Sub

Sub NewDocument()
On Error Resume Next
Dim lpF As New frmChild
SB.Panels(1).Text = "Loading a new document..."
Load lpF
lpF.Show
lpF.RTF1.Text = HTML 'the constant
lpF.RTF1.SelStart = lpF.GetSelStart  'approximately
lpF.bChanged = False 'document not changed
lpF.RTF1.SetFocus
lpF.mnuUpdate_Click
SB.Panels(1).Text = ""
End Sub

Sub OpenDocument()
Dim lpF As New frmChild
On Error GoTo hell
CD.ShowOpen
Load lpF
lpF.Show
lpF.LoadHTMLFile CD.Filename
hell:
End Sub


Private Sub mnuView_Click()
mnuViewToolBar.Checked = (frmMain.CBR.Bands(1).Visible)
mnuViewStatusBar.Checked = frmMain.SB.Visible
mnuViewFileTree.Checked = frmMain.pLeft.Visible
mnuViewTask.Checked = frmMain.SSTab1.TabVisible(3)
mnuViewDocuments.Checked = frmMain.SSTab1.TabVisible(1)
mnuViewScripts.Checked = frmMain.SSTab1.TabVisible(2)
End Sub

Sub mnuViewFileTree_Click()
mnuViewFileTree.Checked = Not mnuViewFileTree.Checked
pLeft.Visible = mnuViewFileTree.Checked
SaveValue "FileTree", pLeft.Visible
End Sub

Private Sub mnuViewStatusBar_Click()
mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
frmMain.SB.Visible = mnuViewStatusBar.Checked
SaveValue "Statusbar", SB.Visible
End Sub

Private Sub mnuViewToolBar_Click()
mnuViewToolBar.Checked = Not mnuViewToolBar.Checked
frmMain.CBR.Bands(1).Visible = mnuViewToolBar.Checked
frmMain.CBR.Bands(2).Visible = frmMain.CBR.Bands(1).Visible
frmMain.CBR.Bands(3).Visible = frmMain.CBR.Bands(2).Visible
SaveValue "Toolbar", frmMain.CBR.Bands(1).Visible
End Sub

Sub GetPrefs()
On Error Resume Next
WindowState = ReadValue("WindowState")
If WindowState = 0 Then
Width = ReadValue("Width")
Height = ReadValue("Height")
Top = ReadValue("Top")
Left = ReadValue("Left")
End If
pLeft.Width = ReadValue("TreeWidth")
pLeft.Visible = ReadValue("FileTree", True)
TB.Style = ReadValue("FlatBar")
TB2.Style = TB.Style
tbView.Style = TB.Style
tbEdit.Style = TB.Style
tbMore.Style = TB.Style
CBR.Bands(4).Visible = ReadValue("HTMLBar", True)
SSTab1.TabVisible(1) = ReadValue("DocumentTree", True)
SSTab1.TabVisible(2) = ReadValue("ScriptView", True)
SSTab1.TabVisible(3) = ReadValue("TaskView", True)
CBR.Bands(1).Visible = ReadValue("Toolbar", True)
CBR.Bands(2).Visible = CBR.Bands(1).Visible
CBR.Bands(3).Visible = CBR.Bands(2).Visible
SB.Visible = ReadValue("Statusbar", True)
' setup print header and footer
sPrintHeader = ReadValue("Header", "", "Print")
sPrintFooter = ReadValue("Footer", "", "Print")
gLeftMargin = ReadValue("gLeftMargin", 25, "Print")
gRightMargin = ReadValue("gRightMargin", 25, "Print")
gTopMargin = ReadValue("gTopMargin", 25, "Print")
gBottomMargin = ReadValue("gBottomMargin", 25, "Print")
LoadBrowserList
ResizeBar
End Sub

Sub SetPrefs()
On Error Resume Next
SaveValue "WindowState", WindowState
If WindowState = 0 Then
SaveValue "Width", Width
SaveValue "Height", Height
SaveValue "Top", Top
SaveValue "Left", Left
End If
SaveValue "Header", sPrintHeader, "Print"
SaveValue "Footer", sPrintFooter, "Print"
SaveValue "gLeftMargin", CStr(gLeftMargin), "Print"
SaveValue "gRightMargin", CStr(gRightMargin), "Print"
SaveValue "gTopMargin", CStr(gTopMargin), "Print"
SaveValue "gBottomMargin", CStr(gBottomMargin), "Print"
End Sub

Sub LoadWeb(lpFilePath As String)
On Error Resume Next

If Dir(lpFilePath, vbDirectory) = "" Then
MsgBox lpFilePath & vbCrLf & "The above path cannot be accessed." & vbCrLf & "It may have been moved or removed.", vbExclamation, "Error 54: Path not found"
Exit Sub
End If

lpFilePath = InitCap(lpFilePath)
ChDir lpFilePath
If Right(lpFilePath, 1) = "\" Then lpFilePath = Left(lpFilePath, Len(lpFilePath) - 1)

SB.Panels(1).Text = "Loading " & lpFilePath & " web..."
MousePointer = 11: SB.Panels(4).Picture = imlBusy.ListImages(2).Picture

Dim i As Long

Fldr.Path = lpFilePath
Fldr.Refresh: Fil.Refresh

tvW.Nodes.Clear
tvW.Nodes.Add , , Fldr.Path, InitCap(Fldr.Path), "closed"
tvW.Nodes(1).ExpandedImage = "open"


For i = 0 To Fldr.ListCount - 1
    tvW.Nodes.Add Fldr.Path, tvwChild, Fldr.list(i), GetFile(Fldr.list(i)), "closed"
    tvW.Nodes.Item(Fldr.list(i)).ExpandedImage = "open"
    tvW.Nodes.Add Fldr.list(i), tvwChild, "", ""
Next i

Fil.Path = lpFilePath

For i = 0 To Fil.ListCount - 1
    If LCase(Fil.list(i)) = "files.inf" Then GoTo n
    If LCase(Ext(Fil.list(i))) = "wbackup" Then LoadBAKFile Fil.list(i): GoTo n
    tvW.Nodes.Add IIf(Len(lpFilePath) <= 2, lpFilePath & "\", lpFilePath), tvwChild, FullPath(Fil.Path, Fil.list(i)), Fil.list(i), FileIcon(Fil.list(i))
    'IIF needed for drives
n:
Next i

SB.Panels(1).Text = ""

AddWebMRU lpFilePath

tvW.Nodes(1).Expanded = True

MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture

ListTasks

Caption = tvW.Nodes(1).Text & " - WonderHTML"


Dim s As String, ss() As String
s = ReadValue("Highlight", "", "Highlight", FullPath(tvW.Nodes(1).Key, "files.inf"))
ss = Split(s, "|")
For i = 0 To UBound(ss)
tvW.Nodes(ss(i)).Bold = True
Next i

End Sub

Private Sub OnNodeClick(ByVal Node As MSComctlLib.Node)

On Error Resume Next
Dim lpF As New frmChild, lpD As Form

MousePointer = 11: SB.Panels(4).Picture = imlBusy.ListImages(2).Picture

SB.Panels(1).Text = InitCap(Node.Key)

Select Case Node.Image

Case 5, "closed"
    If Node.Children > 1 And Node.Child.Text <> "" Then MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture: Exit Sub
    If Node.Child.Text = "" Then tvW.Nodes.Remove Node.Child.Index
    Dim i As Long
    Fldr.Path = Node.Key
    If Fldr.ListCount = 0 Then GoTo nfil
nfld:
    For i = 0 To Fldr.ListCount - 1
    tvW.Nodes.Add Node.Key, tvwChild, Fldr.list(i), GetFile(Fldr.list(i)), "closed"
    tvW.Nodes.Item(Fldr.list(i)).ExpandedImage = "open"
    tvW.Nodes.Add Fldr.list(i), tvwChild, "", ""
    Next i
nfil:
    For i = 0 To Fil.ListCount - 1
    If LCase(Fil.list(i)) = "files.inf" Then GoTo n
    tvW.Nodes.Add Node.Key, tvwChild, FullPath(Fil.Path, Fil.list(i)), Fil.list(i), FileIcon(Fil.list(i))
n:
    Next i

Dim s As String, ss() As String
s = ReadValue("Highlight", "", "Highlight", FullPath(tvW.Nodes(1).Key, "files.inf"))
ss = Split(s, "|")
For i = 0 To UBound(ss)
tvW.Nodes(ss(i)).Bold = True
Next i
MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture
    Exit Sub

Case 4, "open"
    MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture
    Exit Sub
    
Case Else 'files and images

'loop and find if it's already open if yes then set focus to it
    For Each lpD In Forms
            If lpD.Caption = Node.Key Then
                lpD.SetFocus: lpD.RTF1.SetFocus: lpD.PB.SetFocus: MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture: Exit Sub
            End If
    Next lpD
        
        
        Select Case LCase(Ext(Node.Key))
        
        'open based on the extension
        Case "html", "css", "txt", "asp", "htm", "js", "vbs", "xml", "htm_", "htt"
        
        Load lpF
        lpF.LoadHTMLFile Node.Key
        
        Case "jpg", "gif", "bmp", "ico"
        
        'below line confirms if user wants the internal image viewer
        If ReadValue("ImageViewer", 1) = 0 Then GoTo nope
        LoadImage Node.Key
        
        Case Else
        
nope:
        Dim sApp As String, thisFile As String * 260, llen As Long
        sApp = ReadValue(Ext(Node.Key), "", "Editors", FullPath(App.Path, "editors.inf"))
        sApp = Mid$(sApp, InStr(sApp, "|") + 1)
        llen = GetShortPathName(Node.Key, thisFile, 260)
        If sApp <> "" Then
          If ShellExecute(hWnd, "open", sApp, thisFile, "", 10) >= 32 Then GoTo ending Else GoTo cantopen
        End If

        'try to execute otherwise notify that file couldn't be opened
        If ShellExecute(Me.hWnd, "open", Node.Key, "", Up1Level(Node.Key), 10) < 32 Then
cantopen:
          MsgBox "Failed to execute " & GetFile(Node.Key), vbExclamation
        End If
        End Select

End Select
ending:
MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture
End Sub

Sub AddFlags()
'add flags to common dialog
CD.Flags = cdlOFNCreatePrompt + cdlOFNFileMustExist + cdlOFNOverwritePrompt + cdlOFNPathMustExist
End Sub

Function IsWebOpen() As Boolean
'is any web open in the manager
IsWebOpen = (tvW.Nodes.count > 0)
End Function

Function CloseWeb()
'does just what it says
mnuWebDefault.Checked = False
tvW.Nodes.Clear
lvTasks.ListItems.Clear
tbView.Buttons(1).Enabled = False
tbView.Buttons(2).Enabled = False
Caption = "WonderHTML"
End Function

Private Sub FindNodeText()
'this is called when user clicks on the document outline tree
On Error Resume Next
If tvD.SelectedItem.Image = "root" Or tvD.SelectedItem.Image = "category" Then Exit Sub
Dim lF As Long
lF = InStr(1, ActiveForm.RTF1.Text, SB.Panels(1).Text)
If lF = 0 Then Exit Sub
ActiveForm.bUpdateFlag = False 'don't update now
ActiveForm.RTF1.SelStart = lF - 1
If ReadValue("SelectFind", , "Documents") = True Then ActiveForm.RTF1.SelLength = Len(SB.Panels(1).Text) Else ActiveForm.RTF1.SetFocus
ActiveForm.RTF1.SetFocus
ActiveForm.bUpdateFlag = True
End Sub

Private Sub TB2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If IsInQuotes(ActiveForm.RTF1.SelStart) Then GoOutsideQuotes (ActiveForm.RTF1.SelStart)
If Button.Tag = "!" Then Button.Tag = Format(Now, "ddd, dd mmm yyyy at hh:mm AMPM."):   Button.Description = Len(Button.Tag)
Button.Tag = Replace(Button.Tag, "|", vbNewLine)
ActiveForm.RTF1.SelText = Button.Tag
ActiveForm.RTF1.SelStart = ActiveForm.RTF1.SelStart - Len(Button.Tag) + CLng(Button.Description)
ActiveForm.SetFocus: ActiveForm.RTF1.SetFocus
End Sub

Private Sub TB2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then PopupMenu mnuTB2
End Sub

Private Sub tbEdit_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key
Case "undo"
ActiveForm.mnuEditUndo_Click
Case "redo"
ActiveForm.mnuEditRedo_Click
Case "cut"
ActiveForm.mnuEditCut_Click
Case "copy"
ActiveForm.mnuEditCopy_Click
Case "paste"
ActiveForm.mnuEditPaste_Click
Case "edit"
ActiveForm.mnuDocProps_Click
Case "symbol"
ActiveForm.mnuInsertSymbol_Click
Case "date"
ActiveForm.mnuInsertDateTime_Click
Case "cpick"
ActiveForm.mnuClrPicker_Click
Case "cfade"
ActiveForm.mnuColourFade_Click
Case "opensrc"
ActiveForm.mnuOpenThisFile_Click
ActiveForm.RTF1.SetFocus
End Select
End Sub

Private Sub tbMore_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
ActiveForm.mnuGotoLastPos_Click
Button.Enabled = (ActiveForm.Positions.count > 0)
Case 2
ActiveForm.mnuEditDefinition_Click
Case 4
ActiveForm.mnuEditFind_Click
Case 5
ActiveForm.mnuFindNext_Click
Case 7
ActiveForm.mnuDocumentConvert_Click
End Select
End Sub

Private Sub tbView_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Index = 1 Then FileInfo lvTasks.SelectedItem.Key, True
If Button.Index = 2 Then
Dim l As New frmChild
Load l
l.LoadHTMLFile lvTasks.SelectedItem.Key
End If
If Button.Index = 4 Then ListTasks
If Button.Index = 5 Then mnuWebToolsReporter_Click
End Sub

Private Sub tC_Timer()
'resize left bar
Dim lpP As POINTAPI
GetCursorPos lpP
pLeft.Width = lpP.X * Screen.TwipsPerPixelX + pS.Width
End Sub

Private Sub tvD_Collapse(ByVal Node As MSComctlLib.Node)
tvD.SelectedItem = Node
End Sub

Private Sub tvD_DblClick()
FindNodeText 'find the text
End Sub

Private Sub tvD_Expand(ByVal Node As MSComctlLib.Node)
tvD.SelectedItem = Node
End Sub

Private Sub tvD_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If ActiveForm Is Nothing Then Exit Sub
If Button = 2 Then
ActiveForm.mnuSelectBody.Enabled = False
ActiveForm.mnuDeleteBody.Enabled = False
PopupMenu ActiveForm.mnuWhatever
End If
End Sub

Private Sub tvD_NodeClick(ByVal Node As MSComctlLib.Node)
SB.Panels(1).Text = Node.Tag
End Sub

Private Sub tvS_Collapse(ByVal Node As MSComctlLib.Node)
tvS.SelectedItem = Node
End Sub

Private Sub tvS_DblClick()
On Error Resume Next
If tvS.SelectedItem.Image = "folder" Then Exit Sub
If tvS.SelectedItem.parent.Key = "server" Then Exit Sub

tvS.SelectedItem.Expanded = Not tvS.SelectedItem.Expanded
ActiveForm.bUpdateFlag = False
If tvS.SelectedItem.Image = "var" And tvS.SelectedItem.parent.Key <> "globals" Then
  Dim ss() As String
  ss = Split(tvS.SelectedItem.Key, ": local to ")
  ss(0) = Trim(ss(0))
  ss(1) = Trim(ss(1))
  ActiveForm.RTF1.SelStart = InStr(1, ActiveForm.RTF1.Text, ss(1)) - 1
  ActiveForm.RTF1.SelStart = InStr(ActiveForm.RTF1.SelStart + 1, ActiveForm.RTF1.Text, ss(0)) - 1
  If ReadValue("SelectFind", , "Documents") = True Then ActiveForm.RTF1.SelLength = Len(ss(1))
  ActiveForm.RTF1.SetFocus
Exit Sub
End If
ActiveForm.RTF1.SelStart = InStr(1, ActiveForm.RTF1.Text, tvS.SelectedItem.Key) - 1
If ReadValue("SelectFind", , "Documents") = True Then ActiveForm.RTF1.SelLength = Len(tvS.SelectedItem.Key)
ActiveForm.RTF1.SetFocus
ActiveForm.bUpdateFlag = True
End Sub

Private Sub tvS_Expand(ByVal Node As MSComctlLib.Node)
tvS.SelectedItem = Node
End Sub

Private Sub tvS_NodeClick(ByVal Node As MSComctlLib.Node)
SB.Panels(1).Text = Node.Key
End Sub

Private Sub tvW_AfterLabelEdit(Cancel As Integer, NewString As String)
'If tvW.SelectedItem.Image = "closed" Or tvW.SelectedItem.Image = "open" Then NewString = tvW.SelectedItem.Text: Exit Sub
If Cancel > 0 Then Exit Sub
If NewString = GetFile(tvW.SelectedItem.Key) Then Exit Sub

Dim StrPS As String

If Right(tvW.SelectedItem.Key, 1) = "\" Then StrPS = "" Else StrPS = "\"

MousePointer = 11: SB.Panels(4).Picture = imlBusy.ListImages(2).Picture
DoEvents
If Not RenameFile(tvW.SelectedItem.Key, tvW.SelectedItem.parent.Key & StrPS & NewString) Then NewString = tvW.SelectedItem.Text
tvW.SelectedItem.Key = tvW.SelectedItem.parent.Key & StrPS & NewString
DoEvents
MousePointer = 0: SB.Panels(4).Picture = imlBusy.ListImages(1).Picture
End Sub

Private Sub tvW_BeforeLabelEdit(Cancel As Integer)
If tvW.SelectedItem.Index = 1 Then Cancel = True: Exit Sub
End Sub

Private Sub tvW_Collapse(ByVal Node As MSComctlLib.Node)
tvW.SelectedItem = Node
End Sub

Private Sub tvW_DblClick()
On Error Resume Next
OnNodeClick tvW.SelectedItem
End Sub

Private Sub tvW_Expand(ByVal Node As MSComctlLib.Node)
tvW.SelectedItem = Node
OnNodeClick Node
End Sub

Private Sub tvW_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbAltMask And KeyCode = vbKeyReturn Then mnuFileProperties_Click
If KeyCode = vbKeyReturn And Shift = 0 Then mnuTreeOpen_Click
If KeyCode = vbKeyC And Shift = vbCtrlMask Then mnuTreeCopy_Click
If KeyCode = vbKeyX And Shift = vbCtrlMask Then mnuTreeMove_Click
If KeyCode = vbKeyF2 Then mnuTreeRename_Click
If KeyCode = vbKeyDelete Then mnuTreeDelete_Click
If Shift = vbShiftMask And KeyCode = vbKeyReturn Then mnuOpenWith_Click
If KeyCode >= vbKey1 And KeyCode <= vbKey6 And Shift = vbCtrlMask Then mnuFileMRU_Click KeyCode - vbKey1 + 1
End Sub

Private Sub tvW_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim rc As Rect
If Button = 2 And tvW.Nodes.count > 0 Then
mnuTreeOpen.Enabled = (tvW.SelectedItem.Image <> "closed")
mnuOpenWith.Enabled = mnuTreeOpen.Enabled
mnuTreeRename.Enabled = tvW.SelectedItem.Index <> 1
mnuTreeLinkFile.Enabled = Not (ActiveForm Is Nothing)
PopupMenu mnuTree
End If
End Sub

Private Sub tvW_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If SSTab1.Height <> pLeft.ScaleHeight - 15 Then MDIForm_Resize
End Sub

Sub ResizeBar()
On Error Resume Next
SSTab1.Width = pLeft.ScaleWidth - 45
tvD.Width = SSTab1.Width - 190
tvW.Width = tvD.Width
tvS.Width = tvD.Width
lvTasks.Width = tvS.Width
lvTasks.ColumnHeaders(2).Width = lvTasks.Width - 1440 - IIf(lvTasks.ListItems.count * 240 >= lvTasks.Height, 300, 0) '300 for scrollbar
pS.Move pLeft.ScaleWidth - 45, 0, 45, pLeft.ScaleHeight
End Sub

Private Sub tvW_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
If Node.Image = "closed" Then
SB.Panels(1).Text = "Double-click to view items in '" & InitCap(Node.Key) & "'."
Else
SB.Panels(1).Text = Round(FileLen(Node.Key) / 1024, 1) & " KB " & FileType(Node.Key) & ": " & InitCap(Node.Key)
GetSpeedInfo FileLen(Node.Key)
End If
End Sub

Sub LoadCMDLine()
If Command$() <> "" Then
  If Left(Command$, 2) = "-f" Then
    LoadWeb Mid$(Command$, 3) 'from 3rd char
  Else
    Dim lpF As New frmChild
    Load lpF
    lpF.RTF1.LoadFile Command$, rtfText
    lpF.Caption = Command$
    lpF.bChanged = False
  End If
End If
End Sub

Function IsInQuotes(SelStart As Long) As Boolean
On Error Resume Next
Dim posLT As Long, posGT As Long
posLT = InStr(SelStart, ActiveForm.RTF1.Text, "<")
If posLT = 0 Then IsInQuotes = True: Exit Function
posGT = InStr(SelStart, ActiveForm.RTF1.Text, ">")
If posGT = 0 Then IsInQuotes = True: Exit Function
IsInQuotes = (posLT > posGT)
End Function

Sub GoOutsideQuotes(SelStart As Long)
On Error Resume Next
Dim iPos As Long
iPos = InStr(SelStart, ActiveForm.RTF1.Text, ">")
If iPos = 0 Then ActiveForm.RTF1.SelStart = Len(ActiveForm.RTF1.Text): Exit Sub
ActiveForm.RTF1.SelStart = iPos
End Sub

Private Sub tvS_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.mnuSelectBody.Enabled = (tvS.SelectedItem.Image <> "folder") And (tvS.SelectedItem.Image <> "obj")
ActiveForm.mnuDeleteBody.Enabled = ActiveForm.mnuSelectBody.Enabled
If Button = 2 Then PopupMenu ActiveForm.mnuWhatever
End Sub

Sub ListTasks()
'supposed to create a list of files which have pending tasks
On Error GoTo hell
Dim thisline As String, s As String, i As Long

lvTasks.ListItems.Clear

Open FullPath(frmMain.CurrentWeb, "files.inf") For Input As #1
  Do Until EOF(1)
      Line Input #1, thisline
      If thisline = "" Or Left(thisline, 1) <> "[" Or Right(thisline, 1) <> "]" Then GoTo nxt
      thisline = Mid$(thisline, 2, Len(thisline) - 2)
      s = Replace(thisline, CurrentWeb, "", , , vbTextCompare)
      s = FirstSlash(s)
      s = Replace(s, "\", "/")
      If ReadValue("Task", "", thisline, FullPath(CurrentWeb, "files.inf")) <> "" Then lvTasks.ListItems.Add , thisline, s, , 1
nxt:
  Loop
Close #1

For i = 1 To lvTasks.ListItems.count
s = ReadValue("Task", "", lvTasks.ListItems(i).Key, FullPath(CurrentWeb, "files.inf"))
lvTasks.ListItems(i).ListSubItems.Add , , s, , s
n:
Next i

If lvTasks.ListItems.count > 0 Then
    lvTasks.SelectedItem = lvTasks.ListItems(1)
    tbView.Buttons(1).Enabled = True
    tbView.Buttons(2).Enabled = True
End If
Exit Sub
hell:
End Sub

Function CurrentWeb() As String
If IsWebOpen Then CurrentWeb = tvW.Nodes(1).Key Else CurrentWeb = App.Path
End Function

Function PathIsLegal(NodeIndex As Integer) As String
If tvW.Nodes(NodeIndex).Image = "open" Or tvW.Nodes(NodeIndex).Image = "closed" Or tvW.Nodes(NodeIndex).Image = "closed" Then PathIsLegal = tvW.Nodes(NodeIndex).Key Else PathIsLegal = tvW.Nodes(NodeIndex).parent.Key
End Function

Sub LoadMenus()
On Error Resume Next
mnuTreeCopy.Caption = mnuTreeCopy.Caption & vbTab & "Ctrl+C"
mnuTreeRename.Caption = mnuTreeRename.Caption & vbTab & "F2"
mnuTreeMove.Caption = mnuTreeMove.Caption & vbTab & "Ctrl+X"
mnuTreeDelete.Caption = mnuTreeDelete.Caption & vbTab & "Del"
mnuTreeOpen.Caption = mnuTreeOpen.Caption & vbTab & "Enter"
mnuFileProperties.Caption = mnuFileProperties.Caption & vbTab & "Alt+Enter"
mnuFileExit.Caption = mnuFileExit.Caption & vbTab & "Alt+F4"
mnuOpenWith.Caption = mnuOpenWith.Caption & vbTab & "Shift+Enter"
TB.Buttons(1).ButtonMenus(1).Text = TB.Buttons(1).ButtonMenus(1).Text & vbTab & "Ctrl+N"
TB.Buttons(2).ButtonMenus(1).Text = TB.Buttons(2).ButtonMenus(1).Text & vbTab & "Ctrl+O"
End Sub

Sub LoadToolBar()
On Error Resume Next
Dim i As Long, tot As Long
tot = ReadValue("BtnCount", 19, "Buttons")
If tot <= 19 Then Exit Sub
For i = 20 To tot
TB2.Buttons.Add i, , , , ReadValue("Btn" & i & "Img", 0, "Buttons")
TB2.Buttons(i).Description = ReadValue("Btn" & i & "Sel", 0, "Buttons")
TB2.Buttons(i).Tag = ReadValue("Btn" & i, "", "Buttons")
TB2.Buttons(i).ToolTipText = TB2.Buttons(i).Tag
TB2.Buttons(i).Enabled = False
If TB2.Buttons(i).Tag = "" Then TB2.Buttons(i).Visible = False
Next i
End Sub

Sub SelectFunBody()
On Error Resume Next
ActiveForm.bUpdateFlag = False
Dim lpText As String, i As Long
lpText = ActiveForm.RTF1.Text
Dim InPos1 As Long, InPosWhole As Long
InPos1 = InStr(1, lpText, tvS.SelectedItem.Key, vbBinaryCompare)
InPosWhole = InStr(InPos1 + 1, lpText, "}")
If StrCount(Mid$(lpText, InPos1, InPosWhole - InPos1), "{") = 0 Then GoTo n
  For i = 1 To StrCount(Mid$(lpText, InPos1, InPosWhole - InPos1), "{")
    InPosWhole = InStr(InPosWhole + 1, lpText, "}")
  Next i
n:
ActiveForm.RTF1.SelStart = InPos1 - 1
ActiveForm.RTF1.SelLength = InPosWhole - InPos1 + 1
ActiveForm.RTF1.SetFocus
ActiveForm.bUpdateFlag = True
End Sub

Sub LoadBAKFile(File As String)
Dim s As VbMsgBoxResult
s = MsgBox(NoExt(File) & " could not be saved because of an interruption." & vbCrLf & "Do you want to recover the file to what it was backed up?" & vbCrLf & vbCrLf & "You can save this to overwrite the unsaved version of the file.", vbExclamation + vbYesNo, NoExt(FullPath(Fil.Path, File)))
If s = vbYes Then
Dim lpF As New frmChild
Load lpF
ActiveForm.RTF1.LoadFile FullPath(Fil.Path, File)
ActiveForm.Caption = NoExt(ActiveForm.RTF1.Filename)
Else
Kill FullPath(Fil.Path, File)
End If
End Sub

Sub LoadToolbarMRU()
Dim i As Integer
For i = 1 To 6
TB.Buttons(2).ButtonMenus(i + 3).Text = mnuFileMRU(i).Caption
TB.Buttons(2).ButtonMenus(i + 3).Tag = mnuFileMRU(i).Tag
If TB.Buttons(2).ButtonMenus(i + 3).Tag <> "" Then TB.Buttons(2).ButtonMenus(i + 3).Visible = True
Next i
End Sub

Sub AddTemplates()
Dim s As String, i As Integer
s = Fil.Path
Fil.Path = FullPath(App.Path, "Templates")
For i = 0 To Fil.ListCount - 1
TB.Buttons(1).ButtonMenus.Add , FullPath(Fil.Path, Fil.list(i)), NoExt(Fil.list(i))
Next i
Fil.Path = s
End Sub

Sub LoadBrowserList()
On Error Resume Next
Dim i As Integer
For i = 3 To TB.Buttons(12).ButtonMenus.count
TB.Buttons(12).ButtonMenus.Remove 3
Next i
Dim s As Integer
s = ReadValue("count", , "Browsers", FullPath(App.Path, "editors.inf"))
For i = 1 To s
TB.Buttons(12).ButtonMenus.Add , "tmp", ReadValue("item" & i, , "Browsers", FullPath(App.Path, "editors.inf"))
TB.Buttons(12).ButtonMenus("tmp").Tag = TB.Buttons(12).ButtonMenus("tmp").Text
TB.Buttons(12).ButtonMenus("tmp").Text = GetFile(TB.Buttons(12).ButtonMenus("tmp").Tag)
TB.Buttons(12).ButtonMenus("tmp").Text = NoExt(TB.Buttons(12).ButtonMenus("tmp").Text)
If TB.Buttons(12).ButtonMenus("tmp").Tag = "" Then TB.Buttons(12).ButtonMenus.Remove "tmp"
TB.Buttons(12).ButtonMenus("tmp").Key = ""
Next i
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode >= vbKey1 And KeyCode <= vbKey6 And Shift = vbCtrlMask Then mnuFileMRU_Click KeyCode - vbKey1 + 1
End Sub

Private Sub lvTasks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode >= vbKey1 And KeyCode <= vbKey6 And Shift = vbCtrlMask Then mnuFileMRU_Click KeyCode - vbKey1 + 1
End Sub

Sub LoadFonts()
On Error Resume Next
Dim i  As Integer
For i = 0 To Screen.FontCount - 1
cbFonts.AddItem Screen.Fonts(i)
Next i
End Sub

Function InsCustomColors() As String
Dim s As String
'Clipboard.Clear
Load frmCPick
frmCPick.Command1.Caption = "Insert"
frmCPick.Show vbModal
s = Clipboard.GetText()
InsCustomColors = s
End Function
