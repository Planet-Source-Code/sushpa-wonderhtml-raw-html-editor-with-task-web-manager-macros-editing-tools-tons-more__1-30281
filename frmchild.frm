VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChild 
   AutoRedraw      =   -1  'True
   Caption         =   "Untitled"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmchild.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   7845
   Tag             =   "This is the Editor."
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RTF1 
      Height          =   3885
      Left            =   1305
      TabIndex        =   1
      Tag             =   $"frmchild.frx":058A
      Top             =   405
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   6853
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmchild.frx":0615
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser IE1 
      Height          =   1500
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   2760
      ExtentX         =   4868
      ExtentY         =   2646
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer tmrAutoSave 
      Interval        =   60000
      Left            =   1710
      Top             =   4410
   End
   Begin VB.PictureBox p2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   990
      Picture         =   "frmchild.frx":070F
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   180
      Picture         =   "frmchild.frx":2581
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   4950
      Visible         =   0   'False
      Width           =   480
   End
   Begin RichTextLib.RichTextBox rtfTmp 
      Height          =   510
      Left            =   1080
      TabIndex        =   0
      Top             =   4275
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   900
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmchild.frx":35C3
   End
   Begin VB.Menu mnuWhatever 
      Caption         =   "&What"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectBody 
         Caption         =   "&Select"
      End
      Begin VB.Menu mnuDeleteBody 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuWhatSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Update"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewWhat 
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
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "S&ave as..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Sa&ve all"
      End
      Begin VB.Menu mnuFileSepBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRevert 
         Caption         =   "Re&vert..."
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintX 
         Caption         =   "&Print"
         Begin VB.Menu mnuFilePrinterSetup 
            Caption         =   "Set&up..."
         End
         Begin VB.Menu mnuFilePrintPreview 
            Caption         =   "Pr&eview..."
         End
         Begin VB.Menu mnuPSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPrintPage 
            Caption         =   "Page..."
         End
         Begin VB.Menu mnuFilePrint 
            Caption         =   "HTML..."
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu mnuSepMRU 
         Caption         =   "-"
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
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "E&dit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "U&ndo"
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuSpecialPaste 
         Caption         =   "Sp&ecial..."
      End
      Begin VB.Menu mnuPasteSpecialOrdered 
         Caption         =   "Or&dered list"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPasteSpecialUnordered 
         Caption         =   "&Unordered list"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPasteSpecialPRE 
         Caption         =   "P&reformatted"
         Visible         =   0   'False
      End
      Begin VB.Menu mnupasteSpecialBRasP 
         Caption         =   "Breaks as &Paras"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPasteSpecialTagsEntities 
         Caption         =   "Tags as &Entities"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &all"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDocHTML 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCodePage 
         Caption         =   "View &Page"
      End
      Begin VB.Menu mnuEditS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenThisFile 
         Caption         =   "Open &file"
      End
      Begin VB.Menu mnuSepBarSome 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDefinition 
         Caption         =   "De&finition"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGotoLastPos 
         Caption         =   "&Last Position"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuSearchMenu 
      Caption         =   "&Search"
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu asdasd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "R&eplace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditGoTo 
         Caption         =   "&Go to..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindFiles 
         Caption         =   "&Find files..."
      End
   End
   Begin VB.Menu mnuInsertMain 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertSymbol 
         Caption         =   "&Symbol..."
      End
      Begin VB.Menu mnuInsertDateTime 
         Caption         =   "Da&te/Time..."
      End
      Begin VB.Menu mnuInseSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEntitize 
         Caption         =   "E&ntitize text"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuColourFade 
         Caption         =   "Fade&d Text..."
      End
      Begin VB.Menu mnuInsertNetscapeFix 
         Caption         =   "NS &Resize fix..."
      End
      Begin VB.Menu somesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDummy 
         Caption         =   "&Other..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuDocument 
      Caption         =   "D&ocument"
      Begin VB.Menu mnuCheckLinks 
         Caption         =   "C&heck Links"
      End
      Begin VB.Menu mnuValidDoc 
         Caption         =   "&Validate code"
      End
      Begin VB.Menu mnuSepValidate 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDocumentConvert 
         Caption         =   "&Convert..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuDocumentSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClrPicker 
         Caption         =   "Pick &Colour..."
      End
      Begin VB.Menu asdasdasd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStats 
         Caption         =   "Sta&tistics..."
      End
      Begin VB.Menu mnuStyleEdit 
         Caption         =   "CSS &Editor..."
      End
      Begin VB.Menu mnuDocProps 
         Caption         =   "&Properties..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolBar 
         Caption         =   "&Tool Bar"
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
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViiewMode 
         Caption         =   "&View Mode"
         Begin VB.Menu mnuViewMode 
            Caption         =   "&No wrap"
            Index           =   1
         End
         Begin VB.Menu mnuViewMode 
            Caption         =   "&Word wrap"
            Index           =   2
         End
         Begin VB.Menu mnuViewMode 
            Caption         =   "&Printer DC"
            Index           =   3
         End
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Wi&ndow"
      Begin VB.Menu mnuCascadeWin 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "&Tile Horizontal"
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowSepX 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAllWin 
         Caption         =   "&All windows"
         Begin VB.Menu mnuConvert 
            Caption         =   "&Convert..."
         End
         Begin VB.Menu mnuHighlight 
            Caption         =   "&Highlight"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuWinSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuWindowMinimizeAll 
            Caption         =   "&Minimize"
         End
         Begin VB.Menu mnuWindowMaximizeAll 
            Caption         =   "Ma&ximize"
         End
         Begin VB.Menu mnuRestoreAll 
            Caption         =   "&Restore"
         End
         Begin VB.Menu mnuWinUnloadSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuWindowUnloadAll 
            Caption         =   "&Close"
         End
      End
      Begin VB.Menu mnuSwitchTo 
         Caption         =   "&Switch to"
         WindowList      =   -1  'True
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
      Begin VB.Menu mnuHsep 
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
Attribute VB_Name = "frmChild"
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
Public bChanged As Boolean
Public bUpdateFlag As Boolean
Public UndoStack As New Collection, RedoStack As New Collection
Public trapUndo As Boolean
Public Positions As Collection
Dim ThisPath As String
Dim CurrentWord As String

Private Sub Form_GotFocus()
On Error Resume Next
RTF1_GotFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = Asc("W") And Shift = vbCtrlMask Then KeyCode = 0: mnuViewCodePage_Click
If KeyCode = vbKeyReturn And Shift = vbAltMask Then mnuDocProps_Click: KeyCode = 0
If KeyCode >= vbKey1 And KeyCode <= vbKey6 And Shift = vbCtrlMask Then mnuFileMRU_Click KeyCode - vbKey1 + 1
End Sub

Private Sub Form_Load()
On Error Resume Next
bUpdateFlag = False
ThisPath = App.Path
Set Positions = New Collection
EnableFindDialog
IE1.Navigate "about:blank"
SetFont Me
CopyMRUList
WindowState = ReadValue("ChildState", 0)
mnuViewMode_Click CInt(ReadValue("ViewMode", 0, "Documents")) + 1
RTF1.Font.Name = ReadValue("FontName", "Tahoma", "Documents")
RTF1.Font.Size = ReadValue("FontSize", 10, "Documents")
rtfTmp.Font.Name = RTF1.Font.Name
rtfTmp.Font.Size = RTF1.Font.Size
RTF1.SelIndent = 45 'just a little
bChanged = False
SetMenus
mnuEdit_Click
If FormsLeft <= 1 Then EnableBar
trapUndo = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If bChanged Then
Dim bMsg As VbMsgBoxResult
bMsg = MsgBox("The current document has changed." & vbNewLine & "Do you want to save changes to it?", vbExclamation + vbYesNoCancel + vbSystemModal, Caption)
    If bMsg = vbNo Then
    Kill "Untitled.wbackup"
    Kill Caption & ".wbackup"
    Cancel = False
    ElseIf bMsg = vbYes Then
    If Mid(Caption, 2, 1) <> ":" Then Cancel = True
    mnuFileSave_Click
    Else
    Cancel = True
    RTF1.SetFocus
    End If
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
RTF1.Move 0, 0, ScaleWidth, ScaleHeight
IE1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Terminate()
On Error Resume Next
If FormsLeft = 0 Then
DisableBar
frmMain.SB.Panels(2).Text = " Line 0, Col 0, Sel 0 "
frmMain.SB.Panels(3).Text = "0 Lines"
End If
If Not frmMain.ActiveForm Is Nothing Then
Outline frmMain.ActiveForm, frmMain.tvD, frmMain.ActiveForm.RTF1.Text, Not frmMain.SSTab1.TabVisible(1)
AddScripts frmMain.ActiveForm.RTF1.Text, frmMain.tvS, Not frmMain.SSTab1.TabVisible(2)
Else
frmMain.tvD.Nodes.Clear
frmMain.tvS.Nodes.Remove frmMain.tvS.Nodes("Document").Index
End If
frmMain.SB.Panels(5).Text = " 00:00 @ 28.8 K "
frmMain.SB.Panels(1).Text = "Ready"
If FormsLeft = 0 Then frmMain.SSTab1.Tab = 0: frmMain.tvW.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Caption = "Untitled" Then GoTo n
frmMain.tvW.SelectedItem = frmMain.tvW.Nodes(Caption)
frmMain.tvD.Nodes.Clear
frmMain.tvS.Nodes.Clear
n: 'next
Kill FullPath(ThisPath, "temp.html")
frmMain.ActiveForm.SetFocus
frmMain.ActiveForm.RTF1.SetFocus
SaveValue "ChildState", WindowState
End Sub

Private Sub IE1_DownloadComplete()
frmMain.SB.Panels(1).Text = "Press Ctrl+W to toggle between editing mode and the preview window."
End Sub

Private Sub IE1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
frmMain.SB.Panels(1).Text = "" & (Progress * 100 / ProgressMax) & "% Completed"
frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(2).Picture
If Progress = ProgressMax Then frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
End Sub

Private Sub IE1_StatusTextChange(ByVal Text As String)
On Error Resume Next
frmMain.SB.Panels(1).Text = Text
End Sub

Private Sub IE1_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
Dim s As VbMsgBoxResult
s = MsgBox("The document is trying to close the window." & vbCrLf & "Do you want to close this document?" & vbCrLf & vbCrLf & "Changes will be lost if you haven't saved them.", vbYesNo + vbExclamation, IE1.LocationURL)
If s = vbYes Then Unload Me
Cancel = True
End Sub

Private Sub mnuArrangeIcons_Click()
frmMain.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascadeWin_Click()
frmMain.Arrange vbCascade
End Sub

Private Sub mnuCheckLinks_Click()
On Error Resume Next
If Caption = "Untitled" Then MsgBox "You must save the file first.", vbExclamation, "Check links": RTF1.SetFocus: Exit Sub
frmMain.SB.Panels(1).Text = "Analyzing the document..."
Load frmVal
frmVal.ListLinksInDocWhileValidating RTF1.Text, Caption
If frmVal.lvProbs.ListItems.count > 0 Then frmVal.Show Else Unload frmVal: MsgBox "No broken links found.", vbInformation, "Link validity"
frmMain.SB.Panels(1).Text = ""
End Sub

Sub mnuClrPicker_Click()
frmCPick.Show vbModal
End Sub

Sub mnuColourFade_Click()
frmCFade.Show vbModal
End Sub

Private Sub mnuConvert_Click()
On Error Resume Next
Dim lpF As Form
Load frmConv
frmConv.SSTab1.Tab = 1
For Each lpF In Forms
If lpF.Name = "frmChild" And lpF.Caption <> "Untitled" Then frmConv.lsFiles.AddItem lpF.Caption
Next lpF
frmConv.Show vbModal
End Sub

Private Sub mnuDeleteBody_Click()
mnuSelectBody_Click
RTF1.SelText = ""
End Sub

Sub mnuDocProps_Click()
frmProps.Show vbModal
End Sub

Sub mnuDocumentConvert_Click()
On Error Resume Next
Load frmConv
frmConv.rtfTemp.Text = RTF1.Text
frmConv.lbFile.Text = Caption
frmConv.rtfTemp.Text = Replace(frmConv.rtfTemp.Text, Chr(10), "")
frmConv.rtfTemp.Text = Replace(frmConv.rtfTemp.Text, Chr(13), "")
frmConv.PB.Max = Len(RTF1.Text)
frmConv.Show vbModal
End Sub

Private Sub mnuEdit_Click()
On Error Resume Next
mnuGotoLastPos.Enabled = (Positions.count > 0)
mnuEditCut.Enabled = frmMain.tbEdit.Buttons("cut").Enabled
mnuEditCopy.Enabled = frmMain.tbEdit.Buttons("copy").Enabled
mnuEditPaste.Enabled = frmMain.tbEdit.Buttons("paste").Enabled
mnuSelectAll.Enabled = (RTF1.SelLength <> Len(RTF1.Text)) And (Len(RTF1.Text) > 0)
Dim s As String, src As String
s = GetLineText
src = ReadAttrib("src", s)
If src = s Then src = ReadAttrib("href", s)
mnuOpenThisFile.Enabled = (src <> s)
frmMain.tbEdit.Buttons("opensrc").Enabled = mnuOpenThisFile.Enabled
End Sub

Sub mnuEditCopy_Click()
Clipboard.Clear
Clipboard.SetText RTF1.SelText
End Sub

Sub mnuEditCut_Click()
Clipboard.Clear
Clipboard.SetText RTF1.SelText
RTF1.SelText = ""
End Sub

Sub mnuEditDefinition_Click()
On Error Resume Next
If CurrentWord = "" Then Exit Sub
Dim lPos As Long, modeLen As Long
Dim SelLen As Long
SelLen = Len(CurrentWord)
lPos = InStr(1, RTF1.Text, "function " & CurrentWord): modeLen = 9
If lPos = 0 Then lPos = InStr(1, RTF1.Text, "var " & CurrentWord): modeLen = 4
'If lPos = 0 Then lPos = InStr(1, RTF1.Text, "(" & CurrentWord): ModeLen = 1
'If lPos = 0 Then lPos = InStr(1, RTF1.Text, ", " & CurrentWord): ModeLen = 2
'''''above ones aren't perfect'''''
If lPos = 0 Then lPos = InStr(1, RTF1.Text, "name=" & Chr(34) & CurrentWord): modeLen = 7
If lPos = 0 Then lPos = InStr(1, RTF1.Text, "name=" & CurrentWord): modeLen = 5
If lPos = 0 Then lPos = InStr(1, RTF1.Text, "id=" & Chr(34) & CurrentWord): modeLen = 5
If lPos = 0 Then lPos = InStr(1, RTF1.Text, "id=" & CurrentWord): modeLen = 3
If lPos > 0 Then
RTF1.SelStart = lPos - 1
If ReadValue("SelectFind", True, "Documents") = True Then RTF1.SelLength = modeLen + SelLen
Else
MsgBox "Can't define " & CurrentWord & ":" & vbCrLf & "Unrecognized Identifier.", vbExclamation, "Definition"
End If
RTF1.SetFocus
End Sub

Sub mnuEditFind_Click()
Load frmFind
frmFind.SSTab1.Tab = 0
frmFind.Show 'vbModal
End Sub

Private Sub mnuEditGoTo_Click()
Load frmFind
frmFind.SSTab1.Tab = 3
frmFind.Show 'vbModal
End Sub

Sub mnuEditPaste_Click()
On Error Resume Next
Dim strD As String
strD = Clipboard.GetText(vbCFText)
RTF1.SelText = strD
End Sub

Sub mnuEditRedo_Click()
Redo
End Sub

Private Sub mnuEditReplace_Click()
Load frmFind
frmFind.SSTab1.Tab = 1
frmFind.Show 'vbModal
End Sub

Sub mnuEditUndo_Click()
Undo
End Sub

Private Sub mnuEntitize_Click()
Dim s As String
Dim i As Integer
For i = 1 To Len(RTF1.SelText)
s = s & "&#" & Asc(Mid(RTF1.SelText, i, 1)) & ";"
Next i
RTF1.SelText = s
End Sub

Private Sub mnuFileClose_Click()
Unload Me
End Sub

Private Sub mnuFileExit_Click()
Unload frmMain
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
On Error Resume Next
Dim lpF As New frmChild
Load lpF
lpF.LoadHTMLFile mnuFileMRU(Index).Tag
End Sub

Private Sub mnuFileNew_Click()
frmMain.NewDocument
End Sub

Sub LoadHTMLFile(lpFileName As String)
On Error Resume Next
Dim lpF As Form
For Each lpF In Forms
  If LCase(lpF.Caption) = LCase(lpFileName) Then
    lpF.SetFocus
    Unload Me
  End If
Next lpF
RTF1.Text = "Loading " & lpFileName & "...": bChanged = False: trapUndo = False
If Dir(lpFileName) = "" Then MsgBox lpFileName & vbCrLf & "The specified file cannot be found.", vbExclamation, "Error": Unload Me: Exit Sub
Select Case LCase(Right(lpFileName, 3))
Case "gif", "bmp", "jpg", "ico"
LoadImage lpFileName
Unload Me
Exit Sub
Case "htm", "tml", "asp", "xml", "htt"
Icon = p1.Picture
Case "css", "ini", "bat", "cfg", "inf", "txt", ".js", "vbs", "tm_"
Icon = p2.Picture
Case Else
Dim s As VbMsgBoxResult
s = MsgBox("The file you chose to open may not be supported." & vbCrLf & "Do you want to open it in its default program instead?", vbYesNoCancel + vbExclamation, lpFileName)
    If s = vbCancel Then Unload Me: Exit Sub
    If s = vbNo Then Icon = p2.Picture: GoTo loadl
    If s = vbYes Then
        Unload Me
        If ShellExecute(0, "open", lpFileName, "", Up1Level(lpFileName), 10) < 32 Then MsgBox "Could not execute " & GetFile(lpFileName), vbExclamation
    End If
Exit Sub
End Select
loadl:
Caption = lpFileName
RTF1.LoadFile lpFileName, rtfText
RTF1.SetFocus
Form_GotFocus
bChanged = False
GetSpeedInfo Len(RTF1.Text)
Outline Me, frmMain.tvD, RTF1.Text, False
AddScripts RTF1.Text, frmMain.tvS, False
AddFileMRU lpFileName
EnableBar
trapUndo = True
UndoStack.Remove UndoStack.count
mnuStyleEdit.Enabled = (LCase(Ext(Caption)) = "css")
EnableControls
ParseDate
RTF1.SelStart = 0
End Sub

Private Sub mnuFileOpen_Click()
Dim lpF As New frmChild
On Error GoTo hell
With frmMain.CD
.ShowOpen
Load lpF
lpF.LoadHTMLFile .Filename
End With
hell:
End Sub

Sub mnuFilePrint_Click()
    ' Check to se if a printer is installed
    Dim pTest
    
    pTest = Printer.papersize
    
    If pTest = 0 Then   ' no printer installed
        GoTo errHandler
    End If

     Call printText
    
    
    Exit Sub

errHandler:
    MsgBox "No printer found! Printing not possible.", vbCritical
    Exit Sub

End Sub

Private Sub mnuFilePrinterSetup_Click()
    Load frmPageSetup
    frmPageSetup.txtHeader.Text = sPrintHeader 'set defaults
    frmPageSetup.txtFooter.Text = sPrintFooter
    frmPageSetup.Show vbModal
End Sub

Sub mnuFilePrintPreview_Click()
    
    'setup header and footer
    sPrintText = RTF1.Text
    sHeader = SetPrintLine(sPrintHeader)
    sFooter = SetPrintLine(sPrintFooter)
    sPrintText = sHeader & vbCrLf & vbCrLf & sPrintText & vbCrLf & vbCrLf & sFooter
    Me.rtfTmp.Text = sPrintText

    frmDocPreview.Show vbModal

End Sub

Private Sub mnuFileRevert_Click()
On Error Resume Next
If Left(Caption, 8) = "Untitled" Then Exit Sub
If MsgBox("Do you want to revert to the saved version?", vbExclamation + vbYesNo) = vbYes Then
bChanged = False
RTF1.LoadFile Caption
RTF1.SetFocus
End If
End Sub

Sub mnuFileSave_Click()
If Left(Caption, 8) = "Untitled" Or Mid(Caption, 2, 1) <> ":" Then mnuFileSaveAs_Click: Exit Sub
SaveHTMLFile Caption
mnuFileSave.Enabled = False
frmMain.TB.Buttons(3).Enabled = False
End Sub

Sub mnuFileSaveAll_Click()
On Error Resume Next
Dim lpF As Form
For Each lpF In Forms
If lpF.BackColor = &H8000000F Then
    lpF.SetFocus
    lpF.mnuFileSave_Click
End If
Next lpF
End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo hell
If Left(RTF1.Text, 1) <> "<" Then frmMain.CD.FilterIndex = 3 Else frmMain.CD.FilterIndex = 1 'sort of an HTML filter
frmMain.CD.Filename = IIf(Caption <> "Untitled", Caption, "")
frmMain.CD.ShowSave
SaveHTMLFile frmMain.CD.Filename
hell:
RTF1.SetFocus
End Sub

Private Sub mnuFindFiles_Click()
Load frmFind
frmFind.SSTab1.Tab = 2
frmFind.Show vbModal
End Sub

Sub mnuFindNext_Click()
RTF1.Find LastFindText, RTF1.SelStart + 1
End Sub

Sub mnuGotoLastPos_Click()
On Error Resume Next
If Positions.count = 0 Then Exit Sub
RTF1.SelStart = Positions.Item(Positions.count)
Positions.Remove Positions.count
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

Sub mnuInsertDateTime_Click()
frmDate.Show vbModal
End Sub

Private Sub mnuInsertNetscapeFix_Click()
On Error Resume Next
Dim orgSelS As Long
If MsgBox("The Netscape ResizeWindow fix allows you to fix the" & vbCrLf & "problem encountered while resizing NS4 windows." & vbCrLf & vbCrLf & "Insert this script into the document?", vbYesNo + vbQuestion, "NS4 Bug Fix") = vbNo Then RTF1.SetFocus: Exit Sub
orgSelS = RTF1.SelStart + 1
Dim Content As String
Content = vbCrLf & "<SCRIPT language='JavaScript'>" & vbCrLf & "<!--" & vbCrLf & "function MM_reloadPage(init) {  //reloads the window if Nav4 resized" & vbCrLf & "  if (init==true) with (navigator) {if ((appName=='Netscape')&&(parseInt(appVersion)==4)) {" & vbCrLf & "    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}" & vbCrLf & "  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();" & vbCrLf & "}" & vbCrLf & "MM_reloadPage(true);" & vbCrLf & "// -->" & vbCrLf & "</SCRIPT>"
Dim pos As Long
pos = InStr(1, RTF1.Text, "<head>", vbTextCompare)
If pos <> 0 Then
pos = pos + 6
If pos < orgSelS Then orgSelS = orgSelS + Len(Content) + 1
Else
pos = 1
End If
RTF1.SelStart = pos - 1
RTF1.SelText = Content
RTF1.SetFocus
RTF1.SelStart = orgSelS - 1
End Sub

Sub mnuInsertSymbol_Click()
On Error Resume Next
frmSymbols.Show vbModal
RTF1.SetFocus
End Sub

Private Sub mnuOpenRemote_Click()
frmGet.Show vbModal
End Sub

Sub mnuOpenThisFile_Click()
On Error Resume Next
Dim s As String, src As String
'MODES: 1 is web, 0 is local
Dim MODE As Long
s = GetLineText()
src = ReadAttrib("src", s): MODE = 0
If src = s Then src = ReadAttrib("href", s): MODE = 1
If src = s Then MsgBox "The line pointed at does not contain a reference to any document.", vbExclamation: Exit Sub
If MODE = 0 Then
  Dim pt As String
  pt = IIf(Caption = "Untitled", App.Path, Up1Level(Caption))
  pt = FullPath(pt, src)
  If Dir(src) <> "" Then pt = src
  Select Case LCase(Ext(src))
  Case "htm", "html", "asp", "js", "txt"
    Dim l As New frmChild
    Load l
    l.LoadHTMLFile pt
  Case "gif", "jpg", "bmp"
    LoadImage pt
  Case Else
    GoTo shellexec
  End Select
Else
  If MsgBox("HyperText reference found to:" & vbCrLf & src & vbCrLf & vbCrLf & "Click OK to open.", vbOKCancel + vbInformation) = vbCancel Then Exit Sub
shellexec:
  ShellExecute hWnd, "open", src, "", IIf(Caption = "Untitled", App.Path, Up1Level(Caption)), 10
  End If
End Sub

Sub mnupasteSpecialBRasP_Click()
On Error Resume Next
Dim s As String, i As Long
s = Clipboard.GetText(vbCFText)
Dim st() As String
s = Replace(s, Chr(10), "")
st = Split(s, Chr(13))
s = ""
For i = 0 To UBound(st)
s = s & "<P>" & vbCrLf & st(i) & vbCrLf & "</P>" & vbCrLf
Next i
RTF1.SelText = s
End Sub

Sub mnuPasteSpecialOrdered_Click()
On Error Resume Next
Dim s As String, i As Long
s = Clipboard.GetText(vbCFText)
Dim st() As String
s = Replace(s, Chr(10), "")
st = Split(s, Chr(13))
s = ""
RTF1.SelText = "<OL>" & vbCrLf
For i = 0 To UBound(st)
s = s & "<LI>" & st(i) & "</LI>" & vbCrLf
Next i
RTF1.SelText = s & "</OL>" & vbCrLf
End Sub

Sub mnuPasteSpecialPRE_Click()
On Error Resume Next
Dim s As String
s = Clipboard.GetText(vbCFText)
RTF1.SelText = "<PRE>" & vbCrLf & s & "</PRE>" & vbCrLf
End Sub

Sub mnuPasteSpecialTagsEntities_Click()
On Error Resume Next
Dim s As String
s = Clipboard.GetText(vbCFText)
s = Replace(s, "&", "&amp;")
s = Replace(s, "<", "&lt;")
s = Replace(s, "&reg;", "®")
s = Replace(s, ">", "&gt;")
s = Replace(s, "©", "&copy;")
RTF1.SelText = s
End Sub

Sub mnuPasteSpecialUnordered_Click()
On Error Resume Next
Dim s As String, i As Long
s = Clipboard.GetText(vbCFText)
Dim st() As String
s = Replace(s, Chr(10), "")
st = Split(s, Chr(13))
s = ""
RTF1.SelText = "<UL>" & vbCrLf
For i = 0 To UBound(st)
s = s & "<LI>" & st(i) & "</LI>" & vbCrLf
Next i
RTF1.SelText = s & "</UL>" & vbCrLf
End Sub

Sub mnuPreview_Click()
mnuFileSave_Click 'save it first
If Left(Caption, 8) = "Untitled" Or Mid(Caption, 2, 1) <> ":" Then
frmMain.SB.Panels(1).Text = "The file must be saved before you can preview it."
Exit Sub
End If
ShellExecute Me.hWnd, "open", "explorer", Caption, "", 10
End Sub

Private Sub mnuPrintPage_Click()
On Error Resume Next
mnuViewCodePage.Caption = "View &Page"
mnuViewCodePage_Click
IE1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub mnuRestoreAll_Click()
Dim lpF As Form
For Each lpF In Forms
If lpF.Name = "frmChild" Then
lpF.WindowState = vbNormal
End If
Next lpF
End Sub

Sub mnuSelectAll_Click()
RTF1.SelStart = 0
RTF1.SelLength = Len(RTF1.Text)
End Sub

Private Sub mnuSelectBody_Click()
On Error Resume Next
Dim l1 As Long, l2 As Long, braces As Long, i As Long
Dim RTF1Text As String
RTF1Text = NoStrings(NoComments(RTF1.Text))
If frmMain.tvS.Nodes.count <= 1 Then Exit Sub
If frmMain.tvS.SelectedItem Is Nothing Then Exit Sub

If frmMain.tvS.SelectedItem.Image = "var" And frmMain.tvS.SelectedItem.parent.Key = "globals" Then
  Dim s As String
  s = frmMain.tvS.SelectedItem.Key
  s = Trim(s)
  RTF1.SelStart = InStr(RTF1.SelStart + 1, RTF1Text, s) - 1 'var name
  l1 = InStr(RTF1.SelStart + 1, RTF1Text, ";")
  If InStr(RTF1.SelStart + 1, RTF1Text, vbNewLine) < l1 Then l1 = InStr(RTF1.SelStart + 1, RTF1Text, vbNewLine)
  If InStr(RTF1.SelStart + 1, RTF1Text, ",") < l1 Then l1 = InStr(RTF1.SelStart + 1, RTF1Text, ",")
  s = Mid$(RTF1Text, RTF1.SelStart + 1, l1 - RTF1.SelStart + 1)
  RTF1.SelLength = Len(s)
  RTF1.SetFocus
Exit Sub
End If


If frmMain.tvS.SelectedItem.Image = "var" And frmMain.tvS.SelectedItem.parent.Key <> "globals" Then
  Dim ss() As String
  ss = Split(frmMain.tvS.SelectedItem.Key, ": local to ")
  ss(0) = Trim(ss(0))
  ss(1) = Trim(ss(1))
  RTF1.SelStart = InStr(1, RTF1Text, ss(1)) - 1 'function name
  RTF1.SelStart = InStr(RTF1.SelStart + 1, RTF1Text, ss(0)) - 1 'var name
  l1 = InStr(RTF1.SelStart + 1, RTF1Text, ";")
  If InStr(RTF1.SelStart + 1, RTF1Text, vbNewLine) < l1 Then l1 = InStr(RTF1.SelStart + 1, RTF1Text, vbNewLine)
  If InStr(RTF1.SelStart + 1, RTF1Text, ",") < l1 Then l1 = InStr(RTF1.SelStart + 1, RTF1Text, ",")
  ss(1) = Mid$(RTF1Text, RTF1.SelStart + 1, l1 - RTF1.SelStart + 1)
  If ReadValue("SelectFind", , "Documents") = True Then RTF1.SelLength = Len(ss(1))
Exit Sub
End If

l1 = InStr(1, RTF1Text, frmMain.tvS.SelectedItem.Key)
If l1 = 0 Then Exit Sub
l2 = InStr(l1 + 1, RTF1Text, "}")

braces = StrCount(Mid$(RTF1Text, l1, l2 - l1), "{") - 1

If braces = -1 Then GoTo n

  For i = 1 To braces
    l2 = InStr(l2 + 1, RTF1Text, "}")
  Next i
n:

RTF1.SelStart = l1 - 1
RTF1.SelLength = l2 - l1 + 1
RTF1.SetFocus
End Sub

Private Sub mnuSpecialPaste_Click()
frmPS.Show vbModal
End Sub

Private Sub mnuStats_Click()
Load frmStats
frmStats.ShowStats RTF1.Text
frmStats.Show vbModal
End Sub

Private Sub mnuStyleEdit_Click()
On Error Resume Next
frmCSS.Show vbModal
RTF1.SetFocus
End Sub

Private Sub mnuTileHorizontal_Click()
frmMain.Arrange vbHorizontal
End Sub

Private Sub mnuTileVertical_Click()
frmMain.Arrange vbVertical
End Sub

Private Sub mnuTodayTip_Click()
On Error Resume Next
frmTip.Show vbModal
End Sub

Sub mnuUpdate_Click()
On Error Resume Next
AddScripts RTF1.Text, frmMain.tvS, Not frmMain.SSTab1.TabVisible(2)
Outline Me, frmMain.tvD, RTF1.Text, Not frmMain.SSTab1.TabVisible(1)
frmMain.tvD.Nodes.Item("Main").Expanded = True
frmMain.tvD.SelectedItem = frmMain.tvD.Nodes.Item("Main")
End Sub

Private Sub mnuUsingTemplate_Click()
frmTemp.Show vbModal
End Sub

Private Sub mnuValidDoc_Click()
On Error Resume Next
Dim whole As String, probs As String
Dim pos As Long, pos2 As Long, posTmp As Long, posOrg As Long
'check for combinable font tags
frmMain.SB.Panels(1).Text = "Analyzing the document..."
Do
pos = InStr(pos2 + 1, RTF1.Text, "<FONT", vbTextCompare)
posOrg = pos 'save start pos in posOrg
If pos = 0 Then Exit Do
pos = InStr(pos + 1, RTF1.Text, ">")
pos2 = InStr(pos + 1, RTF1.Text, "<FONT", vbTextCompare)
posTmp = InStr(pos + 1, RTF1.Text, "</FONT>", vbTextCompare)
If posTmp < pos2 Then GoTo nxt 'proper, not nested
If pos = 0 Or pos2 = 0 Then Exit Do
RTF1.SelStart = posOrg - 1
probs = probs & "Combinable Nested tag:<FONT>:" & posOrg & vbCrLf
RTF1.SetFocus
nxt:
Loop
'check if titled properly
pos = InStr(1, RTF1.Text, "<TITLE>", vbTextCompare)
pos2 = InStr(1, RTF1.Text, "</TITLE>", vbTextCompare)
If pos = 0 Or pos2 = 0 Or pos2 < pos Then
probs = probs & "Improper or missing tag:<TITLE>:" & pos & vbCrLf
RTF1.SetFocus
RTF1.SelStart = pos - 1
End If
'check ALT in images
Do
pos = InStr(pos2 + 1, RTF1.Text, "<IMG", vbTextCompare)
pos2 = InStr(pos + 1, RTF1.Text, ">") + 1
If pos = 0 Or pos2 = 1 Then Exit Do 'we've already added 1 to pos2
whole = Mid$(RTF1.Text, pos, pos2 - pos)
If ReadAttrib("alt", whole) = whole Then
  RTF1.SelStart = pos - 1
  probs = probs & "Missing [alt] attribute:<IMG>:" & pos & vbCrLf
End If
Loop
'check similar tags that are immediate
Dim pos_pos As Long, pos_pos2 As Long
Dim firsttag, secondtag
Do
pos = InStr(pos2 + 1, RTF1.Text, "<")
pos2 = InStr(pos + 1, RTF1.Text, ">")
If pos = 0 Or pos2 = 0 Then Exit Do
firsttag = Mid$(RTF1.Text, pos, pos2 - pos + 1)
posTmp = InStr(pos2 + 1, RTF1.Text, "</" & Mid$(firsttag, 2), vbTextCompare) 'find closing tag
pos_pos = InStr(posTmp + 1, RTF1.Text, "<")
pos_pos2 = InStr(pos_pos + 1, RTF1.Text, ">")
If pos_pos = 0 Or pos_pos2 = 0 Then GoTo nxt2
secondtag = Mid$(RTF1.Text, pos_pos, pos_pos2 - pos_pos + 1)
If LCase(firsttag) = LCase(secondtag) Then probs = probs & "Similar Immediate Tag:" & UCase(firsttag) & ":" & pos & vbCrLf
nxt2:
Loop
'check redundant nested tags
Do
pos = InStr(pos2 + 1, RTF1.Text, "<")
pos2 = InStr(pos + 1, RTF1.Text, ">")
If pos = 0 Then Exit Do
If pos2 = 0 Then GoTo nxt3
firsttag = Mid$(RTF1.Text, pos, pos2 - pos + 1)
If Left(firsttag, 2) = "<!" Or Left(firsttag, 2) = "</" Then GoTo nxt3
posTmp = InStr(pos2 + 1, RTF1.Text, "</" & Mid$(firsttag, 2), vbTextCompare) 'find closing tag
If posTmp = 0 Then GoTo nxt3
whole = Mid$(RTF1.Text, pos2 + 1, posTmp - pos2 - 1)
If InStr(1, whole, firsttag) > 0 And InStr(1, whole, "</" & Mid$(firsttag, 2)) > InStr(1, whole, firsttag) Then probs = probs & "Redundant nested tags:" & UCase(GetTag(firsttag)) & ":" & pos & vbCrLf
nxt3:
Loop
frmMain.SB.Panels(1).Text = ""
'if all finishes well
If probs = "" Then
  MsgBox "No problems detected.", vbInformation, "Validation"
  RTF1.SetFocus
Else
  Load frmVal
  Dim prbsArr() As String, arrItem() As String
  prbsArr = Split(probs, vbCrLf)
  For pos = 0 To UBound(prbsArr)
    arrItem = Split(prbsArr(pos), ":")
    frmVal.lvProbs.ListItems.Add , "tmp", arrItem(0), , 1
    frmVal.lvProbs.ListItems("tmp").ListSubItems.Add 1, , arrItem(1)
    frmVal.lvProbs.ListItems("tmp").Tag = arrItem(2)
    frmVal.lvProbs.ListItems("tmp").Key = ""
  Next pos
  frmVal.Show
End If
End Sub

Private Sub mnuView_Click()
mnuViewTask.Checked = frmMain.SSTab1.TabVisible(3)
mnuViewScripts.Checked = frmMain.SSTab1.TabVisible(2)
mnuViewDocuments.Checked = frmMain.SSTab1.TabVisible(1)
mnuViewToolBar.Checked = frmMain.CBR.Bands(1).Visible
mnuViewStatusBar.Checked = frmMain.SB.Visible
mnuViewFileTree.Checked = frmMain.pLeft.Visible
End Sub

Sub mnuViewCodePage_Click()
On Error Resume Next
Select Case Left(mnuViewCodePage.Caption, 10)
Case "View &Page"
  If Caption = "Untitled" Then ThisPath = App.Path Else ThisPath = Up1Level(Caption)
  Open FullPath(ThisPath, "temp.html") For Output As #1
    Print #1, RTF1.Text
  Close #1
  'I could have used /temp.html as app.path does not put in / unless it is a root drive, but if i dont, it will be put in the parent folder, but who cares as long as we can access it
  IE1.Navigate FullPath(ThisPath, "temp.html")
  IE1.ZOrder vbBringToFront
  IE1.SetFocus
  mnuViewCodePage.Caption = "View &Code" & vbTab & "Ctrl+W"
Case "View &Code"
  IE1.Navigate "about:blank"
  frmMain.SB.Panels(1).Text = ""
  RTF1.ZOrder vbBringToFront
  RTF1.SetFocus
  mnuViewCodePage.Caption = "View &Page" & vbTab & "Ctrl+W"
End Select
End Sub

Private Sub mnuViewDocuments_Click()
mnuViewDocuments.Checked = Not mnuViewDocuments.Checked
frmMain.SSTab1.TabVisible(1) = mnuViewDocuments.Checked
SaveValue "DocumentTree", mnuViewDocuments.Checked
End Sub

Private Sub mnuViewFileTree_Click()
frmMain.mnuViewFileTree_Click
End Sub

Private Sub mnuViewMode_Click(Index As Integer)
MousePointer = 11: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(2).Picture
Dim i As Integer
For i = 1 To 3
mnuViewMode(i).Checked = False
Next i
mnuViewMode(Index).Checked = True
SetViewMode Index - 1, RTF1
SaveValue "ViewMode", Index - 1, "Documents"
MousePointer = 0: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
End Sub

Private Sub mnuViewOptions_Click()
frmOpts.Show vbModal
End Sub

Private Sub mnuViewStatusBar_Click()
mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
frmMain.SB.Visible = mnuViewStatusBar.Checked
SaveValue "Statusbar", frmMain.SB.Visible
End Sub

Private Sub mnuViewToolBar_Click()
mnuViewToolBar.Checked = Not mnuViewToolBar.Checked
frmMain.CBR.Bands(1).Visible = mnuViewToolBar.Checked
frmMain.CBR.Bands(2).Visible = frmMain.CBR.Bands(1).Visible
frmMain.CBR.Bands(3).Visible = frmMain.CBR.Bands(1).Visible
SaveValue "Toolbar", frmMain.CBR.Bands(1).Visible
End Sub

Private Sub mnuWindowMaximizeAll_Click()
WindowState = vbMaximized 'that's all
End Sub

Private Sub mnuWindowMinimizeAll_Click()
On Error Resume Next
Dim lpC As Form
For Each lpC In Forms
If lpC.BackColor = &H8000000F Then 'is not MDI
lpC.WindowState = vbMinimized
End If
Next lpC
End Sub

Sub mnuWindowUnloadAll_Click()
On Error Resume Next
If MsgBox("Close all document windows?", vbYesNo + vbExclamation, "WonderHTML") = vbNo Then Exit Sub
Dim lpC As Form
For Each lpC In Forms
If lpC.Name = "frmChild" Then
Unload lpC
End If
Next lpC
End Sub

Private Sub RTF1_Change()
On Error Resume Next
bChanged = True

If mnuViewCodePage.Caption = "View &Code" & vbTab & "Ctrl+W" Then mnuViewCodePage_Click

frmMain.TB.Buttons(3).Enabled = True
mnuFileSave.Enabled = True
If Not trapUndo Then Exit Sub 'because trapping is disabled

Dim newElement As New UndoElement    'create new undo element

Dim c%, l&

'remove all redo items because of the change
For c% = 1 To RedoStack.count
    RedoStack.Remove 1
Next c%

'set the values of the new element
newElement.SelStart = RTF1.SelStart
newElement.TextLen = Len(RTF1.Text)
newElement.Text = RTF1.Text

'add it to the undo stack
UndoStack.Add newElement

If UndoStack.count > 50 Then UndoStack.Remove 1
'well.... changed undo limit to 25. would you EVER use 100?

EnableControls
End Sub

Sub RTF1_GotFocus()
On Error Resume Next
frmMain.tvW.SelectedItem = frmMain.tvW.Nodes(Caption)
frmMain.SB.Panels(1).Text = GetTitleFromText(RTF1.Text) & " (" & GetFile(Caption) & ")"
'check if bUpdateFlag and the DocTree's visible property both favour the outlining process
Outline Me, frmMain.tvD, RTF1.Text, (bUpdateFlag) Or (Not frmMain.SSTab1.TabVisible(1)) Or (FormsLeft < 2)
AddScripts RTF1.Text, frmMain.tvS, (bUpdateFlag) Or (Not frmMain.SSTab1.TabVisible(2)) Or (FormsLeft < 2)
noOutline:
frmMain.TB.Buttons(3).Enabled = bChanged
mnuFileSave.Enabled = bChanged
RTF1_SelChange
GetSpeedInfo Len(RTF1.Text)
End Sub

Private Sub RTF1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If ReadValue("AutoIndent", True, "Documents") = True And KeyCode = vbKeyReturn And Shift = 0 Then AutoIndent: KeyCode = 0

If Left(mnuViewCodePage.Caption, 10) = "View &Code" Then mnuViewCodePage_Click

If Shift = vbCtrlMask And KeyCode = vbKeyReturn Then
KeyCode = 0
RTF1.SelText = "<P>" & vbCrLf & vbCrLf & "</P>"
RTF1.SelStart = RTF1.SelStart - 6
End If

If Shift = vbShiftMask And KeyCode = vbKeyReturn Then
KeyCode = 0
RTF1.SelText = "<BR>"
End If

Dim s As String, sel() As String

If Shift = vbShiftMask + vbCtrlMask Then
s = ReadValue("+" & KeyCode & "Text", "", "Macros")
If s = "" Then GoTo n
s = Replace(s, "\n", vbCrLf)
If frmMain.IsInQuotes(RTF1.SelStart) Then frmMain.GoOutsideQuotes RTF1.SelStart
RTF1.SelText = s
s = ReadValue("+" & KeyCode & "Sel", ";", "Macros")
sel = Split(s, ";")
RTF1.SelStart = RTF1.SelStart - CLng(sel(0))
RTF1.SelLength = CLng(sel(1))
KeyCode = 0
End If
n:

If KeyCode = vbKeyTab And Shift = 0 Then
    If RTF1.SelLength > 0 And InStr(RTF1.SelText, Chr(13)) > 0 Then
        KeyCode = 0
        Dim StrP() As String, i As Integer
        Dim lngStart As Long, oldSelL As Long, oldSelS As Long
        lngStart = SendMessage(RTF1.hWnd, EM_LINEINDEX, GetCurrentLine(RTF1) - 1, 0&)
        oldSelL = RTF1.SelLength
        oldSelS = RTF1.SelStart
        RTF1.SelStart = lngStart
        RTF1.SelLength = oldSelL + (oldSelS - lngStart)
        StrP = Split(RTF1.SelText, vbNewLine)
        For i = 0 To UBound(StrP) - 1
        RTF1.SelText = IIf(Shift = 1, vbTab, Space(4)) & StrP(i) & vbNewLine
        Next i
        If StrP(UBound(StrP)) <> "" Then RTF1.SelText = IIf(Shift = 1, vbTab, Space(4)) & StrP(UBound(StrP))
    Else
        RTF1.SelText = Space(4)
        KeyCode = 0
    End If
End If
'don't allow RTF Boxes' default undo proc
If (Shift = vbCtrlMask) Then
    Select Case KeyCode
        Case vbKeyZ
            KeyCode = 0
            If mnuEditUndo.Enabled = True Then mnuEditUndo_Click
        Case vbKeyY
            KeyCode = 0
            If mnuEditRedo.Enabled = True Then mnuEditRedo_Click
        Case vbKeyV
            KeyCode = 0
            If mnuEditPaste.Enabled = True Then mnuEditPaste_Click
    End Select
End If
RTF1.SetFocus
End Sub

Private Sub RTF1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If ReadValue("AutoLink", True, "Documents") = False Then Exit Sub
Dim MoreInf As String, CType As String
CType = CurrentWord
If KeyAscii = 32 Then
    If IsLink(CurrentWord) = True Then
        MoreInf = GetMoreInfo(CType)
        RTF1.SelStart = RTF1.SelStart - Len(CType)
        RTF1.SelText = "<A href=" & Chr(34) & MoreInf & CType & Chr(34) & ">"
        RTF1.SelStart = RTF1.SelStart + Len(CType)
        RTF1.SelText = "</A>"
        KeyAscii = 0
    End If
    CurrentWord = ""
ElseIf KeyAscii = 8 Then
    If Len(CurrentWord) > 1 Then CurrentWord = Left(CurrentWord, Len(CurrentWord) - 1)
Else
    CurrentWord = CurrentWord & Chr$(KeyAscii)
End If
End Sub

Private Sub RTF1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If Button = 2 Then mnuEdit_Click: PopupMenu mnuEdit
End Sub

Sub SaveHTMLFile(lpFileName As String)
On Error Resume Next
ParseDate
Open lpFileName For Output As #1
Print #1, chomp(RTF1.Text)
Close #1
Kill "Untitled.wbackup"
Kill lpFileName & ".wbackup"
Caption = lpFileName
If frmMain.IsWebOpen And GetFile(lpFileName) <> "files.inf" Then frmMain.tvW.Nodes.Add Up1Level(lpFileName), tvwChild, lpFileName, GetFile(lpFileName), FileIcon(lpFileName)
AddFileMRU lpFileName
frmMain.tvS.Nodes("Document").Text = GetFile(Caption)
mnuUpdate_Click
GetSpeedInfo Len(RTF1.Text)
mnuStyleEdit.Enabled = (LCase(Ext(Caption)) = "css")
Outline Me, frmMain.tvD, RTF1.Text, False
AddScripts RTF1.Text, frmMain.tvS, False
bChanged = False
End Sub

Private Sub RTF1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
CurrentWord = RichWordOver(RTF1, X, y)
End Sub

Private Sub RTF1_SelChange()
On Error Resume Next

If Positions.count = 0 Then GoTo yes Else frmMain.tbMore.Buttons(1).Enabled = True
If Positions(Positions.count) = RTF1.SelStart Then GoTo skip
yes:
If Positions.count > 25 Then Positions.Remove 1
Positions.Add RTF1.SelStart
skip:

Dim pC As Integer, lp As POINTAPI
frmMain.SB.Panels(2).Text = " Line " & GetCurrentLine(RTF1) & ", Col " & GetColumnIndex(RTF1) & ", Sel " & RTF1.SelLength & " "
frmMain.SB.Panels(3).Text = " " & GetTotalLines(RTF1) & " Lines "

If Left(frmMain.SB.Panels(1).Text, 32) = "Please wait, converting to ASCII" Then
pC = Round((RTF1.SelStart * 100) / Len(RTF1.Text), 0)
frmMain.SB.Panels(1).Text = "Please wait, converting to ASCII (" & pC & "%)"
End If

GetCaretPos lp
CurrentWord = RichWordOver(RTF1, lp.X * Screen.TwipsPerPixelX + RTF1.Left, lp.y * Screen.TwipsPerPixelY + RTF1.Top)

Dim ln&
If Not trapUndo Then Exit Sub
ln& = RTF1.SelLength
mnuEditCut.Enabled = ln&    'disabled if length of selected text is 0
mnuEditCopy.Enabled = ln&   'disabled if length of selected text is 0
mnuEditPaste.Enabled = Len(Clipboard.GetText(1)) 'disabled if length of clipboard text is 0
mnuSelectAll.Enabled = CBool(Len(RTF1.Text)) 'disabled if length of textbox's text is 0
    
With frmMain.tbEdit
    .Buttons("cut").Enabled = mnuEditCut.Enabled
    .Buttons("copy").Enabled = mnuEditCopy.Enabled
    .Buttons("paste").Enabled = (Clipboard.GetFormat(vbCFText) = True)
    .Buttons("undo").Enabled = UndoStack.count > 1
    .Buttons("redo").Enabled = RedoStack.count > 0
End With
    
mnuEditUndo.Enabled = frmMain.tbEdit.Buttons("undo").Enabled
mnuEditRedo.Enabled = frmMain.tbEdit.Buttons("redo").Enabled
End Sub

Function GetSelStart() As Long
Dim lpS As Long
lpS = RTF1.Find("<BODY", , , rtfNoHighlight)
lpS = RTF1.Find(">", lpS + 1, , rtfNoHighlight)
GetSelStart = lpS + 3
End Function

Sub SetMenus()
On Error Resume Next
'Menus if given mnemonics don't work
mnuFileClose.Caption = mnuFileClose.Caption & vbTab & "Ctrl+F4"
mnuFileExit.Caption = mnuFileExit.Caption & vbTab & "Alt+F4"
mnuEditCut.Caption = mnuEditCut.Caption & vbTab & "Ctrl+X"
mnuEditCopy.Caption = mnuEditCopy.Caption & vbTab & "Ctrl+C"
mnuEditPaste.Caption = mnuEditPaste.Caption & vbTab & "Ctrl+V"
mnuEditUndo.Caption = mnuEditUndo.Caption & vbTab & "Ctrl+Z"
mnuEditRedo.Caption = mnuEditRedo.Caption & vbTab & "Ctrl+Y"
mnuViewCodePage.Caption = mnuViewCodePage.Caption & vbTab & "Ctrl+W"
mnuDocProps.Caption = mnuDocProps.Caption & vbTab & "Alt+Enter"
End Sub

Sub CopyMRUList()
On Error Resume Next
Dim i As Integer

For i = 1 To 6

mnuFileMRU(i).Caption = frmMain.mnuFileMRU(i).Caption
mnuFileMRU(i).Tag = frmMain.mnuFileMRU(i).Tag
mnuFileMRU(i).Visible = (Len(mnuFileMRU(i).Caption) > 4)

Next i
End Sub

Private Sub mnuViewScripts_Click()
mnuViewScripts.Checked = Not mnuViewScripts.Checked
frmMain.SSTab1.TabVisible(2) = mnuViewScripts.Checked
SaveValue "ScriptView", mnuViewScripts.Checked
End Sub

Function IsEmptyElement(ByVal Element As String) As Boolean
If Left(Element, 1) = "<" Then Element = Right(Element, Len(Element) - 1)
If Left(Element, 1) = "/" Then Element = Right(Element, Len(Element) - 1)
If Right(Element, 1) = ">" Then Element = Left(Element, Len(Element) - 1)
If Left(Element, 2) = "!-" Then IsEmptyElement = True: Exit Function

Dim iPos As Long
iPos = InStr(1, Element, " ")
If iPos > 0 Then Element = Left(Element, iPos): Element = Trim(Element)
frmMain.SB.Panels(1).Text = Element

Select Case LCase(Element)
Case "hr", "img", "br", "input", "button", "bgsound", "base", "meta", "!doctype", "!--", "isindex"
IsEmptyElement = True
Case Else
IsEmptyElement = False
End Select
End Function

Function GetTag(ByVal Tag As String) As String
Dim iPos As Long
iPos = InStr(1, Tag, " ")
If iPos > 0 Then
GetTag = Trim(Left(Tag, iPos))
Else
GetTag = Tag
End If
End Function

Sub AddComment(ByVal lpStart As Long, lpEndTag As String)
RTF1.SelText = ">"
RTF1.SelStart = lpStart + 1
RTF1.SelText = vbNewLine & "<!--" & vbNewLine & vbNewLine & "// -->" & vbNewLine & lpEndTag
RTF1.SelStart = RTF1.SelStart - Len(lpEndTag) - 10
End Sub



Function GetVM() As Long
Dim i As Integer
For i = 1 To 3
If mnuViewMode(i).Checked Then GetVM = i: Exit Function
Next i
GetVM = 0
End Function

Private Sub mnuViewTask_Click()
mnuViewTask.Checked = Not mnuViewTask.Checked
frmMain.SSTab1.TabVisible(4) = mnuViewTask.Checked
SaveValue "TaskView", frmMain.SSTab1.TabVisible(3)
End Sub

Public Sub Undo()
On Error Resume Next
Dim chg$, X&
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object, objElement2 As Object

    If UndoStack.count > 1 And trapUndo Then 'we can proceed
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.count - 1).TextLen < UndoStack(UndoStack.count).TextLen
        If DeleteFlag Then  'delete some text
            X& = SendMessage(RTF1.hWnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.count)
            Set objElement2 = UndoStack(UndoStack.count - 1)
            RTF1.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            RTF1.SelLength = objElement.TextLen - objElement2.TextLen
            RTF1.SelText = ""
            X& = SendMessage(RTF1.hWnd, EM_HIDESELECTION, 0&, 0&)
        Else 'append something
            Set objElement = UndoStack(UndoStack.count - 1)
            Set objElement2 = UndoStack(UndoStack.count)
            chg$ = Change(objElement.Text, objElement2.Text, objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            RTF1.SelStart = objElement2.SelStart
            RTF1.SelLength = 0
            RTF1.SelText = chg$
            RTF1.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                RTF1.SelLength = Len(chg$)
            Else
                RTF1.SelStart = RTF1.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add UndoStack(UndoStack.count)
        UndoStack.Remove UndoStack.count
    End If
    EnableControls
    trapUndo = True
    RTF1.SetFocus
End Sub

Public Sub Redo()
Dim chg$
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object
If RedoStack.count > 0 And trapUndo Then
    trapUndo = False
    DeleteFlag = RedoStack(RedoStack.count).TextLen < Len(RTF1.Text)
    If DeleteFlag Then  'delete last item
        Set objElement = RedoStack(RedoStack.count)
        RTF1.SelStart = objElement.SelStart
        RTF1.SelLength = Len(RTF1.Text) - objElement.TextLen
        RTF1.SelText = ""
    Else 'append something
        Set objElement = RedoStack(RedoStack.count)
        chg$ = Change(RTF1.Text, objElement.Text, objElement.SelStart + 1)
        RTF1.SelStart = objElement.SelStart - Len(chg$)
        RTF1.SelLength = 0
        RTF1.SelText = chg$
        RTF1.SelStart = objElement.SelStart - Len(chg$)
        If Len(chg$) > 1 And chg$ <> vbCrLf Then
            RTF1.SelLength = Len(chg$)
        Else
            RTF1.SelStart = RTF1.SelStart + Len(chg$)
        End If
    End If
    UndoStack.Add objElement
    RedoStack.Remove RedoStack.count
    trapUndo = True
End If
EnableControls

RTF1.SetFocus
End Sub

Sub ResetUndo()
Dim i As Long
For i = 1 To UndoStack.count
UndoStack.Remove i
Next i
For i = 1 To RedoStack.count
RedoStack.Remove i
Next i
End Sub

Private Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
On Error Resume Next
Dim tempParam$
Dim d&
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function

Private Sub EnableControls()
On Error Resume Next
    mnuEditUndo.Enabled = UndoStack.count > 1
    mnuEditRedo.Enabled = RedoStack.count > 0
    With frmMain.tbEdit
        .Buttons("cut").Enabled = (RTF1.SelLength <> 0)
        .Buttons("copy").Enabled = (RTF1.SelLength <> 0)
        .Buttons("paste").Enabled = (Clipboard.GetFormat(vbCFText) = True)
        .Buttons("undo").Enabled = mnuEditUndo.Enabled
        .Buttons("redo").Enabled = mnuEditRedo.Enabled
    End With
End Sub


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



Public Sub printText()
    
    On Error GoTo ErrorHandler
    
    'setup header and footer
    sPrintText = RTF1.Text
    sHeader = SetPrintLine(sPrintHeader)
    sFooter = SetPrintLine(sPrintFooter)
    sPrintText = sHeader & sPrintText & sFooter
    Me.rtfTmp.Text = sPrintText
    
    ' This is where the printing is called
    On Error GoTo ErrorHandler
    frmMain.CD.Flags = cdlPDReturnDC + cdlPDNoPageNums
        
    If RTF1.SelLength = 0 Then
        frmMain.CD.Flags = frmMain.CD.Flags + cdlPDAllPages
    Else
        frmMain.CD.Flags = frmMain.CD.Flags + cdlPDSelection
    End If
    frmMain.CD.ShowPrinter
    ' Printing with margin at all four sides.
    ' To use the PrintRTF function we must send it margins in TWIPS. Since the
    ' pagesetup form uses millimeters we must convert them to twips first.
    ' There is aproximatly 57 TWIPS in 1 millimeter.
    PrintRTF rtfTmp, (gLeftMargin * 57), (gTopMargin * 57), (gRightMargin * 57), (gBottomMargin * 57)

    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Sub ParseDate()
On Error GoTo hell
Dim pos As Long, pos2 As Long, ss As Long
Dim temp As String
Dim actual As String
ss = RTF1.SelStart
pos = InStr(1, RTF1.Text, "<!-- WHTMLDATE ")
pos2 = InStr(pos + 1, RTF1.Text, "//-->")
If pos2 = 0 And pos <> 0 Then MsgBox "Invalid Date Script in the document." & vbCrLf & "Cannot update Date/Time.", vbExclamation
If pos = 0 Or pos2 = 0 Then Exit Sub
temp = Mid$(RTF1.Text, pos + 15, pos2 - pos - 15) '20=len(whtmldate...etc)
pos = pos2 + 5
pos2 = InStr(pos, RTF1.Text, "<!-- ENDWHTMLDATE //-->")
If pos2 = 0 Then MsgBox "Invalid Date Script in the document." & vbCrLf & "Cannot update Date/Time.", vbExclamation: Exit Sub
temp = Trim(temp)
actual = Format(Now, temp)
RTF1.SelStart = pos - 1
RTF1.SelLength = pos2 - pos
RTF1.SelText = actual
frmMain.SB.Panels(1).Text = "Date and time in document changed to " & actual
RTF1.SelStart = ss
Exit Sub
hell:
MsgBox "Invalid Date Script in the document." & vbCrLf & "Cannot update Date/Time.", vbExclamation
RTF1.SelStart = ss
End Sub

Sub EnableFindDialog()
On Error Resume Next
If frmFind Is Nothing Then Exit Sub
With frmFind
.SSTab1.Tab = 2
.Label1(1).Enabled = True
.Label1(2).Enabled = True
.Label1(3).Enabled = True
.txF.Enabled = True
.txR.Enabled = True
.chCase.Enabled = True
.chWhole.Enabled = True
.txFindFiles.TabIndex = 0
.cmFind.Enabled = True
.cmRepAll.Enabled = True
.cmRepThis.Enabled = True
.lbG.Enabled = True
.txG.Enabled = True
.cbAB.Enabled = True
.cbAbsRel.Enabled = True
.chCloseGo.Enabled = True
.cmG.Enabled = True
.opG(0).Enabled = True
.opG(1).Enabled = True
.lbDetPos.Enabled = True
.opG(2).Enabled = True
.Label6.Enabled = True
.SSTab1_Click 0
.cbAbsRel_Click
End With
End Sub

Function GetLineText() As String
On Error Resume Next
Dim line As Long, lngStart As Long
Dim start As Long
line = GetCurrentLine(RTF1)
lngStart = SendMessage(RTF1.hWnd, EM_LINEINDEX, line - 1, 0&)
start = lngStart
line = line + 1
lngStart = SendMessage(RTF1.hWnd, EM_LINEINDEX, line - 1, 0&)
If lngStart = -1 Then lngStart = Len(RTF1.Text) + 2
GetLineText = Mid$(RTF1.Text, start + 1, lngStart - start - 2)
End Function

Sub AutoIndent()
Dim s As String, l As Long
s = GetLineText
l = Len(s) - Len(LTrim(s))
RTF1.SelText = vbCrLf & Space(l)
End Sub

Private Sub tmrAutoSave_Timer()
If bChanged = False Then Exit Sub
If ReadValue("AutoSave", True) = False Then Exit Sub
Open Caption & ".wbackup" For Output As #1
'if untitled, it's saved in the working dir.
Print #1, RTF1.Text
Close #1
End Sub

Sub GotoLineProc()
On Error Resume Next
Load frmFind
frmFind.SSTab1.Tab = 3
frmFind.txG.TabIndex = 0
frmFind.Show
End Sub

