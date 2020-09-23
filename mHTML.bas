Attribute VB_Name = "mHTML"
'######################################
'WonderHTML 0.90 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
Option Explicit

Public Const EM_GETFIRSTVISIBLELINE = &HCE
' The in-memory database of tips.
Public Tips As New Collection
' Name of tips file
Public Const TIP_FILE = "whtml.tip"
'Today is 1 October, 2001; 4:15 PM IST. I accidentally
'deleted all functions in this module and saved it. Just
'DirDiver was left. Had to write every one again!
'-----------------------------------------------------------------
'Today is October 12. Nearly had a reprive. I had
'a folder with all my songs in E:\. Accidentally moved it.
'I somehow thought it got copied, so deleted
'the files at the destination. Of course,
'nothing was there at the original location.
'I answered Yes to all when the dumb thing asked
'if I want to delete it permanently cause it's too large for the bin.
'Lost over 70+ wonder songs that I d/l'ed with much effort.
'-------------------------------------------------------------------
'Smile.
'-------------------------------------------------------------------
'Today I'm going to start commenting the code. (15-Oct-2001)
'-------------------------------------------------------------------
'Why do programmers get confused between halloween and christmas?
'cause Dec 25 = Oct 31.
'-------------------------------------------------------------------
'Why do Pascal programmers live in Alaska?
'cause it's above C level.
'------------------------------------------------------------------
'http://sushantshome.tripod.com/laugh/index.html
'------------------------------------------------------------------
'OK, Enough. Code.
'------------------------------------------------------------------

Function SelectDir(Optional NoName As Boolean, Optional Height As Long) As String
'bad function
Load frmWeb
frmWeb.txName.Enabled = Not NoName
frmWeb.lbTemplate.Visible = Not NoName
frmWeb.Label4.Enabled = Not NoName
frmWeb.tvWebs.Enabled = Not NoName
frmWeb.tvWebs.Visible = Not NoName
frmWeb.txName.Visible = frmWeb.txName.Enabled
frmWeb.Label4.Visible = frmWeb.Label4.Enabled
'frmWeb.Height = Height
frmWeb.Show vbModal
SelectDir = ReturnedPath
End Function

Function IsContained(Text As String, list As Object) As Boolean
'if the text is an item in the listbox
IsContained = False
Dim i As Integer
For i = 0 To list.ListCount - 1
    If list.list(i) = Text Then IsContained = True: Exit Function
Next i
End Function

Sub SetFont(F As Form)
'set the damn fonts
On Error Resume Next
Dim lpF As Control, fnt As String
fnt = ReadValue("DisplayFont", "Tahoma")
For Each lpF In F.Controls
If F.Name = "frmDocPreview" And Left(lpF.Name, 3) = "Pic" Then GoTo n
    lpF.Font.Name = fnt
    If lpF.Font.Name <> fnt Then lpF.Font.Name = "MS Sans Serif"
n:
SaveValue lpF.Name, lpF.Caption, F.Name, "C:\default.lng"
Next lpF
End Sub

Public Function ReadValue(Key As String, Optional Default As String, Optional Section As String = "WonderHTML", Optional File)
    ' Read from INI file
    Dim sReturn As String
    If IsMissing(File) Then File = FullPath(App.Path, "settings.ini")
    sReturn = String(255, Chr(0))
    ReadValue = Left(sReturn, GetPrivateProfileString(Section, Key, Default, sReturn, Len(sReturn), File))
End Function

Public Sub SaveValue(Key As String, Value As String, Optional Section As String = "WonderHTML", Optional File)
    ' Write to INI file
    If IsMissing(File) Then File = FullPath(App.Path, "settings.ini")
    WritePrivateProfileString Section, Key, Value, File
End Sub

Function FormsLeft() As Long
'number of frmChilds left
Dim tmp As Long
Dim lpF As Form
For Each lpF In Forms
If lpF.Name = "frmChild" Then tmp = tmp + 1
Next lpF
FormsLeft = tmp
End Function

Function FullPath(lpPath As String, lpFile As String) As String
If Right(lpPath, 1) <> "\" Then lpPath = lpPath & "\"
FullPath = lpPath & lpFile
'fullpath, after resolving the "\" problems
End Function

Sub LoadImage(Path As String)
On Error GoTo n
Dim frmImag As New frmImage
Load frmImag
frmImag.Tag = Path
frmImag.Caption = Path
frmImag.Form_Resize
frmImag.PB.Picture = LoadPicture(Path)
frmImag.Caption = Path '& " (" & Format(FileLen(Path), "0,000") & " Bytes, " & frmImag.pB.Width & "x" & frmImag.pB.Height & ")"
Exit Sub
n:
MsgBox Error, vbExclamation
End Sub

Public Sub SetViewMode(ByVal eViewMode As ERECViewModes, RTF As RichTextBox)
 Select Case eViewMode 'Set View Mode
 Case 0 'to No Wrap
 SendMessageLong RTF.hwnd, EM_SETTARGETDEVICE, 0, 1
 Case 1 'to Word Wrap
 SendMessageLong RTF.hwnd, EM_SETTARGETDEVICE, 0, 0
 Case 2 'to WYSIWYG
 On Error Resume Next
 SendMessageLong RTF.hwnd, EM_SETTARGETDEVICE, Printer.hdc, Printer.Width
 End Select
End Sub

Function GetFirstVisible(RTF As RichTextBox)
Dim l As Long
l = SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
GetFirstVisible = l
'first visible line
End Function

Sub GotoLine(RTF As RichTextBox, LineNum As Long)
On Error Resume Next
Dim lngStart As Long
lngStart = SendMessage(RTF.hwnd, EM_LINEINDEX, LineNum - 1, 0&)
RTF.SelStart = lngStart 'Go To line
End Sub

Public Function RichWordOver(rch As RichTextBox, X As Single, y As Single, Optional start) As String
'find word mouse is over
Dim pt As POINTAPI
Dim pos As Long
Dim start_pos As Long
Dim end_pos As Long
Dim ch As String
Dim txt As String
Dim txtlen As Long

    ' Convert the position to pixels.
    pt.X = X \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(rch.hwnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function

    ' Find the start of the word.
    txt = rch.Text
    For start_pos = pos To 1 Step -1
        ch = Mid$(rch.Text, start_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or ch = "_" Or ch = ":" Or ch = "/" Or ch = "." Or ch = "@" Or ch = "&" Or ch = "+" Or ch = "-") Then Exit For
    Next start_pos
    start_pos = start_pos + 1
    If Not IsMissing(start) Then start = start_pos
    ' Find the end of the word.
    txtlen = Len(txt)
    For end_pos = pos To txtlen
        ch = Mid$(txt, end_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or ch = "_" Or ch = ":" Or ch = "/" Or ch = "." Or ch = "@" Or ch = "&" Or ch = "+" Or ch = "-") Then Exit For
    Next end_pos
    end_pos = end_pos - 1

    If start_pos <= end_pos Then _
        RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1)
End Function

Public Function GetTotalLines(RichTextBox As RichTextBox)
'just what it says
    Dim TotalLines As Long
    TotalLines = SendMessage(RichTextBox.hwnd, EM_GETLINECOUNT, 0, 0&)
    GetTotalLines = TotalLines
End Function

Public Function GetCurrentLine(RichTextBox As RichTextBox)
'just what it says
    Dim CurrentLine As Long
    CurrentLine = SendMessage(RichTextBox.hwnd, EM_LINEFROMCHAR, -1, 0&) + 1
    GetCurrentLine = CurrentLine
End Function

Function GetColumnIndex(RTF As RichTextBox) As Long
'column where caret is
Dim i As Long
i = SendMessage(RTF.hwnd, EM_LINEINDEX, ByVal GetCurrentLine(RTF) - 1, 0&)
GetColumnIndex = RTF.SelStart + 1 - i
End Function

Function GetDay(whatdate)
'I'm given to understand Format(Date,"dddd") works the same
Dim lDay As Long
lDay = Weekday(Date)
Select Case lDay
Case 1
GetDay = "Sunday"
Case 2
GetDay = "Monday"
Case 3
GetDay = "Tuesday"
Case 4
GetDay = "Wednesday"
Case 5
GetDay = "Thursday"
Case 6
GetDay = "Friday"
Case 7
GetDay = "Saturday"
End Select
End Function

Sub Main()
On Error Resume Next

If Dir(FullPath(App.Path, "whtml32.exe")) = "" Then
MsgBox "You're probably running this from within the IDE. Bugs may be encountered, such as icons and the toolbar in the report window not displaying properly. In that case, just open the problematic form in the designer window before running. It all works OK. In fact, when you compile it, such visual problems will not occur at all." & vbCrLf & vbCrLf & "© 2001, Sushant S. Pandurangi. http://sushantshome.tripod.com/vb/wonder.html", vbExclamation
End If

Set frmMain = New frmMDI
If ReadValue("NoSplash", False) = True Then GoTo nosplash
Load frmSplash
frmSplash.tUnload.Enabled = True
Load frmMain 'so it keeps loading while we show splash screen
frmSplash.Show vbModal
nosplash:
frmMain.Show
If ReadValue("StartupTips", 1) = 1 Then frmTip.Show vbModal
End Sub

Function GetMonth(whatdate)
'same as GetDay
Dim lDay As Long
lDay = (whatdate)
Select Case lDay
Case 1
GetMonth = "January"
Case 2
GetMonth = "February"
Case 3
GetMonth = "March"
Case 4
GetMonth = "April"
Case 5
GetMonth = "May"
Case 6
GetMonth = "June"
Case 7
GetMonth = "July"
Case 8
GetMonth = "August"
Case 9
GetMonth = "September"
Case 10
GetMonth = "October"
Case 11
GetMonth = "November"
Case 12
GetMonth = "December"
End Select
End Function

Function FirstSlash(Text As String) As String
If Left(Text, 1) = "\" Or Left(Text, 1) = "/" Then FirstSlash = Right(Text, Len(Text) - 1) Else FirstSlash = Text
'remove first slash from "/text" or "\text"
End Function

Sub FileInfo(Filename As String, Optional TaskView As Boolean = False)
'load the frmfile and show info
Load frmFile
frmFile.Tag = Filename
If TaskView Then frmFile.SSTab1.Tab = 1
frmFile.Proceed 'do the stuff
frmFile.Show vbModal
End Sub

Function HTML()
'default HTML to insert in new doc
Dim s As String
s = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">" & vbCrLf & vbCrLf
s = s & "<HTML>" & vbCrLf & "<HEAD>" & vbCrLf & "<META name=" & Chr(34) & "Author" & Chr(34) & " content=" & Chr(34) & ReadValue("Author", "WonderUser", "Documents") & Chr(34) & ">" & vbCrLf
s = s & "<META name=" & Chr(34) & "Generator" & Chr(34) & " content=" & Chr(34) & "WonderHTML 0.90" & Chr(34) & ">" & vbCrLf
s = s & "<TITLE>HTML Document</TITLE>" & vbCrLf & ReadValue("Comments", , "Documents") & vbCrLf & "</HEAD>"
s = s & vbCrLf & "<BODY" & IIf(ReadValue("BodyAttrib", , "Documents") = "", "", " ") & ReadValue("BodyAttrib", , "Documents") & ">" & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & "</HTML>"
HTML = s
End Function

Function FileType(File As String) As String
Dim s As String
s = GetValue(HKEY_CLASSES_ROOT, "." & Ext(File), "", "")
'get the "(Default)" value from the extension
s = Trim(s)
If Len(s) = 0 Then s = UCase(Ext(File)) & " File": Exit Function
s = GetValue(HKEY_CLASSES_ROOT, s, "", "")
'get the (Default) value from the ID. mostly it is ext+file, so for ".exe" it is "exefile".
s = Trim(s)
If Len(s) = 0 Then s = UCase(Ext(File)) & " File": Exit Function
FileType = s
End Function

Function Ext(File As String) As String
'extension only
On Error Resume Next
Dim i As Long
i = InStr(StrReverse(File), ".")
If i = 0 Then Ext = File: Exit Function
Ext = Right(File, i - 1)
End Function

Sub CleanUp(RTF As RichTextBox, HR As Boolean)
'yep - clean up
'This is for the convert to ASCII feature
On Error Resume Next
Dim pos As Long, pos2 As Long, ListNum As Long
Dim sText As String, bInOL As Boolean, bInUL As Boolean
Do
pos = InStr(1, RTF.Text, "<")
'find first <
pos2 = InStr(pos + 1, RTF.Text, ">")
'find following >
If pos = 0 Then Exit Do
If pos2 = 0 Then Exit Do
If pos2 < pos Then Exit Do
RTF.SelStart = pos - 1
RTF.SelLength = pos2 - pos + 1
If InStr(RTF.SelText, "!--") > 0 Then sText = "": GoTo n
Select Case NoTags(LCase(Tag(RTF.SelText))) 'select what to do
Case "/title"
sText = vbNewLine & vbNewLine
Case "hr"
If HR Then sText = String(30, "-") Else sText = vbNewLine
Case "/dt", "/tr", "br"
sText = vbNewLine
Case "ul"
bInUL = True
sText = ""
Case "/ul"
bInUL = False
sText = vbNewLine
Case "ol"
bInOL = True
sText = ""
Case "/ol"
bInOL = False
sText = vbNewLine
ListNum = 0
Case "li"
If bInOL Then
    ListNum = ListNum + 1
    sText = vbCrLf & ListNum & ". "
ElseIf bInUL Then
    sText = vbCrLf & "¤ "
End If
Case "/li"
sText = vbNewLine
Case "p"
sText = vbNewLine & vbNewLine
Case Else
sText = ""
End Select
n:
RTF.SelText = sText 'replace the tag
nLoop:
Loop
End Sub

Sub ConvEntities(RTF As RichTextBox)
'another for the convert to ASCII thing
'find the &#xxx; entities, and convert them to their char
On Error Resume Next
Dim i As Long, i2 As Long
Do
i = InStr(i2 + 1, RTF.Text, "&#")
If i = 0 Then Exit Do
i2 = InStr(i + 1, RTF.Text, ";")
If i = 0 Then Exit Do
RTF.SelStart = i - 1
RTF.SelLength = i2 - i + 1
RTF.SelText = Chr(CLng(Mid(RTF.Text, i + 2, i2 - i - 2)))
n:
Loop
End Sub

Sub ReplaceStuff(RTF As RichTextBox)
RTF.Text = Replace(RTF.Text, "&amp;", "&")
RTF.Text = Replace(RTF.Text, "&copy;", "©")
RTF.Text = Replace(RTF.Text, "&quot;", Chr(34))
RTF.Text = Replace(RTF.Text, "&reg;", "®")
RTF.Text = Replace(RTF.Text, "&nbsp;", " ")
RTF.Text = Replace(RTF.Text, "&lt;", "<")
RTF.Text = Replace(RTF.Text, "&gt;", ">")
End Sub

Function Tag(sTag As String) As String
'tag name only
'e.g. <BODY a="b" c="d"> returns BODY
Dim l As Long
l = InStr(sTag, " ")
If l = 0 Then Tag = sTag Else Tag = Left(sTag, l)
Tag = Trim(Tag)
End Function

Function NoTags(Tag As String)
'no < and >
NoTags = Replace(Tag, "<", "")
NoTags = Replace(NoTags, ">", "")
End Function


Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    Set Tips = New Collection
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    
    LoadTips = True
    
End Function

Sub ShowHelp(File As String)
On Error Resume Next
Load frmHelp
frmHelp.ShowTopic File
frmHelp.sTree.SelectedItem = frmHelp.sTree.Nodes(File)
frmHelp.Show vbModal
End Sub

Function IsLink(Text As String) As Boolean
Dim i As Integer, Words() As String
Words = Split(Domains, " ")
For i = 0 To UBound(Words())
If InStr(1, Text, Words(i)) > 0 Then IsLink = True: Exit Function
Next i
IsLink = False 'not found anywhere
End Function

Function GetMoreInfo(Text As String) As String
If InStr(1, Text, "@") Then
GetMoreInfo = "mailto:"
ElseIf InStr(1, Text, "http://") Then
GetMoreInfo = ""
ElseIf InStr(1, Text, "://") = 0 Then
GetMoreInfo = "http://"
Else
GetMoreInfo = ""
End If
End Function

Function GetSizeLimit() As Long
GetSizeLimit = CLng(ReadValue("SizeLimit", 0, "Reports")) * 1024
End Function
