Attribute VB_Name = "mMRU"
'######################################
'WonderHTML 1.2 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
Option Explicit
Dim i As Integer
Public assoc_text As String

Sub ShowWebMRU()
    For i = 1 To 4
        If i > frmMain.WebMRU.count Then Exit For
        ' Set menu caption
        frmMain.mnuWebMRU(i).Caption = "&" & i & "  " & GetFile(frmMain.WebMRU(i))
        ' Set menu tag to file name
        frmMain.mnuWebMRU(i).Tag = frmMain.WebMRU(i)
        ' Show menu
        frmMain.mnuWebMRU(i).Visible = True
    Next
    
    For i = frmMain.WebMRU.count + 1 To 4
        frmMain.mnuWebMRU(i).Visible = False 'Hide empty menus
    Next i
    

End Sub

Sub GetWebMRU()
    Dim Filename As String
    
    Set frmMain.WebMRU = New Collection 'Create new collection
    For i = 1 To 4
        Filename = ReadValue("WebMRU" & i, , "MRU Webs")
        If Len(Filename) > 2 Then
            frmMain.WebMRU.Add Filename  'Add file name to collection
        End If
    Next
    ShowWebMRU 'Call DisplayMRUList sub

End Sub

Sub AddWebMRU(Filename As String)

    For i = 1 To 4
        If i > frmMain.WebMRU.count Then Exit For
        If LCase(frmMain.WebMRU(i)) = LCase(Filename) Then 'If filename exist in the
            frmMain.WebMRU.Remove i                     'collection exit sub
            Exit For
        End If
    Next i
    
    If frmMain.WebMRU.count > 0 Then 'If the collection is not empty
        frmMain.WebMRU.Add Filename, , 1  'add file to begining of the collecton
    Else 'else
        frmMain.WebMRU.Add Filename  'just add it
    End If
    
    If frmMain.WebMRU.count > 4 Then 'If there are more items than 8 remove the last one
        frmMain.WebMRU.Remove 5
    End If
    
    For i = 1 To 4
        If i > frmMain.WebMRU.count Then 'If no more files then leave it empty
            Filename = ""
        Else 'else
            Filename = frmMain.WebMRU(i) 'add it
        End If
        ' Add file to the INI
        SaveValue "WebMRU" & i, Filename, "MRU Webs"
    Next i
    GetWebMRU

End Sub

Public Sub GetFileMRU()
    Dim i As Integer
    Dim Filename As String
    
    Set frmMain.FileMRU = New Collection 'Create new collection
    For i = 1 To 6
        Filename = ReadValue("FileMRU" & i, , "MRU Files")
        If Len(Filename) > 2 Then
            frmMain.FileMRU.Add Filename 'Add file name to collection
        End If
    Next
    ShowFileMRU 'Call DisplayMRUList sub
End Sub

Public Sub ShowFileMRU()
On Error Resume Next
    Dim i As Integer
    For i = 1 To 6
        If i > frmMain.FileMRU.count Then Exit For
        ' Set menu caption
        frmMain.mnuFileMRU(i).Caption = "&" & i & "  " & GetFile(frmMain.FileMRU(i))
        ' Set menu tag to file name
        frmMain.mnuFileMRU(i).Tag = frmMain.FileMRU(i)
        ' Show menu
        frmMain.mnuFileMRU(i).Visible = True
    Next
    
    For i = frmMain.FileMRU.count + 1 To 6
        frmMain.mnuFileMRU(i).Visible = False 'Hide empty menus
    Next
    
    frmMain.LoadToolbarMRU
    
End Sub

Public Sub AddFileMRU(ByVal Filename As String)
If Ext(Filename) = "wbackup" Then Exit Sub
    Dim i As Integer

    For i = 1 To 6
        If i > frmMain.FileMRU.count Then Exit For
        If LCase(frmMain.FileMRU(i)) = LCase(Filename) Then 'If filename exist in the
            frmMain.FileMRU.Remove i                     'collection exit sub
            Exit For
        End If
    Next i
    
    If frmMain.FileMRU.count > 0 Then 'If the collection is not empty
        frmMain.FileMRU.Add Filename, , 1  'add file to begining of the collecton
    Else 'else
        frmMain.FileMRU.Add Filename 'just add it
    End If
    
    If frmMain.FileMRU.count > 6 Then 'If there are more items than 8 remove the last one
        frmMain.FileMRU.Remove 7
    End If
    
    For i = 1 To 6
        If i > frmMain.FileMRU.count Then 'If no more files then leave it empty
            Filename = ""
        Else 'else
            Filename = frmMain.FileMRU(i) 'add it
        End If
        ' Add file to the registry
        SaveValue "FileMRU" & i, Filename, "MRU Files"
    Next i
    GetFileMRU
End Sub

Sub Outline(lpForm As Form, lpTreeViewCtl As TreeView, lpStringHTML As String, bExitFlags As Boolean)
On Error Resume Next  'just in case
If bExitFlags Then Exit Sub
frmMain.MousePointer = 11: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(2).Picture
Dim InPos As Long, InPos2 As Long, ThisTag As String
Dim SpacePos As Long, sImg As Long, whole As String
Dim pParent As String, InScript As Boolean
InScript = False

lpTreeViewCtl.Nodes.Clear 'I tried using WM_CLEAR, but it hangs
If lpForm.Icon = lpForm.p2.Picture Then GoTo bye: Exit Sub 'its not an HTML doc

AddMainNodes lpTreeViewCtl, lpForm.Caption

Do

InPos = InStr(InPos2 + 1, lpStringHTML, "<")
InPos2 = InStr(InPos + 1, lpStringHTML, ">")
If InPos = 0 Then Exit Do
If InPos2 = 0 Then Exit Do
'find < and > positions
ThisTag = Mid$(lpStringHTML, InPos + 1, InPos2 - 1 - InPos)
'therefore find the tag text
whole = ThisTag 'at this time, whole tag and thistag are same
SpacePos = InStr(1, ThisTag, " ")
If SpacePos > 0 Then ThisTag = Left(ThisTag, SpacePos)
ThisTag = Trim(ThisTag) 'tag is now only the tag ID
'e.g. <BODY fpp="a" bar="b"> means BODY

'inscript helps to determine if we're in <!-- //-> or not

If Left(ThisTag, 2) = "!-" Then GoTo comments 'comments
If LCase(ThisTag) = "/script" Then InScript = False 'we're out of it
If LCase(ThisTag) = "/style" Then InScript = False 'we're out of it
If Left(ThisTag, 1) = "/" Then GoTo n 'don't want end tags
If Left(ThisTag, 1) = " " Then GoTo n 'is a bogus tag

'now handle the text, set which parent to put under, and what to name it
Select Case LCase(ThisTag)
    Case "body", "head"
        pParent = "document"
        ThisTag = UCase(ThisTag)
    Case "img", "bgsound"
        pParent = "images"
        ThisTag = ReadAttrib("src", whole)
    Case "a"
        pParent = "links"
        ThisTag = ReadAttrib("href", whole)
        If ThisTag = whole Then
            ThisTag = ReadAttrib("name", whole)
            pParent = "bookmarks" 'put in bookmarks
        End If
        If ThisTag = "" Or ThisTag = "#" Then GoTo n
    Case "table"
        pParent = "tables"
    Case "applet"
        pParent = "applets"
        ThisTag = ReadAttrib("code", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("codebase", whole)
    Case "form", "input", "select", "textarea"
        pParent = "forms"
        ThisTag = ReadAttrib("name", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("id", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("value", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("type", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("action", whole)
    Case "font"
        pParent = "styles"
        ThisTag = ReadAttrib("face", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("color", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("size", whole)
        If ThisTag = whole Then ThisTag = "FONT"
    Case "div", "span"
        pParent = "divisions"
        ThisTag = ReadAttrib("id", whole)
        If ThisTag = whole Then ThisTag = "DIV"
    Case "script"
        pParent = "scripts"
        ThisTag = ReadAttrib("src", whole)
        InScript = True
        If ThisTag = whole Then ThisTag = ReadAttrib("language", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("id", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("type", whole)
    Case "object"
        pParent = "Objects"
        ThisTag = ReadAttrib("code", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("classid", whole)
        If ThisTag = whole Then ThisTag = "OBJECT"
    Case "embed"
        pParent = "Plugins"
        ThisTag = ReadAttrib("src", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("type", whole)
        If ThisTag = whole Then ThisTag = "EMBED"
    Case "layer"
        pParent = "layers"
        ThisTag = ReadAttrib("name", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("id", whole)
    Case "title"
        Dim i As Integer, T As String
        i = InStr(InPos2 + 1, lpStringHTML, "</")
        T = Mid$(lpStringHTML, InPos, i - InPos + 9)  '9 len of </title>
        T = Trim(T)
        whole = Mid(T, 2, Len(T) - 3)
        ThisTag = Mid$(whole, 7, Len(whole) - 13) '7 after <title> and 13 counting </title>
        lpTreeViewCtl.Nodes(1).Text = ThisTag & " (" & GetFile(lpTreeViewCtl.Nodes(1).Text) & ")"
        GoTo n
    Case "h1", "h2", "h3", "h4", "h5", "h6"
        pParent = "headings"
        ThisTag = UCase(ThisTag)
    Case "ol", "ul", "d", "dt"
        pParent = "lists"
        ThisTag = UCase(ThisTag)
    Case "p", "br", "li", "html", "b", "i", "u", "em", "blockquote", _
    "", "tr", "td", "th", "center", "param"
        GoTo n 'don't want
    Case "meta", "!doctype"
        pParent = "declare"
        ThisTag = ReadAttrib("name", whole)
        If ThisTag = whole Then ThisTag = ReadAttrib("http-equiv", whole)
        If ThisTag = whole And Left(ThisTag, 1) = "!" Then ThisTag = "!DOCTYPE"
    Case "!--"
comments:
        If InScript Then GoTo n 'inscript helps to not add comments which
        'are inside script or style tags
        pParent = "comments"
        ThisTag = Mid(whole, 4, Len(whole) - 6): ThisTag = Trim(ThisTag)
        If Left(ThisTag, 9) = "WHTMLDATE" Or Left(ThisTag, 12) = "ENDWHTMLDATE" Then GoTo n
        If Len(ThisTag) > 9 Then ThisTag = Left(ThisTag, 9) & "..."
    Case "style"
        InScript = True
        pParent = "other"
    Case Else
        pParent = "other"
End Select
lpTreeViewCtl.Nodes.Add pParent, tvwChild, "temp", ThisTag, pParent '"tag"
lpTreeViewCtl.Nodes.Item("temp").Tag = whole
lpTreeViewCtl.Nodes.Item("temp").Key = ""
'lpTreeViewCtl.Nodes(pParent).Text = GetFunctionName(lpTreeViewCtl.Nodes(pParent).Text) & " (" & lpTreeViewCtl.Nodes(pParent).Children & ")"
n: 'next
Loop
lpTreeViewCtl.Nodes.Item("Main").Expanded = True
lpTreeViewCtl.SelectedItem = lpTreeViewCtl.Nodes.Item("Main")
bye:
frmMain.MousePointer = 0: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
End Sub

Sub AddMainNodes(lpTreeViewCtl As TreeView, Filename As String)
On Error Resume Next
lpTreeViewCtl.Nodes.Add , , "Main", Filename, "root"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "bookmarks", "Anchors", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "applets", "Applets", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "comments", "Comments", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "declare", "Declarations", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "divisions", "Divisions", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "forms", "Forms", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "headings", "Headings", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "images", "Pictures", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "layers", "Layers", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "links", "Hyperlinks", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "lists", "Lists", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "Objects", "Objects", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "Plugins", "Plugins", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "scripts", "Scripts", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "styles", "Styles", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "tables", "Tables", "category"
lpTreeViewCtl.Nodes.Add "Main", tvwChild, "other", "Others", "category"

End Sub

Function ReadAttrib(ByVal lpAttribName As String, ByVal lpString As String) As String
'readattrib("name","<BODY name=foo>") returns foo
On Error Resume Next
Dim lnPos1 As Long, lnPos2 As Long, tmp As String

  lnPos1 = InStr(1, lpString, " " & lpAttribName & "=", vbTextCompare)
  If lnPos1 > 0 Then lnPos1 = lnPos1 + Len(lpAttribName) + 2 Else ReadAttrib = lpString: Exit Function
  
  lnPos2 = InStr(lnPos1 + 1, lpString, " ")
  If Mid$(lpString, lnPos1, 1) = "'" Then lnPos2 = InStr(lnPos1 + 1, lpString, "'")
  If Mid$(lpString, lnPos1, 1) = Chr(34) Then lnPos2 = InStr(lnPos1 + 1, lpString, Chr(34))
  If lnPos2 = 0 Then lnPos2 = Len(lpString) + 1
  tmp = Mid$(lpString, lnPos1, lnPos2 - lnPos1)
  
  If Left(tmp, 1) = Chr(34) Or Left(tmp, 1) = "'" Then tmp = Right(tmp, Len(tmp) - 1)
  If Right(tmp, 1) = Chr(34) Or Right(tmp, 1) = "'" Then tmp = Left(tmp, Len(tmp) - 1)
  
  tmp = Replace(tmp, "\'", "'")
  tmp = Replace(tmp, "\" & Chr(34), Chr(34))
  
  ReadAttrib = tmp
End Function

Public Function ParseInt(Expression As Variant) As Long
'This will return only the Integer portion
Dim pos As Integer, temp As String
For pos = 1 To Len(Expression)
If IsNumeric(Mid(Expression, pos, 1)) = True Then temp = temp & CStr(Mid(Expression, pos, 1))
Next pos
ParseInt = CLng(temp)
End Function

Sub AddScripts(Text As String, lpTree As TreeView, bExitFlags As Boolean)

'14-10-01
'WONDERFULLY UPDATED!

On Error Resume Next
If bExitFlags Then Exit Sub

Dim InPos1 As Long, InPos2 As Long, InPosWhole As Long, i As Long
Dim VarCount, FunCount As Long, lpText As String
Dim ThisFunName As String, ThisFunBody  As String

lpText = Text 'so original is not altered. ByVal didn't work!
lpText = NoStrings(lpText) 'dim out strings, don't want
lpText = NoComments(lpText) 'same for comments // and /* */
lpTree.Nodes.Clear

lpTree.Nodes.Add , , "globals", "Client Scripts", "folder"
lpTree.Nodes.Add , , "server", "Server Objects", "folder"
lpTree.Nodes.Add , , , "Server Scripts", "folder"
''''''''''currently only client scripts displayed.''''''''''''

lpTree.Nodes.Add "server", tvwChild, , "Application", "obj"
lpTree.Nodes.Add "server", tvwChild, , "Request", "obj"
lpTree.Nodes.Add "server", tvwChild, , "Response", "obj"
lpTree.Nodes.Add "server", tvwChild, , "ScriptingContext", "obj"
lpTree.Nodes.Add "server", tvwChild, , "Server", "obj"
lpTree.Nodes.Add "server", tvwChild, , "Session", "obj"
'''''''above ones not implemented yet (14-nov '01)'''''''

Do
InPos1 = InStr(InPos2 + 1, lpText, "function ", vbBinaryCompare)
If InPos1 = 0 Then Exit Do
InPos2 = InStr(InPos1 + 1, lpText, "{")
If InPos2 = 0 Then Exit Do
If InStr(InPos1 + 1, lpText, vbNewLine) < InPos2 Then InPos2 = InStr(InPos1 + 1, lpText, vbNewLine)
ThisFunName = Mid$(lpText, InPos1 + 9, InPos2 - InPos1 - 9)
InPosWhole = InStr(InPos1 + 1, lpText, "}") 'last } pos
'now if function "body" at this stage contains a {, it is a part of the body.
'To determine the entire body, seek upto the next }.
If StrCount(Mid$(lpText, InPos1, InPosWhole - InPos1), "{") = 0 Then GoTo n
  For i = 1 To StrCount(Mid$(lpText, InPos1, InPosWhole - InPos1), "{") - 1
    InPosWhole = InStr(InPosWhole + 1, lpText, "}") 'increase by one
  Next i
n:
ThisFunBody = Mid$(lpText, InPos1, InPosWhole - InPos1 + 1) 'whole body
Mid$(lpText, InPos1, InPosWhole - InPos1) = Space(Len(ThisFunBody))
If InStr(ThisFunName, "(") = 0 Or InStr(ThisFunName, ")") = 0 Then GoTo nxt
lpTree.Nodes.Add "globals", tvwChild, "function " & ThisFunName, GetFunctionName(ThisFunName), "function"
ProcessVariables ThisFunBody, lpTree, "function " & ThisFunName 'add vars
nxt:
Loop

ProcessVariables lpText, lpTree, "globals" 'add global vars
lpTree.Nodes("globals").Expanded = True
lpTree.Nodes("server").Expanded = True

End Sub

Sub ProcessVariables(Text As String, Tree As TreeView, parent As String)
On Error Resume Next
'add variables from text
Dim l1 As Long, l2 As Long
Dim thisVar As String
Do
l1 = InStr(l2 + 1, Text, "var ")
If l1 = 0 Then Exit Do
l2 = InStr(l1 + 1, Text, ";")
If InStr(l1 + 1, Text, "=") < l2 And l2 <> 0 Then l2 = InStr(l1 + 1, Text, "=")
If InStr(l1 + 1, Text, vbNewLine) < l2 And l2 <> 0 Then l2 = InStr(l1 + 1, Text, vbNewLine)
If l2 = 0 Then Exit Do
thisVar = Mid$(Text, l1 + 4, l2 - l1 - 4)
Tree.Nodes.Add parent, tvwChild, "var " & thisVar & IIf(parent = "globals", "", ": local to " & parent), thisVar, "var"
Loop
End Sub

Function FileIcon(lpFileName As String) As String
'matching icon
Select Case LCase(Right(lpFileName, 3))
Case "tml", "htm", "xml", "asp"
FileIcon = "file"
Case "gif", "bmp", "jpg", "wmf", "png", "ico"
FileIcon = "image"
Case ".js", "vbs"
FileIcon = "script"
Case "wav", "mp3", "ram", "ra", "rm", "mid", "rmi"
FileIcon = "audio"
Case "exe", "bin"
FileIcon = "program"
Case "cgi", "pl", "bat"
FileIcon = "shellscript"
Case "doc", "rtf"
FileIcon = "winword"
Case "zip", "arj", "rar", "gz", "sit", "hqx"
FileIcon = "archive"
Case "pdf", "psd"
FileIcon = LCase(Right(lpFileName, 3))
Case "css"
FileIcon = "css"
Case Else
FileIcon = "otherfile"
End Select
End Function

Function DeleteFile(lpFileName As String) As Boolean
'API
If ConfirmDelete(lpFileName) = False Then DeleteFile = False: Exit Function
DoEvents
Dim lpSH As SHFILEOPSTRUCT
With lpSH
.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
.hwnd = frmMain.hwnd
.lpszProgressTitle = "WonderHTML"
.pFrom = lpFileName
.wFunc = FO_DELETE
End With
Screen.MousePointer = 11
DeleteFile = SHFileOperation(lpSH) = 0
Screen.MousePointer = 0
End Function

Function MoveFile(lpFileName As String, lpNewDest As String) As Boolean
'API
Dim lpSH As SHFILEOPSTRUCT
With lpSH
.fFlags = FOF_ALLOWUNDO
.hwnd = frmMain.hwnd
.lpszProgressTitle = "WonderHTML"
.pFrom = lpFileName
.pTo = lpNewDest
.wFunc = FO_MOVE
End With
MoveFile = SHFileOperation(lpSH) = 0
End Function

Function CopyFile(lpFileName As String, lpNewDest As String) As Boolean
Dim lpSH As SHFILEOPSTRUCT
With lpSH
.fFlags = FOF_ALLOWUNDO
.hwnd = frmMain.hwnd
.lpszProgressTitle = "WonderHTML"
.pFrom = lpFileName
.pTo = lpNewDest
.wFunc = FO_COPY
End With
CopyFile = SHFileOperation(lpSH) = 0
End Function

Function RenameFile(lpFileName As String, lpNewName As String) As Boolean
Dim lpSH As SHFILEOPSTRUCT
With lpSH
.fFlags = FOF_ALLOWUNDO
.hwnd = frmMain.hwnd
.lpszProgressTitle = "WonderHTML"
.pFrom = lpFileName
.pTo = lpNewName
.wFunc = FO_RENAME
End With
RenameFile = SHFileOperation(lpSH) = 0
End Function

Function GetVarName(VarName As String) As String
'only var name
On Error Resume Next
Dim EqPos As Long
EqPos = InStr(1, VarName, "=")
If EqPos = 0 Then GetVarName = VarName: Exit Function
GetVarName = Left(VarName, EqPos - 1)
GetVarName = Trim(GetVarName)
End Function

Function GetFunctionName(FunctionName As String) As String
'only func name without args
On Error Resume Next
Dim EqPos As Long
EqPos = InStr(1, FunctionName, "(")
If EqPos = 0 Then GetFunctionName = FunctionName: Exit Function
GetFunctionName = Left(FunctionName, EqPos - 1)
GetFunctionName = Trim(GetFunctionName)
End Function

Function AddAssociation() As String
'returns user-chosen association
frmAss.Show vbModal
AddAssociation = assoc_text
End Function

Function NoStrings(Text As String) As String
'no strings
On Error Resume Next
Dim temp As String, tmps() As String, i As Long
temp = Text
tmps = Split(temp, Chr(34))
temp = ""
If UBound(tmps) = 0 Then GoTo nxt
For i = 0 To UBound(tmps)
If i And 1 Then tmps(i) = Space(Len(tmps(i)) + 2)
temp = temp & tmps(i)
Next i
tmps = Split(temp, "'")
nxt:
If UBound(tmps) = 0 Then GoTo finish
For i = 0 To UBound(tmps)
If i And 1 Then tmps(i) = Space(Len(tmps(i)) + 2)
temp = temp & tmps(i)
Next i
finish:
If temp = "" Then temp = Text
NoStrings = temp
End Function

Function NoComments(Text As String) As String
On Error Resume Next
Dim temp As String, midtemp As String
temp = Text
Dim l1 As Long, l2 As Long
'first find /* */ ones
Do
l1 = InStr(l2 + 1, temp, "/*")
If l1 = 0 Then NoComments = temp: GoTo nxt
l2 = InStr(l1 + 1, temp, "*/") + 2
If l2 = 2 Then NoComments = temp: GoTo nxt
midtemp = Mid$(temp, l1, l2 - l1)
Mid$(temp, l1, l2 - l1) = Space(Len(midtemp))
Loop
nxt:
'now for // ones
l1 = 0: l2 = 0
Do
l1 = InStr(l2 + 1, temp, "//")
If l1 = 0 Then NoComments = temp: Exit Function
l2 = InStr(l1 + 1, temp, vbNewLine)
If l2 = 0 Then NoComments = temp: Exit Function
midtemp = Mid$(temp, l1, l2 - l1)
Mid$(temp, l1, l2 - l1) = Space(Len(midtemp))
Loop
NoComments = temp
End Function

