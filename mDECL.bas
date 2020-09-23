Attribute VB_Name = "mDECL"
'######################################
'WonderHTML 1.2 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
Option Explicit

Public Enum EnumActions
ACTION_FILELIST = 45
ACTION_FINDFILES = 46
ACTION_REPORT = 47
'used large numbers so It won't interfere with other constants
End Enum

Public SearchFlag As Boolean, ExitFlag As Boolean
Public returnedMTYPE As String

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long

Public IfDeleteFile As Boolean

Public Type POINTAPI
        X As Long
        y As Long
End Type

Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As Rect) As Long

Public LastFindText As String

Public Const EM_UNDO = &HC7
Public Const WM_CLEAR = &H303
Public Const WM_KILLFOCUS = &H8
Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = (WM_USER + 63)
Public Const EM_LINEINDEX = &HBB
'Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEFROMCHAR = &HC9

Public Const FLAG_RO = cdlOFNPathMustExist + cdlOFNFileMustExist + cdlOFNOverwritePrompt + cdlOFNReadOnly
Public Const FLAG_RW = cdlOFNPathMustExist + cdlOFNFileMustExist + cdlOFNOverwritePrompt

Public gLeftMargin As Integer           ' Print Preview
Public gRightMargin As Integer          ' Print Preview
Public gTopMargin As Integer            ' Print Preview
Public gBottomMargin As Integer         ' Print Preview
Public gPrint As Boolean                ' Print Preview


Public ReturnedPath As String

Public frmMain As frmMDI

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Enum ERECViewModes
    ercDefault = 0
    ercWordWrap = 1
    ercWYSIWYG = 2
End Enum


Public Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Const EM_CHARFROMPOS = &HD7


Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_NOCONFIRMATION = &H10

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const Domains = "http:// www. .com .net .org .edu .ac .mil .gov ftp:// gopher:// telnet:// news: mailto: wais: javascript:"

Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
 
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE And KEY_ENUMERATE_SUB_KEYS And KEY_NOTIFY And KEY_CREATE_SUB_KEY And KEY_CREATE_LINK And KEY_SET_VALUE
Const REG_OPTION_NON_VOLATILE = 0
Const REG_OPTION_VOLATILE = 1
Const REG_SZ = 1

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long


Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = DI_MASK Or DI_IMAGE

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFilechkattrib As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000


Public Function Findfile(sFileName As String) As WIN32_FIND_DATA
'get the file and its info
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    
    ' Find file and get file data
    plngFirstFileHwnd = FindFirstFile(sFileName, Win32Data)
    If plngFirstFileHwnd = 0 Then
        Findfile.cFileName = "Error"
    Else
        Findfile = Win32Data
    End If
    plngRtn = FindClose(plngFirstFileHwnd)
End Function

Public Function GetFileInfo(dFileName As String, lbModified As Object, lbCreated As Object, lbAccessed As Object)
'get the info
    On Error Resume Next
    Dim FILETIME As SYSTEMTIME
    Dim filedata As WIN32_FIND_DATA
    
    filedata = Findfile(dFileName) 'Find file and get data
        
    'TIME RETURNED IS GMT!
    'I'm not totally confident of the current method, but let's see...
    
    'Modified
    FileTimeToSystemTimeEx filedata.ftLastWriteTime, FILETIME, dFileName
    lbModified = Format(FILETIME.wDayOfWeek + 1, "ddd") & ", " & Format(FILETIME.wDay, "00") & " " & GetMonth(FILETIME.wMonth) & " " & FILETIME.wYear & ", " & NoAMPM(FILETIME.wHour) & ":" & Format(FILETIME.wMinute, "00") & " " & AMPM(FILETIME.wHour)
    'Created
    FileTimeToSystemTimeEx filedata.ftCreationTime, FILETIME, dFileName
    lbCreated = Format(FILETIME.wDayOfWeek + 1, "ddd") & ", " & Format(FILETIME.wDay, "00") & " " & GetMonth(FILETIME.wMonth) & " " & FILETIME.wYear & ", " & NoAMPM(FILETIME.wHour) & ":" & Format(FILETIME.wMinute, "00") & " " & AMPM(FILETIME.wHour)
    ' Accessed
    FileTimeToSystemTimeEx filedata.ftLastAccessTime, FILETIME, dFileName
    lbAccessed = Format(FILETIME.wDayOfWeek + 1, "ddd") & ", " & Format(FILETIME.wDay, "00") & " " & GetMonth(FILETIME.wMonth) & " " & FILETIME.wYear
End Function

Public Sub PaintIcon(dFileName As String, PB As PictureBox)
    Dim lIcon As Long
    ' Extract assocciated icon from file
    lIcon = ExtractAssociatedIcon(App.hInstance, dFileName, 0&)
    DrawIconEx PB.hdc, 0, 0, lIcon, 0, 0, 0, 0, DI_NORMAL        'Draw icon in picturebox
    DestroyIcon lIcon 'Destroy icon
End Sub

Public Function GetValue(hKey As Long, SubKey As String, ValueName As String, Optional Default As String = "")
'get value from registry
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = 0 Then 'If key exist
        sData = String(260, vbNullChar) 'Fill buffer with null chars
        lngRet = RegQueryValueEx(lngResult, ValueName, 0, REG_SZ, ByVal sData, Len(sData))
        If Not lngRet = 0 Then GetValue = Default: Exit Function
        GetValue = Left(sData, InStr(1, sData, vbNullChar) - 1)
        RegCloseKey lngResult
    Else
        GetValue = Default
    End If
End Function

Function AMPM(ByVal Hour As Integer) As String
If Hour = 0 Then Hour = 24
If Hour > 12 Then AMPM = "PM" Else AMPM = "AM"
End Function

Function NoAMPM(ByVal Hour As Integer) As Integer
'e.g. 23:00 returns 11:00 PM
If Hour = 0 Then Hour = 24
If Hour > 12 Then NoAMPM = Hour - 12 Else NoAMPM = Hour
End Function

Function ConfirmDelete(Filename As String) As Boolean
Load frmDel
frmDel.Label1.Caption = frmDel.Label1.Caption & vbCrLf & Filename
frmDel.Show vbModal
ConfirmDelete = IfDeleteFile
End Function

Sub AddFileItem(File As String)
'used in generate file list
If GetFile(File) = "files.inf" Then Exit Sub
Dim s As String, Web As String
Web = (frmMain.tvW.Nodes(1).Key)
s = FirstSlash(Replace(File, Web, ""))
Web = Replace(Web, "\", "/")
s = Replace(s, "\", "/")
frmMain.fileList = "<TR><TD><A href=" & Chr(34) & s & Chr(34) & _
">" & GetTitle(File) & "<BR></A>" & _
s & "<BR></TD><TD>" & Format(FileLen(File), "0,000") & _
" Bytes" & "</TD><TD>" & FileDateTime(File) & _
"</TD></TR>" & vbNewLine & frmMain.fileList
End Sub

Function GenFileList(NewPath As String) As Integer
Screen.MousePointer = 11
DirDiver NewPath, NewPath, True, ACTION_FILELIST, 0, "", vbTextCompare
Screen.MousePointer = 0
End Function

Function GetTitle(Path As String) As String
On Error Resume Next
Dim s As String, l As Long, l2 As Long
If InStr(" htm html asp xml ", " " & Ext(Path) & " ") = 0 Then GetTitle = GetFile(Path): Exit Function
Open Path For Binary As #1
s = Space$(LOF(1))
Get #1, , s
Close #1
l = InStr(1, s, "<TITLE>", vbTextCompare)
If l = 0 Then GetTitle = GetFile(Path): Exit Function
l2 = InStr(l + 1, s, "</TITLE>", vbTextCompare)
If l2 = 0 Then GetTitle = GetFile(Path): Exit Function
s = Mid$(s, l + 7, l2 - l - 7)
GetTitle = s
End Function

Sub ShowReport(Web As String)
On Error Resume Next
Screen.MousePointer = 11: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(2).Picture
frmMain.SB.Panels(1).Text = "Generating Report for " & Web & ", please wait..."
Load frmRept
frmRept.InitializeRept Web
frmRept.Show
frmMain.SB.Panels(1).Text = ""
Screen.MousePointer = 0: frmMain.SB.Panels(4).Picture = frmMain.imlBusy.ListImages(1).Picture
End Sub

Function GetSpeedInfo(Size) As String
On Error Resume Next
Dim l As Single, s As String, div As Long
s = ReadValue("ConnectionSpeed", , "Reports")
Select Case LCase(s)
Case "14.4 k", "28.8 k", "56.6 k", "33.6 k"
div = ParseInt(s) * 100 '100 cause the . is ignored by ParseInt. Else use 1000. It's not 1024 here, as 28800 bps is standard, not 28.8kbps, which is just used in notation. It'd be 28.125kbps if 28800/1024 was used.
Case "ds0 (isdn)"
div = 65536 '64kbps isdn
Case "ds1 (t1)" 'digital signal levels 1 thru 4 (T1-T2-T3-T4 levels)
div = 154400
Case "ds2 (t2)"
div = 6300000
Case "ds3 (t3)"
div = 44746000
Case "ds4 (t4)"
div = 274176000
End Select
div = div / 8 '8 bits per byte. bps is bits per second.
l = Round(Size / div, 1)
frmMain.SB.Panels(5).Text = " " & MMSS(l) & " @ " & s & " "
GetSpeedInfo = MMSS(l) & " @ " & Round((div / 1000), 1) & " KB/s (" & s & ")"
End Function

Function MMSS(whole) As String
Dim mm, ss, sInt As Integer
mm = whole \ 60
ss = ModDecimal(whole, 60)
ss = Round(ss, 0)
If Len(mm) < 2 Then mm = "0" & mm
If Len(ss) < 2 Then ss = "0" & ss
MMSS = mm & ":" & ss
End Function

Function ModDecimal(What, Divider) As Single
Dim l As Single
l = Divider * (What \ Divider)
ModDecimal = What - l
End Function

Sub DoFileAction(File As String, Action As EnumActions, Optional FindString As String, Optional FindCase As VbCompareMethod, Optional ReportAction As Long)
Dim s As String
Select Case Action
Case ACTION_FILELIST
AddFileItem File
frmMain.SB.Panels(1).Text = Up1Level(File, "\")
Case ACTION_FINDFILES
Open File For Binary As #1
s = Space$(LOF(1))
Get #1, , s
Close #1
frmFind.SB.SimpleText = File
If InStr(1, s, FindString, FindCase) > 0 Then frmFind.lvFiles.ListItems.Add , File, FirstSlash(Replace(File, frmFind.txLoc.Text, "", , , vbTextCompare)), , 1
Case ACTION_REPORT
frmRept.DoReportAction File, ReportAction
End Select
End Sub

Sub DoPathAction(Path As String, Action As EnumActions)
'nothing yet
End Sub

Function DirDiver(NewPath As String, BackUp As String, bSub As Boolean, Action As Long, ReportAction As Long, FindString As String, FindCase As VbCompareMethod) As Integer
On Error Resume Next
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String

Dim retval As Integer

    If ExitFlag = True Then
        DirDiver = True
        Exit Function
    End If
    
    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiver = False            ' Set to True if there is an error.
    
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    
    
        If bSub Then DirsToPeek = frmMain.Fldr.ListCount Else DirsToPeek = 0               ' How many directories below this?
    
    Do While DirsToPeek > 0 And SearchFlag = True
    
        OldPath = frmMain.Fldr.Path                      ' Save old path for next recursion.
        frmMain.Fldr.Path = NewPath
        If frmMain.Fldr.ListCount > 0 Then
            ' Get to the node bottom.
            frmMain.Fldr.Path = frmMain.Fldr.list(DirsToPeek - 1)
            retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
            AbandonSearch = DirDiver((frmMain.Fldr.Path), OldPath, bSub, Action, ReportAction, FindString, FindCase)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If frmMain.Fil.ListCount Then
        If Len(frmMain.Fldr.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = frmMain.Fldr.Path                  ' If at root level, leave as is...
        Else
            ThePath = frmMain.Fldr.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        Dim s As String
        For ind = 0 To frmMain.Fil.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath & frmMain.Fil.list(frmMain.Fil.ListCount - ind - 1)
            DoFileAction entry, Action, FindString, FindCase, ReportAction
        Next ind
n:
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        frmMain.Fldr.Path = BackUp
    End If
    Exit Function
End Function

Sub FileTimeToSystemTimeEx(lpF As FILETIME, lpS As SYSTEMTIME, File As String)
'This will call FileTimeToSystemTime, convert into local time
'zone using the difference in hours in the GMT SystemTime
'and the local FileDateTime.
FileTimeToSystemTime lpF, lpS
Static localHours As Integer
Static localMins As Integer
If localHours = 0 And localMins = 0 Then
  localHours = Format(FileDateTime(File), "hh")
  localHours = localHours - lpS.wHour
  localMins = Right(Format(FileDateTime(File), "hh:mm"), 2)
  localMins = localMins - lpS.wMinute
  If localMins < 0 Then localMins = -localMins: localHours = localHours - 1
End If
lpS.wHour = lpS.wHour + localHours
lpS.wMinute = lpS.wMinute + localMins
If lpS.wMinute >= 60 Then lpS.wMinute = lpS.wMinute - 60: lpS.wHour = lpS.wHour + 1
End Sub

Function GetTitleFromText(Text As String) As String
On Error Resume Next
Dim l As Long, l2 As Long, s As String
l = InStr(1, Text, "<TITLE>", vbTextCompare)
If l = 0 Then GetTitleFromText = "Untitled": Exit Function
l2 = InStr(l + 1, Text, "</TITLE>", vbTextCompare)
If l2 = 0 Then GetTitleFromText = "Untitled": Exit Function
s = Mid$(Text, l + 7, l2 - l - 7)
GetTitleFromText = s
End Function

Function chomp(Text As String) As String
If Right(Text, Len(vbNewLine)) = vbNewLine Then chomp = Left(Text, Len(Text) - Len(vbNewLine)) Else chomp = Text
End Function

Function GetMetaType(ID As String, Value As String, Content As String) As String
Load frmMType
frmMType.lbID.Caption = ID
If Value = "name" Then frmMType.opt(0).Value = True Else frmMType.opt(1).Value = True
frmMType.txContent.Text = Content
frmMType.Show vbModal
GetMetaType = returnedMTYPE
returnedMTYPE = ""
End Function

Sub BrowseClr(TB As Object)
On Error Resume Next
Dim s As String
Clipboard.Clear
Load frmCPick
frmCPick.Command1.Caption = "C&ontinue"
frmCPick.Show vbModal
s = Clipboard.GetText
If Len(s) = 7 And Left(s, 1) = "#" Then TB.Text = s
TB.SetFocus
End Sub

