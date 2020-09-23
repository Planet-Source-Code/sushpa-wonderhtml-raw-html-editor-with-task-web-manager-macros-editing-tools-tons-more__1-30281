Attribute VB_Name = "mString"
'===================================
'STRINGS module - String manipulation operations
'===================================
'VB6LIB.DLL - Enhanced Subs & Functions for
'Visual Basic 5.0/6.0 - By Sushant Pandurangi
'===================================
'Visit me at http://sushantshome.tripod.com
'===================================
'Your comments and suggestions are welcome
'You can email me at sushant@phreaker.net
'===================================
Option Explicit
'===================================

Public Function StrCount(sString As String, sChar As String) As Integer
Attribute StrCount.VB_Description = "Count how many times the specified character appears in the string."
Dim pStep As Long, pCount As Long
'Somehow I havent seen such a function in the object
'browser till today. VB needed this one badly.
    pStep = InStr(1, sString, sChar)
'pStep is the first occurence
    If pStep = 0 Then Exit Function 'Char dosent exist
looper:
    Do
        pStep = InStr(pStep + 1, sString, sChar)
        pCount = pCount + 1
    Loop Until pStep = 0
        StrCount = pCount
End Function

Public Function Repeat(sCharacter As String, Length As Long) As String
Attribute Repeat.VB_Description = "Repeats a string character."
'Maybe the String() function I see in VB6 is present in
'VB5 as well. Maybe not, so here is a confirmation.
Dim i As Integer, temp As String
For i = 1 To Length
temp = temp & sCharacter
Next i
Repeat = temp
temp = "": i = 0
End Function

Public Function Reverse(sString As String) As String
Attribute Reverse.VB_Description = "Reverse a given string."
'VB6 has this as an in-built function called
'StrReverse(String) but I am not sure of VB5.
Dim i As Integer, s As String
For i = 1 To Len(sString)
s = s & Mid(sString, Len(sString) + 1 - i, 1)
Next i
Reverse = s
End Function

Public Function CBinary(Expression As Boolean) As Integer
Attribute CBinary.VB_Description = "Convert the given boolean to 0 or 1."
'Useful for converting BOOLs to 0 or 1. CByte() would
'return 255 for true, which wont be useful for setting the
'values of, for instance, a checkbox; as it uses 0 and 1.
If Expression = False Then CBinary = 0 Else CBinary = 1
End Function

Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
frmAbout.Show vbModal
End Sub

Function WdCount(pString As String) As Long
Attribute WdCount.VB_Description = "Count the number of words in the string."
'Number of words; decided using number of spaces and other characters
WdCount = StrCount(pString, " ") + 1 + StrCount(pString, "=") + StrCount(pString, "-") + StrCount(pString, "+") + StrCount(pString, "\") + StrCount(pString, "/") + StrCount(pString, ".")
End Function

Function LnCount(pTextBox As Object) As Integer
Attribute LnCount.VB_Description = "Get the number of lines in a Textbox."
'Number of lines
LnCount = SendMessage(pTextBox.hwnd, &HBA, 0, 0&)
End Function

Function SnCount(pText As String) As Integer
Attribute SnCount.VB_Description = "Get number of sentences."
'Number of sentences
SnCount = StrCount(pText, ".")
End Function

Function Up1Level(sPath As String, Optional Sep As String = "\") As String
Attribute Up1Level.VB_Description = "Return the folder that is up one level from the given one."
'Name of directory up one level from that given
Dim pos As Long, i As Integer, Dummy As String
If Right(sPath, 1) = Sep Then sPath = Left(sPath, Len(sPath) - 1)
Dummy = Reverse(sPath)
pos = InStr(1, Dummy, Sep)
Up1Level = Right$(Dummy, Len(Dummy) - pos)
Up1Level = Reverse(Up1Level)
If Right(Up1Level, 1) = ":" Then Up1Level = Up1Level & Sep
End Function

Function GetFile(sPath As String) As String
Attribute GetFile.VB_Description = "Get the filename portion from the string"
    'Returns only file title
    Dim i, j As Integer
    i = InStr(1, Reverse(sPath), "\")
    If i = 0 Then i = InStr(1, Reverse(sPath), "/")
    If i = 0 Then GetFile = sPath: Exit Function
    GetFile = Right(sPath, i - 1)
End Function

Function GetPath(sPath As String) As String
Attribute GetPath.VB_Description = "Get the pathname portion from the string"
    'Returns only path name without file title
    GetPath = Up1Level(sPath)
End Function

Function InitCap(sString As String) As String
Attribute InitCap.VB_Description = "Returns string with initial capitals."
    'First letter caps
    InitCap = UCase(Left(sString, 1)) & (Right(sString, Len(sString) - 1))
End Function

Public Function WebSafe(Text As String) As String
'This function will return HTML numerical entity values
'based on strings.
Dim pos As Integer, temp As String
For pos = 1 To Len(Text)
temp = temp & "&#" & Asc(Mid(Text, pos, 1)) & ";"
Next pos
WebSafe = temp
End Function

Function NoExt(File As String)
Dim l As Long
l = InStrRev(File, ".")
If l = 0 Then NoExt = File: Exit Function
NoExt = Left(File, l - 1)
End Function

Function Drive(DriveX As String) As String
Dim l As Long
l = InStr(1, DriveX, "[")
If l = 0 Then Drive = DriveX: Exit Function
Drive = Left(DriveX, l - 1)
Drive = Trim(UCase(Drive)) & "\"
End Function

Function GetCSSElementName(ByVal StrCSS As String) As String
Dim pos As Long
pos = InStr(StrCSS, "{")
If pos = 0 Then GetCSSElementName = StrCSS: Exit Function
StrCSS = Left(StrCSS, pos - 1)
StrCSS = Replace(Trim(StrCSS), vbTab, "")
GetCSSElementName = IIf(Left(StrCSS, 1) <> "#" And InStr(StrCSS, ".") = 0, UCase(StrCSS), StrCSS)
End Function


Function GetClrRGBVal(Color As Long, Optional delim As String = vbCrLf) As String
Dim R, G, b
R = Color Mod 256
G = (Color \ 256) Mod 256
b = Color \ 65536
GetClrRGBVal = "Red: " & R & delim & "Green: " & G & delim & "Blue: " & b
End Function

Function GetHexVal(Color As Long) As String
Dim R, G, b
R = Hex(Color Mod 256)
G = Hex((Color \ 256) Mod 256)
b = Hex(Color \ 65536)
If Len(R) < 2 Then R = "0" & R
If Len(G) < 2 Then G = "0" & G
If Len(b) < 2 Then b = "0" & b
GetHexVal = "#" & R & G & b
End Function
