VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUNDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'######################################
'WonderHTML 1.2 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
Option Explicit
Private Undo As Collection
Private Redo As Collection

Sub UndoChange(lpRTF As RichTextBox)
On Error Resume Next
Dim pos As Long, diff As Long
If Undo.Count > 1 Then diff = 1 Else diff = 0
'if only 1 action exists, there is no need to find the difference
'difference is the diference between the last action and the
'current action, which is to be reverted to the last action
pos = InStr(1, Undo.Item(Undo.Count - diff), ";")
'I append ; for saving the selstart
lpRTF.Text = Mid(Undo.Item(Undo.Count - diff), pos + 1, Len(Undo.Item(Undo.Count - diff)))
'the text is what comes after the ;
lpRTF.SelStart = ParseInt(Left(Undo.Item(Undo.Count - diff - 1), pos - 1))
'the selstart is what comes before the ;
'#############
'the current refers to the action or text that is
'currently present in the RTF box. we need only
'previous actions.
'#############
Undo.Remove Undo.Count 'kill the current one, we don't want it
Redo.Add Undo.Item(Undo.Count) 'put it in the redo buffer
Undo.Remove Undo.Count 'now remove it from undo buffer
End Sub

Sub RedoChange(lpRTF As RichTextBox)
On Error Resume Next
Dim pos As Long
'uses the same proc as Undo
pos = InStr(1, Redo.Item(Redo.Count), ";")
lpRTF.Text = Mid(Redo.Item(Redo.Count), pos + 1, Len(Redo.Item(Redo.Count)))
lpRTF.SelStart = ParseInt(Left(Redo.Item(Redo.Count), pos + 1))
Redo.Remove Redo.Count 'now that we're done
Undo.Remove Undo.Count
If Undo.Count > 1 Then Undo.Remove Undo.Count
End Sub

Sub AddAction(ActionText As String, SelectionStart As Long)
'add to undo buffer
If Undo.Count > 100 Then Undo.Remove 1
'limit for undo buffer: 100 actions per document
Undo.Add SelectionStart & ";" & ActionText
'add using a ; delimiter to save selStart as well
'add only if previous one is different
frmMain.TB.Buttons("undo").Enabled = (Undo.Count > 0)
'check if any item exists, if not no need to keep undo enabled
End Sub

Sub ResetAll()
Set Undo = New Collection
Set Redo = New Collection
'finish it off
End Sub

Private Sub Class_Initialize()
ResetAll
End Sub

Private Sub Class_Terminate()
Set Undo = Nothing
Set Redo = Nothing
End Sub

Function UndoAvailable() As Boolean
UndoAvailable = (Undo.Count > 0)
End Function

Function Count() As Long
Count = Undo.Count
End Function

Sub Remove(lpIndex As Long)
Undo.Remove lpIndex
End Sub

Function RedoAvailable() As Boolean
RedoAvailable = (Redo.Count > 0)
End Function
