VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ---------------------------------------------------------------------
'clsSpellCheck: Main class for the spell checking system,
'
'Properties:
' Left - Position of dialogs
' Top - Position of dialogs (realitive to screen)
'
'Methods:
' LoadDict - Loads a dictionary file (must be called before
'            anything else)
' CheckTextBox - Loops through a text box and calls CheckWord for
'                each word in the textbox
' IsWordValid - Returns true if a word is found in the dictionary
'               (and stores the list of suggested words),
'               never displays a UI
' CheckWord - Checks the spelling of a word, displaying a UI if the
'             word is mis-spelled
' GetLastList - Returns the last list of suggested words from
'               IsWordValid or CheckWord
'
'  Created: 2000-07-11 by Scott Seligman <scott@scottandmichelle.net>
' ---------------------------------------------------------------------


Option Explicit

Private m_Words As clsSpellWords 'The collection that
'holds the dictionary

Private m_nLeft As Single 'Position of the dialog
Private m_nTop As Single

Private m_cLastList As Collection 'The last list of suggestions
Private m_sLastWord As String 'The word that was used to get the list
Private m_bWordValid As Boolean 'Is the word mis-spelled?

Public Property Get Left() As Single
Left = m_nLeft
End Property
Public Property Let Left(nLeft As Single)
m_nLeft = nLeft
End Property

Public Property Get Top() As Single
Top = m_nTop
End Property
Public Property Let Top(nTop As Single)
m_nTop = nTop
End Property

Public Sub LoadDict(sDictName As String)
'Loads a dictionary file

Dim nFile As Long
nFile = FreeFile
Open sDictName For Binary As #nFile
m_Words.LoadFromFile nFile
Close #nFile

End Sub

Private Sub Class_Initialize()

Set m_Words = New clsSpellWords

End Sub

Public Function CheckTextBox(txt As TextBox) As Boolean
CheckTextBox = False

Dim sAlphaChar As String 'List of possible characters in a word
Dim nLastPos As Long 'Last position of a character found
Dim nCurPos As Long 'Current position
Dim nEndingPos As Long 'Ending position (ie, Len(txtBox))
Dim bInWord As Boolean 'Are we currently inside of a word?

'Used in calling CheckWord
Dim bReplace As Boolean 'User wants to replace a word
Dim sWord As String 'The mis-spelled word
Dim sReplace As String 'The replacement
Dim bCancel As Boolean 'The user just wants to cancel
'the whole shebang

'Mimics bInWord, expect at the end of the textbox
Dim bInWordCheck As Boolean

'These are the only characters used in our words, everything
' else is considered a space
sAlphaChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'"

'Initialize everything
bInWord = False
nLastPos = 1
nCurPos = 1
nEndingPos = Len(txt)

'Loop through the entire textbox, plus one character
Do Until nCurPos > nEndingPos + 1

'If this is the character after the end, treat it as a space,
' to spell check the final word
If nCurPos > nEndingPos Then
bInWordCheck = False
Else
'See if the current letter is a space, or a letter
bInWordCheck = (InStr(1, sAlphaChar, _
UCase(Mid(txt, nCurPos, 1))) > 0)
End If

If bInWordCheck Then
If Not bInWord Then
'In other words: We aren't inside a word yet, but
' we've found a letter, so mark the position
bInWord = True
nLastPos = nCurPos
End If
Else
If bInWord Then
'In other words, we were inside a word, but not any more

'Grab the current word
sWord = Mid(txt, nLastPos, nCurPos - nLastPos)

'Make sure the word starts with a letter
Do Until sWord = ""
If Asc(UCase(Mid$(sWord, 1, 1))) >= 65 And _
Asc(UCase(Mid$(sWord, 1, 1))) <= 90 Then
Exit Do
Else
sWord = Mid(sWord, 2)
End If
Loop

'Is this word misspelled?
If Not IsWordValid(sWord) Then

'Select it
txt.SelStart = nLastPos - 1
txt.SelLength = nCurPos - nLastPos

'Check the word, this is where the UI is displayed
CheckWord sWord, sReplace, bReplace, bCancel

'If the user clicked cancel, then abort out of
' this function
If bCancel Then
Exit Function
End If

'The user wants to replace the word
If bReplace Then

'Try to make the replacing word mimic the
' current capilization style (first letter, or
' all letters)
If UCase(sWord) = sWord Then

sReplace = UCase(sReplace)

ElseIf Mid(sWord, 1, 1) = _
UCase(Mid(sWord, 1, 1)) Then

Mid(sReplace, 1, 1) = UCase(Mid(sReplace, 1, 1))

End If

'Replace it
On Error GoTo Hell
txt.SelText = sReplace
frmText.mnuUndo.Enabled = False
'If the size differs from the original, make sure
' we note the new ending position
If Len(sReplace) <> Len(sWord) Then
nEndingPos = Len(txt)
End If

'Set the current position to the start of the
' word, this way we can spell check the word
' again (in case the user entered a bad word,
' they can ignore it if they delibertly did that)
nCurPos = nLastPos - 1
End If
End If

bInWord = False
End If
End If
nCurPos = nCurPos + 1
Loop

'Let the callee know we completed successfully (ie, the user
' didn't click cancel)
CheckTextBox = True
MsgBox "Spell Check is complete!", vbInformation

Exit Function
Hell:
HellError
End Function

Public Function IsWordValid(sWord As String) As Boolean
'Is the word valid

Dim vWord As Variant

'Don't care about small or large words
If Len(sWord) <= 1 Or Len(sWord) > 40 Then
IsWordValid = True

Set m_cLastList = Nothing
m_sLastWord = sWord
m_bWordValid = True

Exit Function
End If

'Is this the same word we did last time?
If m_sLastWord <> sWord Then

'Try and get a list of suggestions
Set m_cLastList = m_Words.GetListFromWord(sWord)
m_sLastWord = sWord

'No list, so the word doesn't exist
If m_cLastList Is Nothing Then
IsWordValid = False
Else
'We got a list, see if the word is in the list
IsWordValid = False
For Each vWord In m_cLastList
If vWord = LCase(sWord) Then
IsWordValid = True
Exit For
End If
Next
End If
m_bWordValid = IsWordValid
Else
IsWordValid = m_bWordValid
End If

End Function

Public Sub CheckWord(ByVal sWord As String, sReplace As String, _
bReplace As Boolean, bCancel As Boolean)
'Check a single word, display a UI if need be.
' sWord = Word to check
' sReplace = [out] Replacement word
' bReaplce = [out] sReplace is valid
' bCancel = [out] User click cancel?

'Dont care about small or large words
If Len(sWord) <= 1 Or Len(sWord) > 40 Then
Exit Sub
End If

Dim vWord As Variant

'If the word is a good word, then just exit this function
If IsWordValid(sWord) Then
bReplace = False
sReplace = sWord
Else
'It's a bad word, so pop-up the UI
Dim frm As frmSpellCheck
Set frm = New frmSpellCheck
frm.Left = m_nLeft
frm.Top = m_nTop

frm.ReplaceWord sWord, sReplace, Me, bCancel

'Let the callee know what the users intentions are
If sWord <> sReplace Then
If sReplace = "" Then
bReplace = False
Else
bReplace = True
End If
End If
End If

End Sub

Public Function GetLastList() As Collection
'Just returns the last list of suggestions, used by the UI to
' display the list

Set GetLastList = m_cLastList

End Function
