VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpellWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ---------------------------------------------------------------------
'clsSpellWords: Dictionary class for clsSpellCheck
'  Created: 2000-07-11 by Scott Seligman <scott@scottandmichelle.net>
' ---------------------------------------------------------------------

Option Explicit

'Pointer for a list of words in a block
Private Type DictPointerType
TotalWords As Integer 'How many words?
Block As Byte 'Block number
Offset As Integer 'Offset within block
WordLength As Integer 'Length of words plus null characers
End Type

'Dictionary of compressed values
Private Type DictCompressType
RealSize As Byte 'Number of bytes used
Letter(1 To 5) As Byte 'Each byte
End Type

'All possible pointers for all soundex values
Private m_Pointers(1 To 26, 0 To 6, 0 To 6, 0 To 6) As DictPointerType
Private m_Compress(1 To 255) As DictCompressType

'The blocks themselves, collections of clsSpellBlock objects
Private m_cBlocks As Collection

'Properties for clsSpellBlock so it can expand words
Public Property Get CompressSize(nIndex As Byte) As Byte
CompressSize = m_Compress(nIndex).RealSize
End Property
Public Property Get CompressLetter(nIndex As Byte, _
nOffset As Byte) As Byte
CompressLetter = m_Compress(nIndex).Letter(nOffset)
End Property

Public Sub LoadFromFile(nFile As Long)
'Loas from a file, the file is structured like:
' (1) [long] Number of blocks
' (1 to 26, 0 to 6, 0 to 6, 0 to 6) Pointer into blocks
'    [integer] Number of words
'    [byte]    Block number
'    [integer] Offset within block
'    [integer] Word Lenght
' (1 to 255) Compressed dictionaries
'    [byte]    Number of bytes used
'    (1 to 5) [byte] bytes of expanded token
' (number of blicks) Block of data (32,767 bytes)

Dim nBlocks As Long
Dim i As Long
Dim obj As clsSpellBlock

'Read the header
Get #nFile, , nBlocks
Get #nFile, , m_Pointers()
Get #nFile, , m_Compress

'For each block, read in the block
Set m_cBlocks = New Collection
For i = 1 To nBlocks
Set obj = New clsSpellBlock
obj.LoadFromFile nFile
m_cBlocks.Add obj
Next

End Sub

Public Function Soundex(sWord As String) As String
'This returns a soundex version of the word.  Soundex is an attempt
' to codify how a word sounds, it always returns a four character
' (1 letter, 3 numbers) version of the word.  For more information on
' the format, and why it came about, see:
'
'http://www.nara.gov/genealogy/coding.html

If Asc(UCase(Mid$(sWord, 1, 1))) < 65 Or _
Asc(UCase(Mid$(sWord, 1, 1))) > 90 Then

Soundex = "z000"
Debug.Assert False

Else

Soundex = UCase$(Mid$(sWord, 1, 1)) & "000"

End If

Dim nSoundexPos  As Long
Dim nWordPos As Long
Dim sLastChar As String
Dim sChar As String

nSoundexPos = 2
For nWordPos = 2 To Len(sWord)
sChar = UCase$(Mid$(sWord, nWordPos, 1))
If sLastChar <> sChar Then
Select Case sChar
Case "B", "F", "P", "V"
Mid$(Soundex, nSoundexPos, 1) = "1"
nSoundexPos = nSoundexPos + 1
Case "C", "G", "J", "K", "Q", "S", "X", "Z"
Mid$(Soundex, nSoundexPos, 1) = "2"
nSoundexPos = nSoundexPos + 1
Case "D", "T"
Mid$(Soundex, nSoundexPos, 1) = "3"
nSoundexPos = nSoundexPos + 1
Case "L"
Mid$(Soundex, nSoundexPos, 1) = "4"
nSoundexPos = nSoundexPos + 1
Case "M", "N"
Mid$(Soundex, nSoundexPos, 1) = "5"
nSoundexPos = nSoundexPos + 1
Case "R"
Mid$(Soundex, nSoundexPos, 1) = "6"
nSoundexPos = nSoundexPos + 1
End Select
sLastChar = sChar
End If
If nSoundexPos >= 5 Then
Exit For
End If
Next

End Function

Public Function GetListFromWord(sWord As String) As Collection
'Returns a list of suggestion words given a word.  This just
' returns the soundex list for the given word.

Dim sSoundex As String
sSoundex = Soundex(sWord)

'Dereference the pointer
With m_Pointers(Asc(Mid(sSoundex, 1, 1)) - 64, _
Val(Mid(sSoundex, 2, 1)), Val(Mid(sSoundex, 3, 1)), _
Val(Mid(sSoundex, 4, 1)))

If .TotalWords > 0 Then
'If this a non-null pointer, then call .GetList for the
' proper block

Dim obj As clsSpellBlock
Set obj = m_cBlocks(.Block)
Set GetListFromWord = obj.GetList(LCase(Mid(sSoundex, 1, 1)) _
, CLng(.Offset), CLng(.WordLength), Me)

End If

End With

End Function

