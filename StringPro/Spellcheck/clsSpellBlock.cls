VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpellBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ---------------------------------------------------------------------
'clsSpellBlock: A "block" of words, used by clsSpellWords
'  Created: 2000-07-11 by Scott Seligman <scott@scottandmichelle.net>
' ---------------------------------------------------------------------

Option Explicit

'The block of words
Private m_Block(1 To 32767) As Byte

'Get/Let a specific byte
Public Property Get Block(nOffset As Long) As Byte
Block = m_Block(nOffset)
End Property
Public Property Let Block(nOffset As Long, nValue As Byte)
m_Block(nOffset) = nValue
End Property

Public Sub LoadFromFile(nFile As Long)
'Load the block from a file, using VB's functions

Get #nFile, , m_Block()

End Sub

Public Function GetList(sLetter As String, nStart As Long, nLength _
As Long, parent As clsSpellWords) As Collection
'This returns a list of words from the block.
'
'The words are stored in the list with vbNullChar between each word,
' and each missing the first letter, so this list takes the total
' length of all words (and null characters), the starting position
' within this block, and the first letter for all the words and
' returns a collection of the words.

Set GetList = New Collection

Dim sWord As String
Dim nStrPos As Long
Dim nCurPos As Long
Dim nByte As Byte
Dim i As Byte

sWord = Space(255)
nStrPos = 0

For nCurPos = nStart To nStart + nLength
nByte = m_Block(nCurPos)
If nByte = 0 Then
GetList.Add sLetter & Mid(sWord, 1, nStrPos)
nStrPos = 0
Else
'Exand this letter into it's token
For i = 1 To parent.CompressSize(nByte)

nStrPos = nStrPos + 1

Mid(sWord, nStrPos, 1) = _
Chr(parent.CompressLetter(nByte, i))

Next
End If
Next

End Function
