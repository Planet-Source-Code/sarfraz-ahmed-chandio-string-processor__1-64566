VERSION 5.00
Begin VB.Form frmSpellCheck 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "String Processor"
   ClientHeight    =   4380
   ClientLeft      =   4185
   ClientTop       =   2265
   ClientWidth     =   3720
   ControlBox      =   0   'False
   HelpContextID   =   1330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   3720
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ListBox lstWords 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2370
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2115
   End
   Begin VB.TextBox txtWord 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   75
      Picture         =   "frmSpellCheck.frx":0000
      Top             =   120
      Width           =   3570
   End
   Begin VB.Label lblNoList 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  (no suggestions)"
      Height          =   675
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suggested Entries:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suspicious Entries:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1635
   End
End
Attribute VB_Name = "frmSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I thank Mr. Scott Seligman for providing me with the
'facility of the Spell-Checker.

' ---------------------------------------------------------------------
'frmSpellCheck: Dialog that presents the user with a mis-spelled
'    word, and possible a list of suggested words.
'  Created: 2000-07-11 by Scott Seligman <scott@scottandmichelle.net>
' ---------------------------------------------------------------------

Option Explicit

Private m_sWord As String 'The word in question
Private m_sReplace As String 'The selected word to replace it with
Private m_bCancel As Boolean 'Did the user select cancel?
Private m_cSuggestions As Collection 'Collection of suggested words

Public Sub ReplaceWord(sWord As String, sReplace As String, _
pParent As clsSpellCheck, bCancel As Boolean)
'This is the main entry point for this form:
'  sWord: The mis-spelled word
'  sReplace: [out] The word the user selected to replace it with
'  pParent: The clsSpellCheck master class
'  bCancel: [out] Did the user click on "cancel"?

'Set our module variables
m_sWord = sWord
Set m_cSuggestions = pParent.GetLastList()

'Show the word on the form
txtWord.Text = m_sWord
txtWord.SelStart = 0
txtWord.SelLength = Len(txtWord)

'Add the suggestions, if we can
If Not (m_cSuggestions Is Nothing) Then
Dim vWord As Variant
For Each vWord In m_cSuggestions
lstWords.AddItem vWord
Next
End If

If lstWords.ListCount = 0 Then
'There aren't any suggestions, just put a
' label in the list's place

lblNoList.Top = lstWords.Top
lblNoList.Left = lstWords.Left
lblNoList.Width = lstWords.Width
lblNoList.Height = lstWords.Height
lblNoList.Visible = True
lstWords.Visible = False

End If

'Nothing has been selected yet, so disable this button
cmdChange.Enabled = False

'Show the form
Me.Show vbModal

'Pass the return variables back to the callee
sReplace = m_sReplace
bCancel = m_bCancel

End Sub

Private Sub cmdCancel_Click()
'The user clicked cancel

m_bCancel = True
Unload Me

End Sub

Private Sub cmdChange_Click()
'The user selected to change a word (may also be called indirectly
' via double-clicking on the listbox)

m_sReplace = txtWord
Unload Me

End Sub

Private Sub cmdIgnore_Click()
'Just skip this word

Unload Me

End Sub

Private Sub Form_Activate()
DropShadow cmdCancel, Me
DropShadow cmdIgnore, Me
DropShadow cmdChange, Me
DropShadow txtWord, Me
DropShadow lstWords, Me
End Sub

Private Sub Form_Load()
Me.Icon = frmText.Icon
End Sub

Private Sub Form_Paint()
DropShadow cmdCancel, Me
DropShadow cmdIgnore, Me
DropShadow cmdChange, Me
DropShadow txtWord, Me
DropShadow lstWords, Me
End Sub

Private Sub lstWords_Click()
'Single clicking on the list causes the selected word to be
' displayed in the text-box

If lstWords.ListIndex <> -1 Then
txtWord.Text = lstWords.List(lstWords.ListIndex)
End If

End Sub

Private Sub lstWords_DblClick()
'Double clicking on the list is the same as pressing the
' cmdChange button

If cmdChange.Enabled = True Then
cmdChange.Value = True
End If

End Sub

Private Sub txtWord_Change()
'If the textbox changes, enable the change button, but only
' if the text box contains a new word

If txtWord <> m_sWord Then
cmdChange.Enabled = True
Else
cmdChange.Enabled = False
End If

End Sub

Private Sub txtWord_KeyDown(KeyCode As Integer, Shift As Integer)
'Let the user use the up/down keys from within the textbox

If lstWords.Visible = True Then
If KeyCode = vbKeyUp And Shift = 0 Then
If lstWords.ListIndex = -1 Then
lstWords.ListIndex = 0
Else
If lstWords.ListIndex > 0 Then
lstWords.ListIndex = lstWords.ListIndex - 1
End If
End If
ElseIf KeyCode = vbKeyDown And Shift = 0 Then
If lstWords.ListIndex = -1 Then
lstWords.ListIndex = 0
Else
If lstWords.ListIndex < lstWords.ListCount - 1 Then
lstWords.ListIndex = lstWords.ListIndex + 1
End If
End If
End If
End If

End Sub

