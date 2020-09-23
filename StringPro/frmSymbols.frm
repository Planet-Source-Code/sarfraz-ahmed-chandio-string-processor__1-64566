VERSION 5.00
Begin VB.Form frmSymbols 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Symbol"
   ClientHeight    =   5370
   ClientLeft      =   3675
   ClientTop       =   2610
   ClientWidth     =   7200
   HelpContextID   =   1290
   Icon            =   "frmSymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInsert 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   160
      Width           =   3855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1092
   End
   Begin VB.ComboBox cboFonts 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopy 
      Cancel          =   -1  'True
      Caption         =   "Co&py"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   1092
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3855
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   252
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1092
   End
   Begin VB.PictureBox picHolder 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   5475
      TabIndex        =   3
      Top             =   1080
      Width           =   5535
      Begin VB.Label lblBigDisplay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   720
         Left            =   360
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblsymbols 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   4815
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Insert string:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "All Symbols contained in:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   630
      Width           =   1815
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurrentLabel As Integer
Private noperline As Integer
Private linesout As Integer
Private gignore As Boolean
Private minuschars As Integer
Private fntFont As String
Private blnLoadedFonts As Boolean
Private Const BorderWidth As Integer = 100

Private Sub cboFonts_Click()
lblBigDisplay.Visible = False

Dim i As Integer ' Declare variable.
If lblsymbols(0).FontName <> cboFonts.Text Then
For i = 0 To lblsymbols.count - 1
lblsymbols(i).FontName = cboFonts.Text
Next
End If
If lblBigDisplay.FontName <> cboFonts.Text Then
lblBigDisplay.FontName = cboFonts.Text
End If
'setting fontname in txtInsert
txtInsert.Font = lblBigDisplay.FontName
lblBigDisplay.Visible = False

End Sub

Private Sub cboFonts_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdClose_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblBigDisplay.Visible = False
End Sub

Private Sub cmdCopy_Click()
On Error Resume Next
Clipboard.SetText txtInsert.Text
picHolder.SetFocus
End Sub

Private Sub cmdSymbols_Click(index As Integer)
'    On Error Resume Next
'    'frmMDI.ActiveForm.Text1.InsertContents SF_TEXT, cmdSymbols(Index).Caption
'
'    '...paste the Selected item.
'    frmMDI.ActiveForm.ActiveControl.SelText = ""    'This step is crucial!!! for undoing actions
'    ' Place the text from the Clipboard into the active control.
'    frmMDI.ActiveForm.ActiveControl.SelText = cmdSymbols(index).Caption
'    ' Set focus back to the active window
'    'frmMDI.ActiveForm.ActiveControl.SetFocus

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCopy_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblBigDisplay.Visible = False
End Sub

Private Sub cmdInsert_Click()
On Error Resume Next
picHolder.SetFocus

'...paste the Selected item
frmText.Text1.FontName = cboFonts.Text
frmText.Text1.SelText = ""    'This step is crucial!!! for undoing actions
' Place the text from the Clipboard into the active control.
frmText.Text1.SelText = txtInsert.Text
' Set focus back to the active window
'frmMDI.ActiveForm.ActiveControl.SetFocus
' closing the big display
lblBigDisplay.Visible = False
frmText.mnuUndo.Enabled = False
End Sub

Private Sub cmdInsert_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdInsert_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblBigDisplay.Visible = False
End Sub

Private Sub Form_Activate()
DropShadow cmdClose, Me
DropShadow cmdInsert, Me
DropShadow txtInsert, Me
DropShadow cboFonts, Me
DropShadow cmdCopy, Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyUp
If Shift = 0 Then
picHolder_KeyDown KeyCode, Shift
End If
KeyCode = 0
End Select
End Sub

Private Sub Form_Load()
Me.Icon = frmText.Icon

On Error Resume Next

blnLoadedFonts = False

fntFont = frmText.Text1.Font

lblMessage = "Symbols contained in: "
' set the big display to the same font
lblBigDisplay.Font = fntFont
noperline = 0
' set font and size
lblsymbols(0).Font = fntFont
FillSymbols (0)
gignore = True
VScroll1.Max = linesout
VScroll1.Min = 0
gignore = False
' Set the currently selected label to 0
CurrentLabel = 0

' adding one item named the active font name, just to show it
' then selecting it. The hole list vil be rebildt the first time
' the user click dhe dropdovn button
cboFonts.AddItem (fntFont), 0
cboFonts.ListIndex = 0


For x = 1 To Screen.FontCount
cboFonts.AddItem Screen.Fonts(x)
Next
cboFonts.RemoveItem (0)

For x = 0 To cboFonts.ListCount - 1
Exit For
Next

End Sub
Sub FillSymbols(ByVal startnumber As Integer)
gignore = False
' use minus chars to take away left co-or
minuschars = 1
' number of lines
numberoflines = 1
' hide the first symbol
lblsymbols(0).Left = -5000
' number of lines off screen
linesout = 0
' number of symbols per line
'noperline = 0
' Hide the picture box
picHolder.Visible = False
For i = 1 To 223
' Load the new symbol label
'On Error Resume Next
Load lblsymbols(i)
On Error GoTo 0
' change the current char - miss out
' the first 32 chars
currentchar = i + startnumber + 32
If currentchar > 255 Then Exit For
' Set caption to char
lblsymbols(i).Caption = Chr(currentchar)
' New left position
' (i - 1) [to allow left to start at 0
' - minuschars [to take away the previous
' symbols from prev. lines
' * (lblsymbols(i).Width - 12)
' [To move number from left plus
' line width
NewLeftPos = BorderWidth + ((i) - minuschars) * (lblsymbols(i).Width - 20)
' If the new left pos is bigger than
' the container width - new symbol
' then start a new line
If NewLeftPos > picHolder.Width - lblsymbols(i).Width Then
' Add the number of current symbols
' minus the one just created
minuschars = lblsymbols.count - 1
' Set the number per line (excluding
' current symbol, if it is not set
' -1 for currentsymbol
' -1 for first label which is not shown
If noperline = 0 Then noperline = lblsymbols.count - 2
' increment the number of lines
numberoflines = numberoflines + 1
' new top position (new line)
' lines - 1 [allow for top =0
' (lblsymbols(i).Height - 12)
' [number of lines - thick line
newtop = (numberoflines) * (lblsymbols(i).Height - 20)
' If the new top pos is greater than
' picHolder bottom line then increment
' lines out of screen
If newtop + lblsymbols(i).Height > picHolder.Height Then
linesout = linesout + 1
End If
' Set the new left to include the new
' minuschar value
'NewLeftPos = ((i) - minuschars) * (lblsymbols(i).Width - 12)
NewLeftPos = BorderWidth + (i - minuschars) * (lblsymbols(i).Width - 20)
End If
' Refresh pic1
'picHolder.Refresh
' set top pos of symbol
lblsymbols(i).Top = (numberoflines - 0.7) * (lblsymbols(i).Height - 20)
' set new left
lblsymbols(i).Left = NewLeftPos
' make is visible
lblsymbols(i).Visible = True
Next
' Show the picture again
picHolder.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblBigDisplay.Visible = False
End Sub

Private Sub Form_Paint()
DropShadow cmdClose, Me
DropShadow cmdInsert, Me
DropShadow txtInsert, Me
DropShadow cboFonts, Me
DropShadow cmdCopy, Me
End Sub

Private Sub lblBigDisplay_Click()
'    lblBigDisplay.Visible = False

End Sub

Private Sub lblBigDisplay_DblClick()
txtInsert.Text = txtInsert.Text & lblsymbols(CurrentLabel).Caption
End Sub

Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblBigDisplay.Visible = False
End Sub

Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblBigDisplay.Visible = False
End Sub

Private Sub lblStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblBigDisplay.Visible = False
End Sub

Private Sub lblsymbols_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo errHandler
lblBigDisplay.Left = lblsymbols(index).Left - ((lblBigDisplay.Width - lblsymbols(index).Width) / 2)
lblBigDisplay.Top = lblsymbols(index).Top - ((lblBigDisplay.Height - lblsymbols(index).Height) / 2)
lblBigDisplay.Caption = lblsymbols(index).Caption
lblBigDisplay.Visible = True
CurrentLabel = index
fred = lblsymbols(index).Caption
lblStatus.Caption = "Special Character " & Asc(fred)
errHandler:

End Sub

Private Sub picHolder_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not Shift = 0 Then Exit Sub
'    If KeyCode = vbKeyLeft And Not CurrentLabel = 1 Then
'        lblsymbols_Click (CurrentLabel - 1)
'    ElseIf KeyCode = vbKeyRight And Not CurrentLabel = lblsymbols.Count - 2 Then
'        lblsymbols_Click (CurrentLabel + 1)
'    ElseIf KeyCode = vbKeyUp And CurrentLabel > noperline Then
'        lblsymbols_Click (CurrentLabel - noperline)
'    ElseIf KeyCode = vbKeyDown And CurrentLabel < (lblsymbols.Count - 2 + noperline) Then
'        lblsymbols_Click (CurrentLabel + noperline)
'    End If
End Sub

Private Sub picHolder_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub txtInsert_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub txtInsert_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblBigDisplay.Visible = False
End Sub

Private Sub VScroll1_Change()
If Not gignore Then
MousePointer = vbHourglass
For Each Label In lblsymbols
If Not Label.index = 0 Then
Unload Label
End If
Next
charstart = VScroll1.Value * noperline
FillSymbols (charstart)
MousePointer = vbDefault
End If
lblBigDisplay.Visible = False
picHolder.SetFocus
End Sub

Private Sub VScroll1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub
