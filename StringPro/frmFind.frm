VERSION 5.00
Begin VB.Form frmFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   1530
   ClientLeft      =   2700
   ClientTop       =   3585
   ClientWidth     =   6075
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtReplace 
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
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindAgain 
      Caption         =   "Find &Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtFind 
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
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Match Case"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Replace w&ith"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fi&nd what"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim compare As String

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then frmFind.Hide
End Sub

Private Sub cmdCancel_Click()
frmFind.Hide
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then frmFind.Hide
End Sub

Private Sub cmdFind_Click()

Position = 0
If Check1.Value = 1 Then
compare = vbBinaryCompare
Else
compare = vbTextCompare
End If
On Error GoTo Hell
Position = InStr(Position + 1, frmText.Text1.Text, txtFind.Text, compare)

If Position > 0 Then
frmText.Text1.SelStart = Position - 1
frmText.Text1.SelLength = Len(txtFind.Text)
cmdFindAgain.Enabled = True
If txtFind.Text <> "" And txtReplace.Text <> "" Then
cmdReplace.Enabled = True
cmdReplaceAll.Enabled = True
End If
frmText.SetFocus
Else
MsgBox "StringPro has finished searching the string  " + "  ' " + txtFind.Text + " '", vbInformation
cmdFindAgain.Enabled = False
cmdReplace.Enabled = False
cmdReplaceAll.Enabled = False
txtReplace.Enabled = False
End If

If cmdFind.Value = True Then
frmText.mnuFindAgain.Enabled = True
Else
frmText.mnuFindAgain.Enabled = False
End If

If Position > 0 Then
txtReplace.Enabled = True
Label2.Enabled = True
Else
txtReplace.Enabled = False
Label2.Enabled = False
End If

Exit Sub
Hell:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdFind_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then frmFind.Hide
End Sub

Private Sub cmdFindAgain_Click()
If Check1.Value = 1 Then
compare = vbBinaryCompare
Else
compare = vbTextCompare
End If
On Error GoTo Hell
Position = InStr(Position + 1, frmText.Text1.Text, txtFind.Text, compare)

If Position > 0 Then
frmText.Text1.SelStart = Position - 1
frmText.Text1.SelLength = Len(txtFind.Text)
cmdFindAgain.Enabled = True
If txtFind.Text <> "" And txtReplace.Text <> "" Then
cmdReplace.Enabled = True
cmdReplaceAll.Enabled = True
End If
frmText.SetFocus
Else
MsgBox "StringPro has finished searching the string  " + "  ' " + txtFind.Text + " '", vbInformation
frmText.SetFocus
cmdFindAgain.Enabled = False
cmdReplace.Enabled = False
cmdReplaceAll.Enabled = False
End If

If Position > 0 Then
txtReplace.Enabled = True
Label2.Enabled = True
Else
txtReplace.Enabled = False
Label2.Enabled = False
End If

Exit Sub
Hell:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdFindAgain_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then frmFind.Hide
End Sub

Private Sub cmdReplace_Click()
On Error GoTo Hell
If Position > 0 Then frmText.Text1.SelText = txtReplace.Text
Position = Position + Len(txtReplace.Text) - 1

If Check1.Value = 1 Then
compare = vbBinaryCompare
Else
compare = vbTextCompare
End If
Position = InStr(Position + 1, frmText.Text1.Text, txtFind.Text, compare)
frmText.mnuUndo.Enabled = False
If Position > 0 Then

frmText.Text1.SelStart = Position - 1
frmText.Text1.SelLength = Len(txtFind.Text)

cmdFindAgain.Enabled = True
If txtFind.Text <> "" And txtReplace.Text <> "" Then
cmdReplace.Enabled = True
cmdReplaceAll.Enabled = True
End If
frmText.SetFocus

Else
MsgBox "StringPro has finished searching the string  " + "  ' " + txtFind.Text + " '", vbInformation
frmText.SetFocus
cmdFindAgain.Enabled = False
cmdReplace.Enabled = False
cmdReplaceAll.Enabled = False
End If

If Position > 0 Then
txtReplace.Enabled = True
Label2.Enabled = True
Else
txtReplace.Enabled = False
Label2.Enabled = False
End If

Exit Sub
Hell:
HellError
End Sub

Private Sub cmdReplace_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then frmFind.Hide
End Sub

Private Sub cmdReplaceAll_Click()
Do While Position > 0
On Error GoTo Hell
If Position > 0 Then frmText.Text1.SelText = txtReplace.Text
Position = Position + Len(txtReplace.Text) - 1

If Check1.Value = 1 Then
compare = vbBinaryCompare
Else
compare = vbTextCompare
End If
Position = InStr(Position + 1, frmText.Text1.Text, txtFind.Text, compare)
frmText.mnuUndo.Enabled = False
If Position > 0 Then

frmText.Text1.SelStart = Position - 1
frmText.Text1.SelLength = Len(txtFind.Text)
cmdFindAgain.Enabled = True

If txtFind.Text <> "" And txtReplace.Text <> "" Then
cmdReplace.Enabled = True
cmdReplaceAll.Enabled = True
End If
frmText.SetFocus

Else
MsgBox "StringPro has finished searching the string  " + "  ' " + txtFind.Text + " '", vbInformation
frmText.SetFocus
cmdFindAgain.Enabled = False
cmdReplace.Enabled = False
cmdReplaceAll.Enabled = False
End If
Loop

If Position > 0 Then
txtReplace.Enabled = True
Label2.Enabled = True
Else
txtReplace.Enabled = False
Label2.Enabled = False
End If

Exit Sub
Hell:
HellError
End Sub

Private Sub cmdReplaceAll_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then frmFind.Hide
End Sub

Private Sub Form_Activate()
DropShadow txtFind, Me
DropShadow txtReplace, Me
DropShadow cmdFind, Me
DropShadow cmdFindAgain, Me
DropShadow cmdReplace, Me
DropShadow cmdReplaceAll, Me
DropShadow cmdCancel, Me
End Sub

Private Sub Form_Load()
Position = 0
txtReplace.Enabled = False
Label2.Enabled = False

If frmFind.txtFind = "" Then
cmdFindAgain.Enabled = False
End If
End Sub

Private Sub Form_Paint()
DropShadow txtFind, Me
DropShadow txtReplace, Me
DropShadow cmdFind, Me
DropShadow cmdFindAgain, Me
DropShadow cmdReplace, Me
DropShadow cmdReplaceAll, Me
DropShadow cmdCancel, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmFind.Hide
frmText.mnuFindAgain.Enabled = False
End Sub

Private Sub txtFind_Change()
If txtFind.Text <> "" Then
cmdFind.Enabled = True
cmdReplace.Enabled = False
cmdReplaceAll.Enabled = False
cmdFindAgain.Enabled = False
txtReplace.Enabled = False
Label2.Enabled = False
Else
cmdFind.Enabled = False
cmdReplace.Enabled = False
cmdReplaceAll.Enabled = False
cmdFindAgain.Enabled = False
frmText.mnuFindAgain.Enabled = False
End If

End Sub

Private Sub txtReplace_Change()
If txtFind.Text <> "" And txtReplace.Text <> "" And Position > 0 Then
cmdReplace.Enabled = True
cmdReplaceAll.Enabled = True
Else
cmdReplace.Enabled = False
cmdReplaceAll.Enabled = False
End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then frmFind.Hide
End Sub

Private Sub txtReplace_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then frmFind.Hide
End Sub
