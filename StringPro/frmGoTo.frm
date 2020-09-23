VERSION 5.00
Begin VB.Form frmGoTo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Goto Line"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "frmGoTo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&GO"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   900
      Width           =   1095
   End
   Begin VB.TextBox txtGo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Go to what:"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   75
      TabIndex        =   3
      Top             =   150
      Width           =   1260
      Begin VB.OptionButton GotoStart 
         Caption         =   "&Start"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   720
      End
      Begin VB.OptionButton GotoLine 
         Caption         =   "&Line"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton GoToEnd 
         Caption         =   "End"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter the line number to go to."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1635
      TabIndex        =   7
      Top             =   150
      Width           =   2655
   End
End
Attribute VB_Name = "frmGoTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdGo_Click()
On Error GoTo GoToError
Dim lngStart As Long

With frmText.Text1
If GotoLine Then 'If Go To line is checked
'Get pos of start of the line
lngStart = SendMessage(.hWnd, EM_LINEINDEX, txtGo.Text - 1, 0&)
If lngStart = -1 Then 'Invalid line number
MsgBox "Invalid Line Number!", vbExclamation, "Input Error"
txtGo.Text = ""
txtGo.SetFocus
Exit Sub
End If

.SelStart = lngStart 'Go To line
ElseIf GotoStart Then
.SelStart = 0
ElseIf GoToEnd Then
.SelStart = Len(.Text)
End If
.SetFocus
End With
GoToError:
If Err.Number = 13 Then
MsgBox "You can only type Numbers!", vbExclamation, "Input Error"
txtGo.Text = ""
txtGo.SetFocus
Exit Sub
End If
End Sub

Private Sub cmdGo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Activate()
DropShadow cmdGo, Me
DropShadow cmdClose, Me
DropShadow txtGo, Me
End Sub

Private Sub Form_Paint()
DropShadow cmdGo, Me
DropShadow cmdClose, Me
DropShadow txtGo, Me
End Sub

Private Sub GoToEnd_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub


Private Sub GotoLine_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub GotoLine_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtGo.SetFocus
End Sub

Private Sub GotoStart_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub txtGo_Change()
On Error Resume Next
If Not txtGo.Text = "" Or txtGo.Enabled = False Then
cmdGo.Enabled = True
ElseIf txtGo.Text = "" Then
cmdGo.Enabled = False
End If
End Sub

Private Sub GotoLine_Click()
txtGo.Enabled = True
txtGo_Change
End Sub

Private Sub GotoStart_Click()
txtGo.Enabled = False
txtGo_Change
End Sub

Private Sub GoToEnd_Click()
txtGo.Enabled = False
txtGo_Change
End Sub

Private Sub Form_Load()
Me.Icon = frmText.Icon
GotoLine_Click
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub txtGo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub
