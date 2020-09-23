VERSION 5.00
Begin VB.Form frmScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtScreen 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Me.Caption = frmText.Caption
Form_Resize
txtScreen.ForeColor = frmText.Text1.ForeColor
txtScreen.BackColor = frmText.Text1.BackColor
txtScreen.FontBold = frmText.Text1.FontBold
txtScreen.FontItalic = frmText.Text1.FontItalic
txtScreen.FontName = frmText.Text1.FontName
txtScreen.FontSize = frmText.Text1.FontSize
txtScreen.FontStrikethru = frmText.Text1.FontStrikethru
txtScreen.FontUnderline = frmText.Text1.FontUnderline

End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Left = -30
Me.Top = -30
Me.Height = Screen.Height + 30
Me.Width = Screen.Width + 30
txtScreen.Width = Me.Width - 30
txtScreen.Height = Me.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
If txtScreen <> frmText.Text1 Then
On Error GoTo Handler
frmText.Text1 = txtScreen
txtScreen = ""
frmText.mnuUndo.Enabled = False
Exit Sub
Else
Unload Me
End If

Handler:
If Err.Description = "Out of memory" Or Err.Number = 7 Then
MsgBox "File is too large!!", vbExclamation
Else
End If
End Sub

Private Sub txtScreen_Change()
If frmText.mnuTypingSound.Checked = False Then
Exit Sub
Else
Dim Play As String
On Error Resume Next
Play = sndPlaySound(App.path + "\TypingSound.wav", SND_ASYNC)
End If
End Sub

Private Sub txtScreen_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub
