VERSION 5.00
Begin VB.Form frmCompare 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compare Text"
   ClientHeight    =   7425
   ClientLeft      =   2175
   ClientTop       =   420
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3180
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H005A3F2E&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Close #1
frmText.cd.filename = Jimmy
Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Call Command1_Click
End Sub

Private Sub Form_Activate()
DropShadow Command1, Me
End Sub

Private Sub Form_Load()
Me.Icon = frmText.Icon
End Sub

Private Sub Form_Paint()
DropShadow Command1, Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Command1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command1_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Call Command1_Click
End Sub
