VERSION 5.00
Begin VB.Form frmWordCount 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Word Count"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   2655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   780
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblChars 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblWords 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblSpaces 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblLines 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Characters:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Spaces:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Lines:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Words:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmWordCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Activate()
DropShadow Command1, Me
End Sub

Private Sub Form_Load()
Me.Icon = frmText.Icon

Dim WordCount$, ChrCount
Dim a() As String
Dim b() As String

'Get Words
a() = Split(frmText.Text1.Text, " ") 'Split text to " "
WordCount = UBound(a)
For i = 0 To UBound(a)
If a(i) = "" Then
WordCount = WordCount - 1
End If
Next
b() = Split(frmText.Text1.Text, Chr$(10))
WordCount = WordCount + UBound(b)
For i = 0 To UBound(b)
If b(i) = "" Then
WordCount = WordCount - 1
End If
Next

'Get Characters
ChrCount = SendMessageLong(frmText.Text1.hWnd, WM_GETTEXTLENGTH, 0, 0)
 

If WordCount = -2 Then WordCount = -1
lblLines.Caption = Format(SendMessage(frmText.Text1.hWnd, EM_GETLINECOUNT, 0, 0&), "###,###,###,###")
lblSpaces.Caption = Format(CountSpaces(frmText.Text1.Text), "###,###,###,###,###")
lblWords.Caption = Format(WordCount + 1, "###,###,###,###")
lblChars.Caption = Format(ChrCount, "###,###,###,###,###")

Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Paint()
DropShadow Command1, Me
End Sub
