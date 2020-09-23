VERSION 5.00
Begin VB.Form frmProperties 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Properties"
   ClientHeight    =   5700
   ClientLeft      =   3330
   ClientTop       =   1455
   ClientWidth     =   5400
   Icon            =   "Properties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1125
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   435
      Left            =   2120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Caption         =   "Attributes"
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
      Height          =   1155
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   5145
      Begin VB.CheckBox chkReadOnly 
         Caption         =   "&Read-Only"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1305
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "&System"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkArchive 
         Caption         =   "&Archive"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkHidden 
         Caption         =   "&Hidden"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dates"
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
      Height          =   1395
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   5145
      Begin VB.Label lblAccessed 
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblModified 
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lblCreated 
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Accessed:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Created:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
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
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5145
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Properties.frx":1CFA
         Top             =   960
         Width           =   3645
      End
      Begin VB.Label Label8 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "DOS Path:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblShortPath 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label lblShortName 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "DOS Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblSize 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   600
         Width           =   3645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Location:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblFileName 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   3645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String

Private Sub chkArchive_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkHidden_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkReadOnly_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkSystem_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdApply_Click()
On Error Resume Next
SetAttr (frmText.cd.filename), vbArchive * chkArchive.Value + _
vbSystem * chkSystem.Value + vbReadOnly * chkReadOnly.Value + vbHidden * chkHidden.Value
cmdOK.SetFocus
End Sub

Private Sub cmdApply_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdOK_Click()
cmdApply_Click
Unload Me
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
DropShadow cmdOK, Me
DropShadow cmdApply, Me
DropShadow cmdCancel, Me
End Sub

Private Sub Form_Load()
Dim Temp As Integer
Dim fil As File
Dim fso As New FileSystemObject
Dim FSize As Long

Screen.MousePointer = 11

If frmText.Text1.Text <> "" And Jimmy <> "Untitled" Then
On Error Resume Next
Set fil = fso.GetFile(Jimmy)

FSize = CStr(fil.Size)

On Error Resume Next
Temp = GetAttr(frmText.cd.filename)
If (Temp And vbArchive) <> 0 Then chkArchive.Value = 1
If (Temp And vbReadOnly) <> 0 Then chkReadOnly.Value = 1
If (Temp And vbHidden) <> 0 Then chkHidden.Value = 1
If (Temp And vbSystem) <> 0 Then chkSystem.Value = 1

On Error Resume Next
lblFileName = fil.Name
lblFileName.ToolTipText = frmText.cd.filename
lblSize = FormatSize(FSize)
lblShortName = fil.ShortName
lblCreated = fil.DateCreated
lblShortPath = fil.ShortPath
lblModified = fil.DateLastModified
lblAccessed = fil.DateLastAccessed
Text1 = fil.path
Text1.ToolTipText = frmText.cd.filename
lblShortPath.ToolTipText = fil.ShortPath
Else
End If

Set fil = Nothing

Me.Show
Exit Sub
End Sub

'Shows file size in different formats such as KB
Public Function FormatSize(ByVal Amount As Long) As String
Dim Buffer As String
Dim Result As String

Buffer = Space$(255) 'Fill buffer
Result = StrFormatByteSize(Amount, Buffer, Len(Buffer)) 'Format file size
If InStr(Result, vbNullChar) > 1 Then
FormatSize = Left$(Result, InStr(Result, vbNullChar) - 1)
End If
End Function

Private Sub Form_Paint()
DropShadow cmdOK, Me
DropShadow cmdApply, Me
DropShadow cmdCancel, Me
End Sub
