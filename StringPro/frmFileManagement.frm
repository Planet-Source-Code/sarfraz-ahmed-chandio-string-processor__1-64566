VERSION 5.00
Begin VB.Form frmFileManagement 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Management"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Bro&wse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtCopyName 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   6255
   End
   Begin VB.CommandButton cmdSaveTo 
      Caption         =   "&BackUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Rename"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FileName"
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
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy To"
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmFileManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
On Error GoTo Handler
ShowFolders

If Right$(txtCopyName, 1) = "\" Then
Exit Sub
Else
txtCopyName = txtCopyName + "\"
End If

Exit Sub
Handler:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdBrowse_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdCopy_Click()

If Trim$(txtCopyName) = "" Then
MsgBox "No destination folder specified !!", vbExclamation
Exit Sub
End If

'Copy the file
On Error GoTo Handler
 If GetFTitle(txtCopyName) = "" Then
    FileCopy txtFileName, txtCopyName + GetFTitle(Jimmy)
    MsgBox "File was successfully copied to specified destination!", vbInformation
    Exit Sub
 ElseIf Right$(txtCopyName, 4) <> ".txt" Then
    FileCopy txtFileName, txtCopyName + ".txt"
    MsgBox "File was successfully copied to specified destination!", vbInformation
    Exit Sub
 Else
    FileCopy txtFileName, txtCopyName
    MsgBox "File was successfully copied to specified destination!", vbInformation
 End If

Exit Sub
Handler:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdCopy_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim Temp%

Temp = MsgBox("Are you sure you want to delete the opened file permanently?", vbQuestion + vbYesNo, "Delete File?")

If Temp = vbYes Then
On Error GoTo Handler
Close #1
Kill (Jimmy)
frmText.Text1.Text = ""
frmText.Text1.Tag = ""
Jimmy = "Untitled"
frmText.Caption = Jimmy & " - String Processor"
Sarfraz = False
frmText.cd.filename = ""
frmText.mnuUndo.Enabled = False
cmdCopy.Enabled = False
cmdBrowse.Enabled = False
cmdDelete.Enabled = False
cmdRename.Enabled = False
cmdSaveTo.Enabled = False
txtFileName = ""
txtCopyName = ""
Else
Exit Sub
End If

Exit Sub
Handler:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdExit_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdRename_Click()
txtFileName.SetFocus

If Not Right$(txtFileName, 4) = ".txt" And cmdRename.Caption = "Sa&ve" And cmdRename.Caption <> "&Rename" Then
MsgBox "Please specify the correct path and file name along with its extention", vbExclamation
Exit Sub
Else

On Error GoTo Handler
Name Jimmy As txtFileName.Text
frmText.cd.filename = txtFileName.Text
cmdRename.Caption = "&Rename"
Jimmy = txtFileName.Text
frmText.Caption = txtFileName + " - String Processor"
frmText.AddToHis
End If

Exit Sub
Handler:
MsgBox "Make sure that you specify the correct path!", vbCritical
End Sub

Private Sub cmdRename_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdSaveTo_Click()
On Error GoTo Handler
FileCopy Jimmy, Jimmy + ".bak"

If FileExists(Jimmy + ".bak") Then
MsgBox "File was successfully backedup!", vbInformation
End If

Exit Sub
Handler:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdSaveTo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Activate()
DropShadow cmdCopy, Me
DropShadow cmdBrowse, Me
DropShadow cmdRename, Me
DropShadow cmdDelete, Me
DropShadow cmdSaveTo, Me
DropShadow cmdExit, Me
DropShadow txtFileName, Me
DropShadow txtCopyName, Me
End Sub

Private Sub Form_Load()
Me.Icon = frmText.Icon
txtFileName.Text = Jimmy
txtCopyName.Text = StripPath(Jimmy)
End Sub

Private Sub Form_Paint()
DropShadow cmdCopy, Me
DropShadow cmdBrowse, Me
DropShadow cmdRename, Me
DropShadow cmdDelete, Me
DropShadow cmdSaveTo, Me
DropShadow cmdExit, Me
DropShadow txtFileName, Me
DropShadow txtCopyName, Me
End Sub

Private Sub txtCopyName_GotFocus()
txtCopyName.SelStart = 0
txtCopyName.SelLength = Len(txtCopyName.Text)
End Sub

Private Sub txtCopyName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub txtFileName_Change()
If txtFileName.Text <> Jimmy Then
cmdRename.Caption = "Sa&ve"
Else
cmdRename.Caption = "&Rename"
End If
End Sub

Private Sub txtFilename_GotFocus()
txtFileName.SelStart = 0
txtFileName.SelLength = Len(txtFileName.Text)
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub


'Shows the 'Browse in Folder' dialog
Private Sub ShowFolders()
Dim fld As Long
Dim Br As BROWSEINFO
Dim Path As String
Dim pos1 As Integer

Br.ulFlags = BIF_RETURNONLYFSDIRS
fld = SHBrowseForFolder(Br)
Path = Space(MAX_PATH)
If SHGetPathFromIDList(ByVal fld, ByVal Path) Then
txtCopyName = Path
End If

End Sub

