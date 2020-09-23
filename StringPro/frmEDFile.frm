VERSION 5.00
Begin VB.Form frmEDFile 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "String Processor"
   ClientHeight    =   3030
   ClientLeft      =   3675
   ClientTop       =   3015
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   1800
      Top             =   1320
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "D&one"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1850
      Width           =   2535
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncypt 
      Caption         =   "&Encrypt"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Super Encryption Algorithm"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   833
      TabIndex        =   4
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   120
      Picture         =   "frmEDFile.frx":0000
      Stretch         =   -1  'True
      Top             =   -840
      Width           =   4455
   End
End
Attribute VB_Name = "frmEDFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDecrypt_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdEncypt_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdQuit_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdDecrypt_Click()
Dim MyText As String
Dim TempEncKey As String
Dim tempChar As String
Dim EncStr As String
Dim EncKey As String
Dim EncLen As Integer
Dim EncPos As Integer
Dim EncKeyPos As Integer
Dim TA As Integer
Dim x As Long
Dim TB As Integer
Dim TC As Integer
Dim Temp As Integer

If Jimmy = "Untitled" Then
cmdDecrypt.Enabled = True
cmdEncypt.Enabled = True
MsgBox "No Text file loaded in StringPro!!", vbInformation
Exit Sub
End If
Temp = GetAttr(frmText.cd.filename)
If (Temp And vbReadOnly) <> 0 Then
MsgBox "Unable to decrypt the file!" & vbNewLine & "File seems to have Read-Only attribute set." & vbNewLine & "Remove the Read-Only attribute from the file and try again.", vbInformation
Exit Sub

Else
TempEncKey = InputBox("Enter the decryption key.This is the key that was typed for encryption.", "Decrypt")
If TempEncKey = "" Then Exit Sub
Me.Caption = "Decrypting......please wait"
EncStr = ""
EncPos = 1
EncKeyPos = 1
For x = 1 To Len(TempEncKey)
EncKey = EncKey & Asc(Mid$(TempEncKey, x, 1))
Next
EncLen = Len(EncKey)

MyText = frmText.Text1.Text

For x = 1 To Len(MyText) Step 2
TB = Asc(Mid$(EncKey, EncKeyPos, 1))
EncKeyPos = EncKeyPos + 1
If EncKeyPos > EncLen Then EncKeyPos = 1
tempChar = Mid$(MyText, x, 2)
TA = Val("&H" + tempChar)
TC = TB Xor TA
EncStr = EncStr & Chr$(TC)
Next
On Error GoTo Hell
MyText = EncStr
Me.Caption = "String Processor"
frmText.Text1.Text = MyText
frmText.mnuUndo.Enabled = False

cmdDecrypt.Enabled = False
cmdEncypt.Enabled = True
cmdEncypt.SetFocus

On Error GoTo Hell
Open Jimmy For Output As #1
Print #1, frmText.Text1.Text
frmText.StatusBar1.Panels(1).Text = "File Decrypted!!"
Sarfraz = False
Close #1
End If

Exit Sub
Hell:
HellError
Me.Caption = "String Processor"
End Sub

Private Sub cmdEncypt_Click()
Dim MyText As String
Dim TempEncKey As String
Dim tempChar As String
Dim EncStr As String
Dim EncKey As String
Dim EncLen As Integer
Dim EncPos As Integer
Dim EncKeyPos As Integer
Dim TA As Integer
Dim x As Long
Dim TB As Integer
Dim TC As Integer
Dim Temp As Integer

If Jimmy = "Untitled" Then
cmdEncypt.Enabled = True
cmdDecrypt.Enabled = True
MsgBox "No Text file loaded in StringPro!!", vbInformation
Exit Sub
End If
Temp = GetAttr(frmText.cd.filename)
If (Temp And vbReadOnly) <> 0 Then
MsgBox "Unable to encrypt the file!" & vbNewLine & "File seems to have Read-Only attribute set." & vbNewLine & "Remove the Read-Only attribute from the file and try again.", vbInformation
Exit Sub

Else
TempEncKey = InputBox("Enter the encryption key.This key will be vital for decrypting this text later.", "Encrypt")
If TempEncKey = "" Then Exit Sub
Me.Caption = "Encrypting......please wait"
EncStr = ""
EncPos = 1
EncKeyPos = 1

For x = 1 To Len(TempEncKey)
EncKey = EncKey & Asc(Mid$(TempEncKey, x, 1))
Next

EncLen = Len(EncKey)

MyText = frmText.Text1.Text

For x = 1 To Len(MyText)
TB = Asc(Mid$(EncKey, EncKeyPos, 1))
EncKeyPos = EncKeyPos + 1
If EncKeyPos > EncLen Then EncKeyPos = 1
TA = Asc(Mid$(MyText, x, 1))
TC = TB Xor TA
tempChar = Hex$(TC)
If Len(tempChar) < 2 Then tempChar = "0" & tempChar
EncStr = EncStr & tempChar
Next
On Error GoTo Hell
MyText = EncStr
Me.Caption = "String Processor"
frmText.Text1.Text = MyText
frmText.mnuUndo.Enabled = False


Open "temp.txt" For Output As #2
Close #2
i = MsgBox("Do you want to create a BACKUP file?", vbQuestion + vbYesNo, "Create Backup File?")
If i = vbYes Then FileCopy Jimmy, Jimmy + ".bak"

On Error GoTo HandleIt
Kill Jimmy
On Error GoTo HandleIt
Name "temp.txt" As Jimmy

On Error GoTo Hell
Open Jimmy For Output As #1
Print #1, frmText.Text1.Text
frmText.StatusBar1.Panels(1).Text = "File Encrypted!!"
Sarfraz = False
Close #1

cmdDecrypt.Enabled = True
cmdEncypt.Enabled = False
End If

Exit Sub
Hell:
HellError
Me.Caption = "String Processor"
HandleIt:
MsgBox Err.Description, vbCritical, "Error"
Me.Caption = "String Processor"
End Sub

Private Sub Form_Activate()
DropShadow cmdEncypt, Me
DropShadow cmdDecrypt, Me
DropShadow cmdQuit, Me
End Sub

Private Sub Form_Load()
Me.Icon = frmText.Icon
End Sub

Private Sub Form_Paint()
DropShadow cmdEncypt, Me
DropShadow cmdDecrypt, Me
DropShadow cmdQuit, Me
End Sub

Private Sub Timer1_Timer()
Label1.Move Label1.Left - 25
If Label1.Left < -Label1.Width Then Label1.Left = Image1.Width
End Sub

