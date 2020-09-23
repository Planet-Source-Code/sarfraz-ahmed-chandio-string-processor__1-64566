VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "About "
   ClientHeight    =   6135
   ClientLeft      =   3300
   ClientTop       =   1185
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmAboutStringProcessor.frx":0000
   ScaleHeight     =   4234.485
   ScaleMode       =   0  'User
   ScaleWidth      =   5197.651
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   400
      Left            =   360
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   960
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   345
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Sys Info..."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   345
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4080
      TabIndex        =   12
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1/1/04"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3720
      TabIndex        =   1
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Registered To :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF80&
      BorderStyle     =   6  'Inside Solid
      Height          =   1935
      Left            =   120
      Top             =   3480
      Width           =   5295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:- sarfrazahmed_pk@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmAboutStringProcessor.frx":E4E9
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Port Qasim Housing Complex Karachi-48 (Zip: 75020) PAKISTAN."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4920
      Width           =   5415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Voice:- +92-0300-2409462"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmAboutStringProcessor.frx":E92B
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "If you have any suggestions,compliments or questions,you can contact me at following particulars:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   5055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "frmAboutStringProcessor.frx":EC35
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "sarfrazahmed_pk@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1200
      MouseIcon       =   "frmAboutStringProcessor.frx":15BCC
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:-If you detect any errors or bugs in the program, you can always contact me at my e-mail address."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.angelfire.com/ultra/sarfraz"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   435
      MouseIcon       =   "frmAboutStringProcessor.frx":1600E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "For other magnificent programs offered by the same developer,please visit the website:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function tapiRequestMakeCall Lib "tapi32.dll" _
         (ByVal stNumber As String, ByVal stDummy1 As String, _
         ByVal stDummy2 As String, ByVal stDummy3 As String) As Long

Dim AnimatedText As String
Dim a As Integer
Dim b As Integer
Dim UseSound As String

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = frmText.Icon

Label1.Visible = False
AnimatedText = Label1.Caption
a = Len(AnimatedText)
b = 1

'Show UserName
Dim sBuffer As String
Dim lSize As Long
sBuffer = Space$(255)
lSize = Len(sBuffer)
On Error Resume Next
Call GetUserName(sBuffer, lSize)
If lSize > 0 Then
Label8.Caption = Left$(sBuffer, lSize)
Else
Label8.Caption = ""
End If

End Sub

Public Sub StartSysInfo()
On Error GoTo SysInfoErr

Dim rc As Long
Dim SysInfoPath As String

' Try To Get System Info Program Path\Name From Registry...
If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
' Try To Get System Info Program Path Only From Registry...
ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
' Validate Existance Of Known 32 Bit File Version
If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
SysInfoPath = SysInfoPath & "\MSINFO32.EXE"

' Error - File Can Not Be Found...
Else
GoTo SysInfoErr
End If
' Error - Registry Entry Can Not Be Found...
Else
GoTo SysInfoErr
End If

Call Shell(SysInfoPath, vbNormalFocus)

Exit Sub
SysInfoErr:
MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
Dim i As Long                                           ' Loop Counter
Dim rc As Long                                          ' Return Code
Dim hKey As Long                                        ' Handle To An Open Registry Key
Dim hDepth As Long                                      '
Dim KeyValType As Long                                  ' Data Type Of A Registry Key
Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
'------------------------------------------------------------
' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
'------------------------------------------------------------
rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key

If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...

tmpVal = String$(1024, 0)                             ' Allocate Variable Space
KeyValSize = 1024                                       ' Mark Variable Size

'------------------------------------------------------------
' Retrieve Registry Key Value...
'------------------------------------------------------------
rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors

If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
Else                                                    ' WinNT Does NOT Null Terminate String...
tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
End If
'------------------------------------------------------------
' Determine Key Value Type For Conversion...
'------------------------------------------------------------
Select Case KeyValType                                  ' Search Data Types...
Case REG_SZ                                             ' String Registry Key Data Type
KeyVal = tmpVal                                     ' Copy String Value
Case REG_DWORD                                          ' Double Word Registry Key Data Type
For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
Next
KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
End Select

GetKeyValue = True                                      ' Return Success
rc = RegCloseKey(hKey)                                  ' Close Registry Key
Exit Function                                           ' Exit

GetKeyError:      ' Cleanup After An Error Has Occured...
KeyVal = ""                                             ' Set Return Val To Empty String
GetKeyValue = False                                     ' Return Failure
rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label12.ForeColor = vbWhite
Label11.ForeColor = vbWhite
Label11.FontBold = False
Label12.FontBold = False
End Sub

Private Sub Label1_Click()
On Error Resume Next
Shell "start http://www.angelfire.com/ultra/sarfraz", vbHide
End Sub

Private Sub Label11_Click()
Unload Me
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label11.ForeColor = vbYellow
Label11.FontBold = True
End Sub

Private Sub Label12_Click()
Call StartSysInfo
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label12.ForeColor = vbYellow
Label12.FontBold = True
End Sub

Private Sub Label3_Click()
On Error Resume Next
Shell "start mailto:sarfrazahmed_pk@yahoo.com", vbHide
End Sub

Private Sub Label5_Click()
Call Label3_Click
End Sub

Private Sub Label7_Click()
On Error GoTo Handler
DialNumber ("+92 (0300) 2409462")

Exit Sub
Handler:
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
UseSound = "Yes"

Label1.Visible = True
On Error Resume Next
Label1 = Left(AnimatedText, b)
b = b + 1

If UseSound = "Yes" And Label1 <> Left(AnimatedText, b) Then
Dim Play As String
On Error Resume Next
Play = sndPlaySound(App.path + "\TypingSound.wav", SND_ASYNC)
Else
End If

End Sub

Private Sub Timer2_Timer()
Static intCounter As Integer

If intCounter = 0 Then
Label8.ForeColor = vbYellow
intCounter = 1
Exit Sub

ElseIf intCounter = 1 Then
Label8.ForeColor = vbGreen
intCounter = 2
Exit Sub

ElseIf intCounter = 2 Then
Label8.ForeColor = vbMagenta
intCounter = 3
Exit Sub

ElseIf intCounter = 3 Then
Label8.ForeColor = vbBlue
intCounter = 4
Exit Sub

ElseIf intCounter = 4 Then
Label8.ForeColor = vbRed
intCounter = 5
Exit Sub

ElseIf intCounter = 5 Then
Label8.ForeColor = vbWhite
intCounter = 6
Exit Sub

ElseIf intCounter = 6 Then
Label8.ForeColor = 8421631
intCounter = 7
Exit Sub

ElseIf intCounter = 7 Then
Label8.ForeColor = vbCyan
intCounter = 8
Exit Sub

ElseIf intCounter = 8 Then
Label8.ForeColor = vbBlack
intCounter = 0
Exit Sub

End If
End Sub


'Following function receives a phone number to dial.
Function DialNumber(PhoneNumber)

Dim Msg As String
Dim retVal As Long

On Error Resume Next
' Send the telephone number to the modem.
retVal = tapiRequestMakeCall(PhoneNumber, "", "", "")
If retVal < 0 Then
Msg = "Unable to dial number " & PhoneNumber
End If

End Function

