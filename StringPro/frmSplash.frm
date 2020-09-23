VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "String Processor v1.00"
   ClientHeight    =   4365
   ClientLeft      =   3210
   ClientTop       =   2025
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   2160
      Top             =   1920
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   240
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sarfraz Ahmed Chandio"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.angelfire.com/ultra/sarfraz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   2520
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Runs"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D39E5C&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   420
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown  Company"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown User"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This product is liscened to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   0
      Picture         =   "frmSplash.frx":1CFA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   -360
      Picture         =   "frmSplash.frx":DCDC
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim TimesRun
Dim keyHandle As Long
Dim retVal As Long, Read As Boolean
Dim strData As String

'Show UserName and CompanyName
On Error Resume Next
retVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", 0, 0, keyHandle)
If retVal = ERROR_SUCCESS Then
strData = Space(300)
retVal = 2
retVal = RegQueryValueEx(keyHandle, "RegisteredOwner", 0, REG_DWORD, ByVal strData, 300)
If retVal = ERROR_SUCCESS Then
Label2.Caption = strData
strData = Space(300)
retVal = RegQueryValueEx(keyHandle, "RegisteredOrganization", 0, REG_DWORD, ByVal strData, 300)
On Error Resume Next
Label3.Caption = strData
Read = True
End If
End If

On Error Resume Next
RegCloseKey (keyHandle)
If Read = False Then
retVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", 0, 0, keyHandle)

strData = Space(300)
retVal = RegQueryValueEx(keyHandle, "RegisteredOwner", 0, REG_DWORD, ByVal strData, 300)
Label2.Caption = strData

strData = Space(300)
retVal = RegQueryValueEx(keyHandle, "RegisteredOrganization", 0, REG_DWORD, ByVal strData, 300)
On Error Resume Next
Label5.Caption = strData
End If
RegCloseKey (keyHandle)


'Counts number of times app has run
On Error Resume Next
SaveSetting "StringPro", "StartupForm", "TimesRun", GetSetting("StringPro", "StartupForm", "TimesRun", 0) + 1
TimesRun = GetSetting("StringPro", "StartupForm", "TimesRun")
Label7.Caption = TimesRun

End Sub

Private Sub Timer1_Timer()
Load frmText
Unload Me
frmText.Show
End Sub
