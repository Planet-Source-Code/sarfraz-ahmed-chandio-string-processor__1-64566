VERSION 5.00
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6120
   ClientLeft      =   3165
   ClientTop       =   1425
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2333
      TabIndex        =   16
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCleanReg 
      Caption         =   "&Restore Defaults"
      Height          =   375
      Left            =   1913
      TabIndex        =   14
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Others"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Width           =   5535
      Begin VB.CheckBox chkHtml 
         Caption         =   "&HTML Menu Visible"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   440
         Width           =   1815
      End
      Begin VB.CheckBox chkPlay 
         Caption         =   "&Play Typing Sound"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   440
         Width           =   1815
      End
      Begin VB.CheckBox chkToolbar 
         Caption         =   "&Toolbar"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   440
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "StartUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   5535
      Begin VB.CheckBox chkSplash 
         Caption         =   "&Show Splash Screen"
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cboStartup 
         Height          =   315
         ItemData        =   "frmOptions.frx":0000
         Left            =   240
         List            =   "frmOptions.frx":000D
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Select StringPro StartUp Position"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton cmdFC 
         Caption         =   "Selec&t"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox chkBC 
         Caption         =   "Page C&olor"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkFC 
         Caption         =   "T&ext Color"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdBC 
         Caption         =   "Se&lect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboFontName 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "&Underline"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "&Italic"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "&Bold"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Font Size"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Font Name"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   420
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFontName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cboFontSize_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cboStartup_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkBC_Click()
If chkBC.Value = 1 Then
cmdBC.Enabled = True
Else
cmdBC.Enabled = False
End If
End Sub

Private Sub chkBold_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkFC_Click()
If chkFC.Value = 1 Then
cmdFC.Enabled = True
Else
cmdFC.Enabled = False
End If
End Sub

Private Sub chkHtml_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkItalic_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkPlay_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkSplash_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkToolbar_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkUnderline_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdApply_Click()

frmText.Text1.FontName = cboFontName.Text
frmText.Text1.FontSize = cboFontSize.Text

If chkBold.Value = 1 Then frmText.Text1.FontBold = True
If chkItalic.Value = 1 Then frmText.Text1.FontItalic = True
If chkUnderline.Value = 1 Then frmText.Text1.FontUnderline = True
If chkToolbar.Value = 1 Then frmText.Toolbar1.Visible = True
If chkHtml.Value = 1 Then frmText.mnuHtml.Visible = True

If chkFC.Value = 0 Then frmText.Text1.ForeColor = vbBlack
If chkBC.Value = 0 Then frmText.Text1.BackColor = vbWhite
If chkBold.Value = 0 Then frmText.Text1.FontBold = False
If chkItalic.Value = 0 Then frmText.Text1.FontItalic = False
If chkUnderline.Value = 0 Then frmText.Text1.FontUnderline = False
If chkToolbar.Value = 0 Then frmText.Toolbar1.Visible = False
If chkHtml.Value = 0 Then frmText.mnuHtml.Visible = False

If frmText.Toolbar1.Visible Then
frmText.Text1.Top = 445
frmText.Text1.Height = 7575
frmText.mnuToolbar.Checked = True
Else
frmText.Text1.Top = 1
frmText.Text1.Height = 7800
frmText.mnuToolbar.Checked = False
End If

On Error GoTo Handler
SaveSetting App.EXEName, "FontNameSize", "FontName", cboFontName.Text
SaveSetting App.EXEName, "FontNameSize", "FontSize", cboFontSize.Text
SaveSetting App.EXEName, "FontStyle", "FontBold", chkBold.Value
SaveSetting App.EXEName, "FontStyle", "FontItalic", chkItalic.Value
SaveSetting App.EXEName, "FontStyle", "FontUnderline", chkUnderline.Value
SaveSetting App.EXEName, "StartupPos", "Position", cboStartup.ListIndex
SaveSetting App.EXEName, "Toolbar", "Visible", chkToolbar.Value
SaveSetting App.EXEName, "StartupPos", "Splash", chkSplash.Value
SaveSetting App.EXEName, "TypingSound", "Play", chkPlay.Value
SaveSetting App.EXEName, "Color", "ForeColor", chkFC.Value
SaveSetting App.EXEName, "Color", "BackColor", chkBC.Value
SaveSetting App.EXEName, "Color", "TextForeColor", frmText.Text1.ForeColor
SaveSetting App.EXEName, "Color", "TextBackColor", frmText.Text1.BackColor
SaveSetting App.EXEName, "HtmlMenu", "visible", chkHtml.Value

cmdOk.SetFocus

Exit Sub
Handler:
frmText.Text1.FontName = "Verdana"
frmText.Text1.FontSize = 11
frmText.Text1.FontBold = False
frmText.Text1.FontItalic = False
frmText.Text1.FontUnderline = False
frmText.WindowState = 2
frmText.Toolbar1.Visible = True
frmText.mnuToolbar.Checked = True
frmText.mnuTypingSound.Checked = False
frmText.Text1.ForeColor = vbBlack
frmText.Text1.BackColor = vbWhite
frmText.mnuHtml.Visible = False
End Sub

Private Sub cmdApply_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdBC_Click()
On Error GoTo Handler
frmText.cd.CancelError = True
frmText.cd.ShowColor
frmText.Text1.BackColor = frmText.cd.color

Exit Sub
Handler:
Dim a
a = GetSetting(App.EXEName, "Color", "TextBackColor", "")
If a = "" Then
frmText.Text1.BackColor = vbWhite
Else
frmText.Text1.BackColor = GetSetting(App.EXEName, "Color", "TextBackColor", "")
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdCleanReg_Click()

On Error Resume Next
DeleteSetting App.EXEName, "StartupPos"
DeleteSetting App.EXEName, "FontNameSize"
DeleteSetting App.EXEName, "FontStyle"
DeleteSetting App.EXEName, "Toolbar"
DeleteSetting App.EXEName, "TypingSound"
DeleteSetting App.EXEName, "file"
DeleteSetting App.EXEName, "LastFile"
DeleteSetting App.EXEName, "save"
DeleteSetting App.EXEName, "Color"
DeleteSetting App.EXEName, "HtmlMenu"
DeleteSetting App.EXEName, "StartupPos"

MsgBox "The default settings have successfully been restored and will be applied the next time you run the StringPro.", vbExclamation
Unload Me
End Sub

Private Sub cmdCleanReg_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdFC_Click()
On Error GoTo Handler
frmText.cd.CancelError = True
frmText.cd.ShowColor
frmText.Text1.ForeColor = frmText.cd.color

Exit Sub
Handler:
Dim b
b = GetSetting(App.EXEName, "Color", "TextForeColor", "")
If b = "" Then
frmText.Text1.ForeColor = vbBlack
Else
frmText.Text1.ForeColor = GetSetting(App.EXEName, "Color", "TextForeColor", "")
End If
End Sub

Private Sub cmdOK_Click()
cmdApply_Click
Unload Me
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Activate()
DropShadow cmdApply, Me
DropShadow cmdOk, Me
DropShadow cmdCleanReg, Me
DropShadow cmdCancel, Me
End Sub

Private Sub Form_Load()
Me.Icon = frmText.Icon

Dim x%
'Get FontName
cboFontName = frmText.Text1.FontName

For x = 1 To Screen.FontCount
cboFontName.AddItem Screen.Fonts(x)
Next
cboFontName.RemoveItem (0)

For x = 0 To cboFontName.ListCount - 1
Exit For
Next

'Get FontSize
cboFontSize = frmText.Text1.FontSize

For x = 8 To 72
cboFontSize.AddItem Str$(x)
Next

For x = 0 To cboFontSize.ListCount - 1
Exit For
Next


Dim Values(11)

Values(0) = GetSetting(App.EXEName, "HtmlMenu", "visible", "")
If Values(0) = "" Then
chkHtml.Value = 0
Else
chkHtml.Value = GetSetting(App.EXEName, "HtmlMenu", "visible", "")
End If

Values(1) = GetSetting(App.EXEName, "FontNameSize", "FontName", "")
If Values(1) = "" Then
cboFontName.Text = "Verdana"
Else
cboFontName = GetSetting(App.EXEName, "FontNameSize", "FontName", "")
End If

Values(2) = GetSetting(App.EXEName, "FontNameSize", "FontSize", "")
If Values(2) = "" Then
cboFontSize = 11
Else
cboFontSize = GetSetting(App.EXEName, "FontNameSize", "FontSize", "")
End If

Values(3) = GetSetting(App.EXEName, "Toolbar", "Visible", "")
If Values(3) = "" Then
chkToolbar.Value = 1
Else
chkToolbar.Value = GetSetting(App.EXEName, "Toolbar", "Visible", "")
End If

Values(4) = GetSetting(App.EXEName, "StartupPos", "Position", "")
If Values(4) = "" Then
cboStartup = "Maximized"
Else
If Values(4) = 0 Then cboStartup = "Normal"
If Values(4) = 2 Then cboStartup = "Maximized"
If Values(4) = 1 Then cboStartup = "Minimized"
End If

Values(5) = GetSetting(App.EXEName, "FontStyle", "FontBold", "")
If Values(5) = "" Then
chkBold.Value = 0
Else
If Values(5) = "True" Then chkBold.Value = 1
chkBold.Value = GetSetting(App.EXEName, "FontStyle", "FontBold", "")
End If


Values(6) = GetSetting(App.EXEName, "FontStyle", "FontItalic", "")
If Values(6) = "" Then
chkItalic.Value = 0
Else
If Values(6) = "True" Then chkItalic.Value = 1
chkItalic.Value = GetSetting(App.EXEName, "FontStyle", "FontItalic", "")
End If


Values(7) = GetSetting(App.EXEName, "FontStyle", "FontUnderline", "")
If Values(7) = "" Then
chkUnderline.Value = 0
Else
If Values(7) = "True" Then chkUnderline.Value = 1
chkUnderline.Value = GetSetting(App.EXEName, "FontStyle", "FontUnderline", "")
End If

Values(8) = GetSetting(App.EXEName, "StartupPos", "Splash", "")
If Values(8) = "" Then
chkSplash.Value = 1
Else
'On Error Resume Next
chkSplash.Value = GetSetting(App.EXEName, "StartupPos", "Splash", "")
End If

Values(9) = GetSetting(App.EXEName, "TypingSound", "Play", "")
If Values(9) = "" Then
chkPlay.Value = 0
Else
'On Error Resume Next
chkPlay.Value = GetSetting(App.EXEName, "TypingSound", "Play", "")
End If

Values(10) = GetSetting(App.EXEName, "Color", "ForeColor", "")
If Values(10) = 0 Then
chkFC.Value = 0
Else
On Error Resume Next
chkFC.Value = GetSetting(App.EXEName, "Color", "ForeColor", "")
End If

On Error Resume Next
Values(11) = GetSetting(App.EXEName, "Color", "BackColor", "")
If Values(11) = 0 Then
chkBC.Value = 0
Else
chkBC.Value = GetSetting(App.EXEName, "Color", "BackColor", "")
End If

End Sub

Private Sub Form_Paint()
DropShadow cmdApply, Me
DropShadow cmdOk, Me
DropShadow cmdCleanReg, Me
DropShadow cmdCancel, Me
End Sub
