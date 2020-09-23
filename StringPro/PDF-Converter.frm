VERSION 5.00
Begin VB.Form frmConvertToPDF 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "String Processor"
   ClientHeight    =   5265
   ClientLeft      =   2625
   ClientTop       =   1395
   ClientWidth     =   6660
   Icon            =   "PDF-Converter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   4440
      Top             =   4680
   End
   Begin VB.CommandButton btnConvert 
      Caption         =   "&Convert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5190
      TabIndex        =   0
      Top             =   4200
      Width           =   1350
   End
   Begin VB.TextBox txtTitle 
      Height          =   360
      Left            =   1440
      TabIndex        =   6
      Top             =   3045
      Width           =   5100
   End
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5190
      TabIndex        =   1
      Top             =   4730
      Width           =   1350
   End
   Begin VB.ComboBox cmbPageSize 
      Height          =   315
      ItemData        =   "PDF-Converter.frx":1CFA
      Left            =   5040
      List            =   "PDF-Converter.frx":1D07
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3600
      Width           =   1545
   End
   Begin VB.ComboBox cmbFontSize 
      Height          =   315
      ItemData        =   "PDF-Converter.frx":1D2E
      Left            =   3240
      List            =   "PDF-Converter.frx":1D47
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3600
      Width           =   930
   End
   Begin VB.ComboBox cmbRotation 
      Height          =   315
      ItemData        =   "PDF-Converter.frx":1D74
      Left            =   4200
      List            =   "PDF-Converter.frx":1D84
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3600
      Width           =   810
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      ItemData        =   "PDF-Converter.frx":1D9D
      Left            =   1440
      List            =   "PDF-Converter.frx":1DAA
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Frame frmTitle 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   0
      TabIndex        =   18
      Top             =   -15
      Width           =   6720
      Begin VB.Image Image1 
         Height          =   750
         Left            =   120
         Picture         =   "PDF-Converter.frx":1DCB
         Stretch         =   -1  'True
         Top             =   120
         Width           =   5415
      End
      Begin VB.Image imgIcon 
         Height          =   750
         Left            =   5640
         Picture         =   "PDF-Converter.frx":357A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.TextBox txtOutputFile 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   -480
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtFilename 
      Height          =   255
      Left            =   6270
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtSubject 
      Height          =   360
      Left            =   1440
      TabIndex        =   5
      Top             =   2565
      Width           =   5100
   End
   Begin VB.TextBox txtKeywords 
      Height          =   360
      Left            =   1440
      TabIndex        =   4
      Top             =   2100
      Width           =   5100
   End
   Begin VB.TextBox txtCreator 
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   1635
      Width           =   5100
   End
   Begin VB.TextBox txtAuthor 
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   1185
      Width           =   5100
   End
   Begin VB.Label lblStringPro 
      BackStyle       =   0  'Transparent
      Caption         =   "String Processor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Label lblFonts 
      BackStyle       =   0  'Transparent
      Caption         =   "Fonts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblKeyword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Keywords"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   1110
   End
   Begin VB.Label lblCreator 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Creator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1110
   End
   Begin VB.Label lblAuthor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Author Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   -120
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmConvertToPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type OPENFILENAME
lStructSize As Long
hwndOwner As Long
hInstance As Long
lpstrFilter As String
lpstrCustomFilter As String
nMaxCustFilter As Long
nFilterIndex As Long
lpstrFile As String
nMaxFile As Long
lpstrFileTitle As String
nMaxFileTitle As Long
lpstrInitialDir As String
lpstrTitle As String
flags As Long
nFileOffset As Integer
nFileExtension As Integer
lpstrDefExt As String
lCustData As Long
lpfnHook As Long
lpTemplateName As String
End Type

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000 ' new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Dim Position As Long
Dim pageNo As Long
Dim lineNo As Long
Dim pageHeight As Long
Dim pageWidth As Long
Dim location(1 To 5000) As Long
Dim pageObj(1 To 5000) As Long
Dim lines As Long
Dim obj As Long
Dim Tpages As Long
Dim encoding As Long
Dim resources As Long
Dim pages As Variant
Dim author As String
Dim creator As String
Dim keywords As String
Dim subject As String
Dim Title As String
Dim BaseFont As String
Dim pointSize As Currency
Dim vertSpace As Currency
Dim rotate As Integer
Dim info As Long
Dim root As Long
Dim npagex As Double
Dim npagey As Long
Dim filetxt As String
Dim filepdf As String
Dim linelen As Long
Dim cache As String
Dim cmdline As String

Const AppName = "String Processor v 1.00"

Private Sub Form_Activate()
DropShadow txtAuthor, Me
DropShadow txtCreator, Me
DropShadow txtSubject, Me
DropShadow txtTitle, Me
DropShadow txtKeywords, Me
DropShadow cmbFont, Me
DropShadow cmbFontSize, Me
DropShadow cmbRotation, Me
DropShadow cmbPageSize, Me
DropShadow btnConvert, Me
DropShadow btnClose, Me
End Sub

Private Sub Form_Load()

Dim filename As String
On Local Error Resume Next
filename = Jimmy
If Len(filename) Then
txtFileName.Text = filename
filename = txtFileName.Text
txtOutputFile.Text = Left(filename, Len(filename) - 3) & "pdf"
End If

txtCreator.Text = AppName
cmbFont.ListIndex = 1
cmbFontSize.ListIndex = 1
cmbRotation.ListIndex = 0
cmbPageSize.ListIndex = 0

cmdline = LCase(Command)
If cmdline Like """*""" Then
cmdline = Mid(cmdline, 2, Len(cmdline) - 2)
End If

If FileExists(cmdline) Then
txtFileName.Text = cmdline
txtOutputFile.Text = Left(cmdline, Len(cmdline) - 4) & ".pdf"
btnConvert_Click
End If
End Sub

Private Sub Form_Paint()
DropShadow txtAuthor, Me
DropShadow txtCreator, Me
DropShadow txtSubject, Me
DropShadow txtTitle, Me
DropShadow txtKeywords, Me
DropShadow cmbFont, Me
DropShadow cmbFontSize, Me
DropShadow cmbRotation, Me
DropShadow cmbPageSize, Me
DropShadow btnConvert, Me
DropShadow btnClose, Me
End Sub

Private Sub Timer1_Timer()
Static Sarfraz%

If Sarfraz = 0 Then
lblStringPro.ForeColor = vbRed
Sarfraz = 1
Exit Sub

ElseIf Sarfraz = 1 Then
lblStringPro.ForeColor = vbCyan
Sarfraz = 2
Exit Sub

ElseIf Sarfraz = 2 Then
lblStringPro.ForeColor = vbMagenta
Sarfraz = 3
Exit Sub

ElseIf Sarfraz = 3 Then
lblStringPro.ForeColor = vbWhite
Sarfraz = 4
Exit Sub

ElseIf Sarfraz = 4 Then
lblStringPro.ForeColor = vbGreen
Sarfraz = 5
Exit Sub

ElseIf Sarfraz = 5 Then
lblStringPro.ForeColor = 8421631
Sarfraz = 6
Exit Sub

ElseIf Sarfraz = 6 Then
lblStringPro.ForeColor = vbBlack
Sarfraz = 7
Exit Sub

ElseIf Sarfraz = 7 Then
lblStringPro.ForeColor = vbYellow
Sarfraz = 0
Exit Sub

End If
End Sub

Private Sub txtAuthor_GotFocus()
txtAuthor.SelStart = 0
txtAuthor.SelLength = Len(txtAuthor.Text)
End Sub

Private Sub txtCreator_GotFocus()
txtCreator.SelStart = 0
txtCreator.SelLength = Len(txtCreator.Text)
End Sub

Private Sub txtSubject_GotFocus()
txtSubject.SelStart = 0
txtSubject.SelLength = Len(txtSubject.Text)
End Sub

Private Sub txtTitle_GotFocus()
txtTitle.SelStart = 0
txtTitle.SelLength = Len(txtTitle.Text)
End Sub

Private Sub txtKeywords_GotFocus()
txtKeywords.SelStart = 0
txtKeywords.SelLength = Len(txtKeywords.Text)
End Sub

Private Sub txtFilename_GotFocus()
txtFileName.SelStart = 0
txtFileName.SelLength = Len(txtFileName.Text)
End Sub

Private Sub txtOutputFile_GotFocus()
txtOutputFile.SelStart = 0
txtOutputFile.SelLength = Len(txtOutputFile.Text)
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnConvert_Click()
If txtFileName.Text <> "" And txtOutputFile.Text <> "" And frmText.Text1.Text <> "" And Jimmy <> "Untitled" Then
ConvertToPDF txtFileName.Text, txtOutputFile.Text, _
txtAuthor.Text, txtCreator.Text, txtKeywords.Text, _
txtSubject.Text, txtTitle.Text, _
cmbFont.Text, Val(cmbFontSize.Text), Val(cmbRotation.Text), _
Val(cmbPageSize.Text), Val(Right(cmbPageSize.Text, 3))
If Not FileExists(frmText.cd.filename) Then
Unload Me

ElseIf MsgBox("The file was successfully converted to PDF format!!" & vbCr & vbCr & "Do you want to open the generated PDF file?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
ShellExecute 0, vbNullString, txtOutputFile.Text, vbNullString, vbNullString, 1
End If

Else
MsgBox "No Text file loaded!!" & vbCr & "Please load a Text file in the StringPro first!", vbExclamation
Unload Me
End If
End Sub

Public Sub ConvertToPDF(filename As String, outputfile As String, _
Optional TextAuthor As String, Optional TextCreator As String, Optional TextKeywords As String, _
Optional TextSubject As String, Optional TextTitle As String, _
Optional FontName As String = "Courier", Optional FontSize As Integer = 10, Optional Rotation As Integer, _
Optional pwidth As Single = 8.5, Optional pheight As Single = 11)
On Error GoTo er

If Not FileExists(filename) Then
MsgBox "File '" & filename & "' does not exist.", vbExclamation
Exit Sub
ElseIf FileExists(outputfile) Then
Kill outputfile
End If

initialize FontName, FontSize, Rotation, pwidth, pheight

author = TextAuthor
creator = TextCreator
keywords = TextKeywords
subject = TextSubject
Title = TextTitle
filetxt = filename
filepdf = outputfile

Call WriteStart
Call WriteHead
Call WritePages
Call endpdf
Exit Sub
er:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub initialize(FontName As String, FontSize As Integer, Rotation As Integer, pwidth As Single, pheight As Single)
pageHeight = 72 * pheight
pageWidth = 72 * pwidth

BaseFont = FontName ' Courier, Times-Roman, Arial
pointSize = FontSize ' Font Size; Don't change it
vertSpace = FontSize * 1.2 ' Vertical spacing
rotate = Rotation ' degrees to rotate; try setting 90,180,etc
lines = (pageHeight - 72) / vertSpace ' no of lines on one page

Select Case LCase(FontName)
Case "courier": linelen = 1.5 * pageWidth / pointSize
Case "arial": linelen = 2 * pageWidth / pointSize
'Case "Times-Roman": linelen = 2.2 * pageWidth / pointSize
Case Else: linelen = 2.2 * pageWidth / pointSize
End Select

obj = 0
npagex = pageWidth / 2
npagey = 25
pageNo = 0
Position = 0
cache = ""
End Sub

Private Sub writepdf(stre As String, Optional flush As Boolean)
On Local Error Resume Next
Position = Position + Len(stre)
cache = cache & stre & vbCr
If Len(cache) > 32000 Or flush Then
Open filepdf For Append As #1
Print #1, cache;
Close #1
cache = ""
End If
End Sub

Private Sub WriteStart()
writepdf ("%PDF-1.2")
writepdf ("%âãÏÓ")
End Sub

Private Sub WriteHead()
Dim CreationDate As String
On Error GoTo er
CreationDate = "D:" & Format(Now, "YYYYMMDDHHNNSS")
obj = obj + 1
location(obj) = Position
info = obj

writepdf (obj & " 0 obj")
writepdf ("<<")
writepdf ("/Author (" & author & ")")
writepdf ("/CreationDate (" & CreationDate & ")")
writepdf ("/Creator (" & creator & ")")
writepdf ("/Producer (" & AppName & ")")
writepdf ("/Title (" & Title & ")")
writepdf ("/Subject (" & subject & ")")
writepdf ("/Keywords (" & keywords & ")")
writepdf (">>")
writepdf ("endobj")

obj = obj + 1
root = obj
obj = obj + 1
Tpages = obj
encoding = obj + 2
resources = obj + 3

obj = obj + 1
location(obj) = Position
writepdf (obj & " 0 obj")
writepdf ("<<")
writepdf ("/Type /Font")
writepdf ("/Subtype /Type1")
writepdf ("/Name /F1")
writepdf ("/Encoding " & encoding & " 0 R")
writepdf ("/BaseFont /" & BaseFont)
writepdf (">>")
writepdf ("endobj")

obj = obj + 1
location(obj) = Position
writepdf (obj & " 0 obj")
writepdf ("<<")
writepdf ("/Type /Encoding")
writepdf ("/BaseEncoding /WinAnsiEncoding")
writepdf (">>")
writepdf ("endobj")

obj = obj + 1
location(obj) = Position
writepdf (obj & " 0 obj")
writepdf ("<<")
writepdf ("  /Font << /F1 " & obj - 2 & " 0 R >>")
writepdf ("  /ProcSet [ /PDF /Text ]")
writepdf (">>")
writepdf ("endobj")
Exit Sub
er:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub WritePages()
Dim i As Integer
Dim line As String, tmpline As String, beginstream As String
On Error GoTo er
Open filetxt For Input As #2
beginstream = StartPage
lineNo = -1
Do Until EOF(2)
Line Input #2, line
lineNo = lineNo + 1

'page break
If lineNo >= lines Or InStr(line, Chr(12)) > 0 Then
writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
writepdf ("(" & pageNo & ") Tj")
writepdf ("/F1 " & pointSize & " Tf")
endpage (beginstream)
beginstream = StartPage
End If

line = ReplaceText(ReplaceText(line, "(", "\("), ")", "\)")
line = Trim(line)

If Len(line) > linelen Then

'word wrap
Do While Len(line) > linelen
tmpline = Left(line, linelen)
For i = Len(tmpline) To Len(tmpline) \ 2 Step -1
If InStr("*&^%$#,. ;<=>[])}!""", Mid(tmpline, i, 1)) Then
tmpline = Left(tmpline, i)
Exit For
End If
Next

line = Mid$(line, Len(tmpline) + 1)
writepdf ("T* (" & tmpline & vbCrLf & ") Tj")
lineNo = lineNo + 1

'page break
If lineNo >= lines Or InStr(line, Chr(12)) > 0 Then
writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
writepdf ("(" & pageNo & ") Tj")
writepdf ("/F1 " & pointSize & " Tf")
endpage (beginstream)
beginstream = StartPage
End If
Loop

lineNo = lineNo + 1
writepdf ("T* (" & line & vbCrLf & ") Tj")

Else

writepdf ("T* (" & line & vbCrLf & ") Tj")

End If
Loop
Close #2
writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
writepdf ("(" & pageNo & ") Tj")
writepdf ("/F1 " & pointSize & " Tf")
endpage (beginstream)
Exit Sub
er:
MsgBox Err.Description, vbCritical, "Error"
Close
End Sub

Private Function StartPage() As String
Dim strmpos As Long
On Error GoTo er
obj = obj + 1
location(obj) = Position
pageNo = pageNo + 1
pageObj(pageNo) = obj

writepdf (obj & " 0 obj")
writepdf ("<<")
writepdf ("/Type /Page")
writepdf ("/Parent " & Tpages & " 0 R")
writepdf ("/Resources " & resources & " 0 R")
obj = obj + 1
writepdf ("/Contents " & obj & " 0 R")
writepdf ("/Rotate " & rotate)
writepdf (">>")
writepdf ("endobj")

location(obj) = Position
writepdf (obj & " 0 obj")
writepdf ("<<")
writepdf ("/Length " & obj + 1 & " 0 R")
writepdf (">>")
writepdf ("stream")
strmpos = Position
writepdf ("BT")
writepdf ("/F1 " & pointSize & " Tf")
writepdf ("1 0 0 1 50 " & pageHeight - 40 & " Tm")
writepdf (vertSpace & " TL")

StartPage = strmpos
Exit Function
er:
MsgBox Err.Description, vbCritical, "Error"
End Function

Function endpage(streamstart As Long) As String
Dim streamEnd As Long
On Error GoTo er
writepdf ("ET")
streamEnd = Position
writepdf ("endstream")
writepdf ("endobj")
obj = obj + 1
location(obj) = Position
writepdf (obj & " 0 obj")
writepdf (streamEnd - streamstart)
writepdf "endobj"
lineNo = 0
Exit Function
er:
MsgBox Err.Description, vbCritical, "Error"
End Function

Sub endpdf()
Dim ty As String, i As Integer, xreF As Long
On Error GoTo er
location(root) = Position
writepdf (root & " 0 obj")
writepdf ("<<")
writepdf ("/Type /Catalog")
writepdf ("/Pages " & Tpages & " 0 R")
writepdf (">>")
writepdf ("endobj")
location(Tpages) = Position
writepdf (Tpages & " 0 obj")
writepdf ("<<")
writepdf ("/Type /Pages")
writepdf ("/Count " & pageNo)
writepdf ("/MediaBox [ 0 0 " & pageWidth & " " & pageHeight & " ]")
ty = ("/Kids [ ")
For i = 1 To pageNo
ty = ty & pageObj(i) & " 0 R "
Next i
ty = ty & "]"
writepdf (ty)
writepdf (">>")
writepdf ("endobj")
xreF = Position
writepdf ("0 " & obj + 1)
writepdf ("0000000000 65535 f ")
For i = 1 To obj
writepdf (Format(location(i), "0000000000") & " 00000 n ")
Next i
writepdf ("trailer")
writepdf ("<<")
writepdf ("/Size " & obj + 1)
writepdf ("/Root " & root & " 0 R")
writepdf ("/Info " & info & " 0 R")
writepdf (">>")
writepdf ("startxref")
writepdf (xreF)
writepdf "%%EOF", True
Exit Sub
er:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Public Function FileExists(ByVal filename As String) As Boolean
On Error Resume Next
FileExists = FileLen(filename) > 0
Err.Clear
End Function

Public Function ReplaceText(Text As String, TextToReplace As String, NewText As String) As String
Dim mtext As String, SpacePos As Long
mtext = Text
SpacePos = InStr(mtext, TextToReplace)
Do While SpacePos
mtext = Left(mtext, SpacePos - 1) & NewText & Mid(mtext, SpacePos + Len(TextToReplace))
SpacePos = InStr(SpacePos + Len(NewText), mtext, TextToReplace)
Loop
ReplaceText = mtext
End Function
