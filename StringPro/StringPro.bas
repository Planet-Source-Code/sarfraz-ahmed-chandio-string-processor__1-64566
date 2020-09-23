Attribute VB_Name = "StringPro"
Option Explicit
Option Compare Text
Option Base 0

Public Position As Double
Public i As Long
Public Jimmy As String
Public Sarfraz As Boolean
Public Const EM_LINEFROMCHAR = &HC9
Public Const SND_ASYNC = &H1
Public Const WM_GETTEXTLENGTH = &HE


Const FirstLines = 15 'Top shadow roundless      *
Const EndLines = 15   'Bottom shadow roundless   *
Const Measure = 15    'Form measure (p.e. twips) *
Const Desp = 65       'Shadow desp               *
Const Band = 60       'Shadow brightness         *



'For 'Browse in Folder' dialog

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

'Folders Show
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const MAX_PATH = 260
Public Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type


Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'Declarations for UserName and Company
Public Const ERROR_SUCCESS = 0&
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_DWORD = 4                      ' 32-bit number

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" _
Alias "RegOpenKeyExA" (ByVal hKey As Long, _
ByVal lpSubKey As String, ByVal ulOptions As Long, _
ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
Alias "RegQueryValueExA" (ByVal hKey As Long, _
ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'Declarations for User-defined PopUpMenu
Public Const GWL_WNDPROC = (-4)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_CONTEXTMENU = &H7B

Public origWndProc As Long



'The following two lines are what are responsible for
'Undoing.
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_UNDO = &HC7
Public Const EM_LINEINDEX = &HBB


'This function counts number of lines.
Public Declare Function SendMessageByVal Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINELENGTH = &HC1


'Get UserName
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long


Public Sub Main()

If frmOptions.chkSplash.Value = 1 Then
 frmSplash.Show
Else
 Load frmText
 frmText.Show
End If

End Sub



'Get Current Directory
Public Function GetDirectory(filename As String) As String
Dim strpos As Long
Dim CurPos As Integer
CurPos = 1

GetDirectory = ""
If InStr(CurPos, filename, "\", vbTextCompare) <= 0 Then
GetDirectory = ""
Exit Function
End If
Do While True
strpos = InStr(CurPos, filename, "\", vbTextCompare)
If strpos <> 0 Then
CurPos = strpos + 1
Else
GetDirectory = Left$(filename, CurPos - 1)
Exit Do
End If
Loop
End Function


'True if text is selected
Public Function TextSelected()
TextSelected = frmText.Text1.SelText <> ""
End Function


'True if text is not selected
Function TextNotSelected() As Boolean
If Len(frmText.Text1.SelText) = 0 Then
MsgBox "Text Not Selected!", vbExclamation
TextNotSelected = True
End If
End Function

'Get FileName from the Path
Public Function GetFTitle(strFilename As String)
On Error Resume Next
Dim cbBuf As String
    
cbBuf = String(250, vbNullChar) 'Fill buffer with null chars
GetFileTitle strFilename, cbBuf, Len(cbBuf) 'Get file title
GetFTitle = Left(cbBuf, InStr(1, cbBuf, vbNullChar) - 1) 'Extract file title from buffer

End Function

'Unload all forms
Function UnloadForms()
Dim Form As Form
For Each Form In Forms
Unload Form
Set Form = Nothing
Next Form
End Function

'Checks whether a file exists
Function FileExists(ByVal strFilePath As String) As Boolean
strFilePath = Trim(strFilePath)
If strFilePath = "" Then Exit Function
If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

'This function converts numbers to their corresponding
'roman values
Public Function NumericToRoman(ByVal Value As Long) As String

Dim iPos As Integer, sBuffer As String, iReference As Integer
Dim sLowChar As String, sMidChar As String, sHighChar As String

On Error Resume Next
sBuffer = String$(Value \ 1000, "M")
Value = Value Mod 1000

iReference = 100
Do Until iReference = 0
If iReference = 100 Then
sHighChar = "M"
sMidChar = "D"
sLowChar = "C"
ElseIf iReference = 10 Then
sHighChar = "C"
sMidChar = "L"
sLowChar = "X"
Else
sHighChar = "X"
sMidChar = "V"
sLowChar = "I"
End If
iPos = Value \ iReference
If (iPos > 0) And (iPos < 4) Then
sBuffer = sBuffer & String$(iPos, sLowChar)
ElseIf iPos = 4 Then
sBuffer = sBuffer & sLowChar & sMidChar
ElseIf iPos = 5 Then
sBuffer = sBuffer & sMidChar
ElseIf (iPos > 5) And (iPos < 9) Then
sBuffer = sBuffer & sMidChar & String$(iPos - 5, sLowChar)
ElseIf iPos = 9 Then
sBuffer = sBuffer & sLowChar & sHighChar
End If
Value = Value - iReference * iPos
iReference = iReference \ 10
Loop
NumericToRoman = sBuffer

End Function


'Show the error message
Public Function HellError()
If Err.Number = 7 Or Err.Description = "Out of Memory" Then
MsgBox "File is too large!!", vbExclamation, "File Size Large"
frmText.mnuUndo.Enabled = False
Else
MsgBox Err.Description, vbCritical, "Error"
End If
End Function

'Disables the TextBox default PopUpMenu and enables the
'user-defined one.
Public Sub SetHook(hWnd, bSet As Boolean)
If bSet Then
On Error Resume Next
origWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf AppWndProc)
ElseIf origWndProc Then
Dim lRet As Long
lRet = SetWindowLong(hWnd, GWL_WNDPROC, origWndProc)
End If
End Sub

Public Function AppWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case Msg
Case WM_CONTEXTMENU
frmText.PopupMenu frmText.mnuEdit
AppWndProc = 0
Exit Function
End Select
On Error Resume Next
AppWndProc = CallWindowProc(origWndProc, hWnd, Msg, wParam, lParam)
End Function


Public Function TrackChanges(input1 As String, input2 As String) As String
'This function tracks differences between two similar
'strings for example "I AM LIVING IN ATHENS.GREETINGS"
'and "I AM LIVING IN LONDON.GREETINGS" will return
'"LONDON"

Dim pastedstr, spos, fpos, floko, fluki, epos, flupos

pastedstr = ""

spos = 1
fpos = 1
Do While spos < Len(input1)
floko = Mid(input1, spos, 1)
fluki = Mid(input2, fpos, 1)
If floko = fluki Then
spos = spos + 1
fpos = fpos + 1
Else
Exit Do
End If
Loop

epos = Len(input1)
flupos = Len(input2)

Do While epos > spos
floko = Mid(input1, epos, 1)
fluki = Mid(input2, flupos, 1)
If fluki = floko Then
epos = epos - 1
flupos = flupos - 1
Else
Exit Do
End If
Loop

If flupos = 0 Then
TrackChanges = input1
Else
TrackChanges = Mid(input2, spos, flupos - spos + 1)
End If

End Function


' This routine also works with open files
' and raises an error if the file doesn't exist.
Function GetAttribute(filename As String) As String
Dim Result As String, attr As Long

attr = GetAttr(filename)
' GetAttr also works with directories.
If attr And vbDirectory Then Result = Result & " Directory"
If attr And vbReadOnly Then Result = Result & " ReadOnly"
If attr And vbHidden Then Result = Result & " Hidden"
If attr And vbSystem Then Result = Result & " System"
If attr And vbArchive Then Result = Result & " Archive"
' Discard the first (extra) space.
GetAttribute = Mid$(Result, 2)

End Function


'This is how to set the attribute of a file with SetAttr.
'You can't use the SetAttr function on open files.

' Mark a file as Archive and Read-only.
'     SetAttr cd.Filename, vbArchive + vbReadOnly
' Change a file from hidden to visible, and vice versa.
'     SetAttr cd.Filename, GetAttr(cd.Filename) Xor vbHidden

'''''''''''''''''''''''''''''''''''''''''''''''''''''

'This function shows Path other than FileName.
Public Function StripPath(ByVal FullPath As String) As String
If InStr(FullPath, "\") = 0 Then
StripPath = FullPath
Exit Function
End If
StripPath = Left(FullPath, InStrRev(FullPath, "\"))
End Function

'Note:This function might not work with non-Latin alphabets.
Public Function CountSpaces(Text As String) As Long
Dim b() As Byte, i As Long
b() = Text
For i = 0 To UBound(b) Step 2
'Consider only even-numbered items.
'Save time and code using the function name as a local
'variable.
If b(i) = 32 Then CountSpaces = CountSpaces + 1
Next
End Function

'Removes everything but valid chars in a string
Public Function OnlyChars(ByVal txt As String, ByVal Uchars As Boolean, ByVal Lchars As Boolean, ByVal Digits As Boolean, Optional Remplacement As String) As String
Dim h As Long
Dim tempon As String
Dim Num

For h = 1 To Len(txt)
    Num = Asc(Mid$(txt, h, 1))
    'suppression des uchars
    If Uchars Then
        If (Num >= 65 And Num <= 90) Or Num = 32 Then
            tempon = tempon & Mid$(txt, h, 1)
        Else
            tempon = tempon & Remplacement
        End If
     
    End If
    
    If Lchars Then
        If (Num >= 97 And Num <= 122) Or Num = 32 Or Num = 224 Or Num = 225 Or Num = 232 Or Num = 233 Or Num = 234 Or Num = 235 Or Num = 249 Or Num = 250 Or Num = 251 Then
        tempon = tempon & Mid$(txt, h, 1)
        Else
        tempon = tempon & Remplacement
            
        End If
    
    End If
    
    If Digits Then
        If (Num >= 48 And Num <= 57) Or Num = 32 Then
            tempon = tempon & Mid$(txt, h, 1)
        
        Else
        tempon = tempon & Remplacement
            
        End If
    
    End If
   'If Uchars And Lchars And Digits Then tempon = tempon & Mid$(txt, h, 1)
Next
OnlyChars = tempon


End Function


'Drops shadow over controls
Sub DropShadow(Control As Object, Formu As Form)
    Dim n
 On Error Resume Next
    For n = 0 To 120 Step Measure
        DrawRect Formu, Control.Left + Desp + n / 2, Control.Top + Desp + n / 2, Control.Width - n, Control.Height - n, RGB(256 - (n + Band), 256 - (n + Band), 256 - (n + Band))
    Next
End Sub

Sub DrawRect(Control As Form, l, t, w, h, color)
    Dim x, xx
   On Error Resume Next
    For x = t To t + h Step Measure
        xx = x - t
        
        Select Case xx
            Case Is < FirstLines
                Control.Line (l + (FirstLines - xx), xx + t)-(l + w + xx - FirstLines, xx + t), color
            Case Is > h - EndLines
                Control.Line (l - h + EndLines + xx, xx + t)-(l + w - (EndLines + xx - h), xx + t), color
            Case Else
                Control.Line (l, xx + t)-(l + w, xx + t), color
        End Select

    Next
        
End Sub

Sub ShadowControls(Formu As Form)
    Dim n
    On Error Resume Next
    For n = 0 To Formu.Controls.count - 1
        DropShadow Formu.Controls(n), Formu
    Next
End Sub

