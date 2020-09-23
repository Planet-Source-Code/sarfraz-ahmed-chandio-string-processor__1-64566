VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmText 
   Caption         =   "Untitled - String Processor"
   ClientHeight    =   6660
   ClientLeft      =   1620
   ClientTop       =   1095
   ClientWidth     =   8880
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8970
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer WatchDog 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8040
      Top             =   4320
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6300
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9049
            Text            =   "String Processor v1.00"
            TextSave        =   "String Processor v1.00"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   882
            TextSave        =   "7/16/04"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   11955
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   4320
      Top             =   3120
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5760
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   12582912
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontStrikeThru  =   -1  'True
      FontUnderLine   =   -1  'True
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5340
      Top             =   3915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   30
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":1CFA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":1E0C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":1F1E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2030
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2142
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2254
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2366
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2478
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":258A
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":269C
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":27AE
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":28C0
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":29D2
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2AE4
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2BF6
            Key             =   "Drawing"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2D08
            Key             =   "Spell Check"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2E1A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":2F2C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":303E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":3150
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":3262
            Key             =   "Mystery Network Neighborhood"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":3A7C
            Key             =   "pdf"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":4016
            Key             =   "Note1"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":4330
            Key             =   "Encrypted File $#!"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":450A
            Key             =   "Text Document TXT"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":46E4
            Key             =   "MSGBOX041"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":49FE
            Key             =   "KEYS03"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":4D18
            Key             =   "Help File HLP1"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":4EF2
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmText.frx":520C
            Key             =   "MSGBOX02"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Selected Text"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Search"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Symbols"
            Object.ToolTipText     =   "Insert Symbols"
            ImageKey        =   "KEYS03"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Spell Check"
            Object.ToolTipText     =   "Spell Check"
            ImageKey        =   "Spell Check"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Encrypt/Decrypt"
            Object.ToolTipText     =   "Encrypt/Decrypt File"
            ImageKey        =   "Encrypted File $#!"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Convert To PDF"
            Object.ToolTipText     =   "Convert To PDF "
            ImageKey        =   "pdf"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About StringPro"
            ImageKey        =   "MSGBOX041"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "MSGBOX02"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFileS 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Index           =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuNewInstance 
         Caption         =   "New &Instance"
         Shortcut        =   ^I
      End
      Begin VB.Menu df 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveText 
         Caption         =   "&Save"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Index           =   0
      End
      Begin VB.Menu mnuSaveSelectionAs 
         Caption         =   "Save S&election As..."
      End
      Begin VB.Menu fgfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRevertSaved 
         Caption         =   "&Revert To Original"
         Shortcut        =   {F4}
      End
      Begin VB.Menu jh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu fghgfhgf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileManagement 
         Caption         =   "File &Management..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHis 
         Caption         =   "His 1"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHis 
         Caption         =   "His 2"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHis 
         Caption         =   "His 3"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHis 
         Caption         =   "His 4"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHis 
         Caption         =   "His 5"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSaveExit 
         Caption         =   "Save and Ex&it"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuClose 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHtml 
      Caption         =   "Ht&ml"
      Visible         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "Previe&w in Browser"
         Shortcut        =   {F5}
      End
      Begin VB.Menu fdgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParagraph 
         Caption         =   "<P> New &Paragraph"
      End
      Begin VB.Menu mnuLineBreak 
         Caption         =   "<BR> Line &Break"
      End
      Begin VB.Menu mnuFontTag 
         Caption         =   "<FONT> &Fonts"
      End
      Begin VB.Menu mnuCenterItem 
         Caption         =   "<CENTER> Centrali&ze"
      End
      Begin VB.Menu mnuHeadings 
         Caption         =   "<H1-6> &Headings"
      End
      Begin VB.Menu mnuFlatLine 
         Caption         =   "<HR> Fl&at Line"
      End
      Begin VB.Menu mnuBold 
         Caption         =   "<B> B&old"
      End
      Begin VB.Menu mnuItalic 
         Caption         =   "<I> &Italic"
      End
      Begin VB.Menu mnuUnderline 
         Caption         =   "<U> &Underline"
      End
      Begin VB.Menu mnuBulletList 
         Caption         =   "<BL> Bul&leted List"
      End
      Begin VB.Menu mnuNumberedList 
         Caption         =   "<OL> &Numbered List"
      End
      Begin VB.Menu mnuPredefined 
         Caption         =   "<PRE> Pre&defined "
      End
      Begin VB.Menu mnuTable 
         Caption         =   "<TABLE> &Table"
      End
      Begin VB.Menu mnuFrames 
         Caption         =   "<FRAMESET> Fram&e"
      End
      Begin VB.Menu fgdfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtherTags 
         Caption         =   "Other Ta&gs"
         Begin VB.Menu mnuAnimation 
            Caption         =   "<MARQUEE> &Animation"
         End
         Begin VB.Menu mnuEmphasis 
            Caption         =   "<EM> &Emphasis"
         End
         Begin VB.Menu mnuTypeWriterText 
            Caption         =   "<TT> Type&Writer Text"
         End
         Begin VB.Menu mnuMonoSpaceType 
            Caption         =   "<CODE> Monospace &Type"
         End
         Begin VB.Menu mnuDefination 
            Caption         =   "<DFN> &Defination"
         End
         Begin VB.Menu mnuCite 
            Caption         =   "<CITE> &Citation"
         End
         Begin VB.Menu mnuCitation 
            Caption         =   "<BLOCKQUOTE> &Long Citation"
         End
         Begin VB.Menu mnuSignature 
            Caption         =   "<ADDRESS> Si&gnature"
         End
         Begin VB.Menu mnuMonoSpace 
            Caption         =   "<KBD> &Monospace"
         End
         Begin VB.Menu mnuDIV 
            Caption         =   "<DIV> Break Doc&ument"
         End
         Begin VB.Menu mnuSTYLING 
            Caption         =   "<STYLE> Redefi&ne Tags"
         End
         Begin VB.Menu mnuFloatingFrame 
            Caption         =   "<IFRAME> Fl&oating Frame"
         End
      End
      Begin VB.Menu hfghfgh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMiscAttributes 
         Caption         =   "&Misc Attributes"
         Begin VB.Menu mnuBgColor 
            Caption         =   "BGCOLOR"
         End
         Begin VB.Menu mnuBackGround 
            Caption         =   "BACKGROUND"
         End
         Begin VB.Menu mnuTextLink 
            Caption         =   "TEXT"
         End
         Begin VB.Menu mnuLink 
            Caption         =   "LINK"
         End
         Begin VB.Menu mnuAlink 
            Caption         =   "ALINK"
         End
         Begin VB.Menu mnuVLINK 
            Caption         =   "VLINK"
         End
         Begin VB.Menu mnuAlt 
            Caption         =   "ALT"
         End
         Begin VB.Menu mnuSize 
            Caption         =   "SIZE"
         End
         Begin VB.Menu mnuFace 
            Caption         =   "FACE"
         End
         Begin VB.Menu mnuCOLORText 
            Caption         =   "COLOR"
         End
         Begin VB.Menu mnuName 
            Caption         =   "NAME"
         End
         Begin VB.Menu mnuSTYLE 
            Caption         =   "STYLE"
         End
         Begin VB.Menu mnuAlign 
            Caption         =   "ALIGN"
         End
         Begin VB.Menu mnuWidth 
            Caption         =   "WIDTH"
         End
         Begin VB.Menu mnuHeight 
            Caption         =   "HEIGHT"
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "BORDER"
         End
         Begin VB.Menu mnuVSpace 
            Caption         =   "VSPACE"
         End
         Begin VB.Menu mnuHSpace 
            Caption         =   "HSPACE"
         End
         Begin VB.Menu mnuUseMap 
            Caption         =   "USEMAP"
         End
         Begin VB.Menu mnuScrolling 
            Caption         =   "SCROLLING"
         End
         Begin VB.Menu mnuNoShade 
            Caption         =   "NOSHADE"
         End
         Begin VB.Menu mnuChecked 
            Caption         =   "CHECKED"
         End
         Begin VB.Menu mnuAction 
            Caption         =   "ACTION"
         End
         Begin VB.Menu mnuMethod 
            Caption         =   "METHOD"
         End
         Begin VB.Menu mnuBorderColor 
            Caption         =   "BORDERCOLOR"
         End
         Begin VB.Menu mnuBorderColorHeight 
            Caption         =   "BORDERCOLORHEIGHT"
         End
         Begin VB.Menu mnuBorderColorDark 
            Caption         =   "BORDERCOLORDARK"
         End
         Begin VB.Menu mnuLeft 
            Caption         =   "LEFT"
         End
         Begin VB.Menu mnuCenter 
            Caption         =   "CENTER"
         End
         Begin VB.Menu mnuRight 
            Caption         =   "RIGHT"
         End
         Begin VB.Menu mnuVAlign 
            Caption         =   "VALIGN"
         End
         Begin VB.Menu mnuCellSpacing 
            Caption         =   "CELLSPACING"
         End
         Begin VB.Menu mnuCellPadding 
            Caption         =   "CELLPADDING"
         End
         Begin VB.Menu mnuRowSpan 
            Caption         =   "ROWSPAN"
         End
         Begin VB.Menu mnuColSpan 
            Caption         =   "COLSPAN"
         End
         Begin VB.Menu mnuRows 
            Caption         =   "ROWS"
         End
         Begin VB.Menu mnuCols 
            Caption         =   "COLS"
         End
         Begin VB.Menu mnuMarginHeight 
            Caption         =   "MARGINHEIGHT"
         End
         Begin VB.Menu mnuMarginWidth 
            Caption         =   "MARGINWIDTH"
         End
         Begin VB.Menu mnuFameSpacing 
            Caption         =   "FRAMESPACING"
         End
         Begin VB.Menu mnuFrameBorder 
            Caption         =   "FRAMEBORDER"
         End
         Begin VB.Menu mnuCodeBase 
            Caption         =   "CODEBASE"
         End
         Begin VB.Menu mnuTypeDisc 
            Caption         =   "TYPE=DISC"
         End
         Begin VB.Menu mnuCLASS 
            Caption         =   "CLASS"
         End
         Begin VB.Menu mnuDynSrc 
            Caption         =   "DYNSRC"
         End
         Begin VB.Menu mnuLoopInfinite 
            Caption         =   "LOOP=INFINITE"
         End
         Begin VB.Menu mnuStart 
            Caption         =   "START"
         End
         Begin VB.Menu mnuTFoot 
            Caption         =   "TFOOT"
         End
      End
      Begin VB.Menu ghfdgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHyperlinImags 
         Caption         =   "HyperLin&ks"
         Begin VB.Menu mnuWebLink 
            Caption         =   "&Website"
         End
         Begin VB.Menu mnuEmailLink 
            Caption         =   "&Email"
         End
         Begin VB.Menu mnuImages 
            Caption         =   "&Picture"
         End
      End
      Begin VB.Menu hf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControls 
         Caption         =   "&Controls"
         Begin VB.Menu mnuScript 
            Caption         =   "&Script"
            Begin VB.Menu mnuJava 
               Caption         =   "&Jave Script"
            End
            Begin VB.Menu mnuVb 
               Caption         =   "&VB Script "
            End
         End
         Begin VB.Menu gfgf 
            Caption         =   "-"
         End
         Begin VB.Menu mnuForm 
            Caption         =   "&Form"
         End
         Begin VB.Menu mnuButton 
            Caption         =   "&Button Type"
            Begin VB.Menu mnuGeneralButton 
               Caption         =   "&General Button"
            End
            Begin VB.Menu mnuButtons 
               Caption         =   "&Submit Button"
            End
            Begin VB.Menu mnuResetButton 
               Caption         =   "&Reset Button"
            End
         End
         Begin VB.Menu mnuTextBoxType 
            Caption         =   "&TextBox Type"
            Begin VB.Menu mnuText 
               Caption         =   "&TextBox"
            End
            Begin VB.Menu mnuTextArea 
               Caption         =   "Text &Area"
            End
            Begin VB.Menu mnuPasswordBox 
               Caption         =   "&Password Box"
            End
         End
         Begin VB.Menu mnuRadio 
            Caption         =   "&Radio Button"
         End
         Begin VB.Menu mnuCheckb 
            Caption         =   "&Check Box"
         End
         Begin VB.Menu mnuListBox 
            Caption         =   "&List Box"
         End
         Begin VB.Menu mnuDownbutton 
            Caption         =   "Co&mbo Box"
         End
      End
      Begin VB.Menu ghgfh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Co&lors"
         Begin VB.Menu mnuRed 
            Caption         =   "&Red"
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu mnuGreen 
            Caption         =   "&Green"
         End
         Begin VB.Menu mnuBlack 
            Caption         =   "B&lack"
         End
         Begin VB.Menu mnuYellow 
            Caption         =   "&Yellow"
         End
         Begin VB.Menu mnuWhite 
            Caption         =   "&White"
         End
         Begin VB.Menu mnuSilver 
            Caption         =   "&Silver"
         End
         Begin VB.Menu mnuTeal 
            Caption         =   "&Teal"
         End
         Begin VB.Menu mnuPurple 
            Caption         =   "&Purple"
         End
         Begin VB.Menu mnuOlive 
            Caption         =   "&Olive"
         End
         Begin VB.Menu mnuNavy 
            Caption         =   "&Navy"
         End
         Begin VB.Menu mnuMaroon 
            Caption         =   "&Maroon"
         End
         Begin VB.Menu mnuLime 
            Caption         =   "L&ime"
         End
         Begin VB.Menu mnuGray 
            Caption         =   "Gr&ay"
         End
         Begin VB.Menu mnuFuchsia 
            Caption         =   "&Fuchsia"
         End
         Begin VB.Menu mnuAqua 
            Caption         =   "A&qua"
         End
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu dfgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSelectLine 
         Caption         =   "Copy  &Line"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu gf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clea&r All"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu ii 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTypingSound 
         Caption         =   "Typing &Sound"
         Shortcut        =   {F9}
      End
      Begin VB.Menu hfgh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelection 
         Caption         =   "Selecti&on to Notepad"
      End
      Begin VB.Menu fgf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHtmlEditing 
         Caption         =   "&HTML Editing"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindAgain 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFindinFiles 
         Caption         =   "Find in File&s..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuOccurrences 
         Caption         =   "Find &Occurrences"
      End
      Begin VB.Menu jhjhjhjh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuickRepalce 
         Caption         =   "&Quick Replace"
         Shortcut        =   {F6}
      End
      Begin VB.Menu fgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto Line..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu hfghfghfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontIncrease 
         Caption         =   "&Increase Font"
      End
      Begin VB.Menu mnuFontDecrease 
         Caption         =   "&Decrease Font"
      End
      Begin VB.Menu ioi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullScreen 
         Caption         =   "Full &Screen"
         Shortcut        =   {F11}
      End
      Begin VB.Menu jghjghjg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&File Properties..."
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuiInsertTD 
         Caption         =   "&Date and Time"
         Begin VB.Menu mnuSampleDate 
            Caption         =   "Time/Date Formats"
            Index           =   0
         End
      End
      Begin VB.Menu gh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertTextFile 
         Caption         =   "&Text File..."
      End
      Begin VB.Menu kj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSymbols 
         Caption         =   "&Symbols..."
         Shortcut        =   ^B
      End
      Begin VB.Menu gdfgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepetitive 
         Caption         =   "&Repetitive Characters"
      End
   End
   Begin VB.Menu mnuFormate 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFont 
         Caption         =   "&Fonts..."
         Shortcut        =   ^T
      End
      Begin VB.Menu nbnbnb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackColor 
         Caption         =   "&Page Color..."
      End
      Begin VB.Menu mnuForeColor 
         Caption         =   "Te&xt Color..."
      End
      Begin VB.Menu hg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeCase 
         Caption         =   "Select Ca&se"
         Begin VB.Menu mnuUpperCase 
            Caption         =   "&UPPER CASE"
         End
         Begin VB.Menu mnuLowerCase 
            Caption         =   "&lower case"
         End
         Begin VB.Menu mnuProperCapitalize 
            Caption         =   "&Proper Case"
         End
         Begin VB.Menu mnuLowerCaps 
            Caption         =   "Lower &caps"
         End
      End
      Begin VB.Menu gdfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAllowChars 
         Caption         =   "&Allow Only Characters"
      End
      Begin VB.Menu mnuRemoveExtraSpaces 
         Caption         =   "&Remove Extra Spaces"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuPdf 
         Caption         =   "Text to P&DF..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTexttoExe 
         Caption         =   "Text to &EXE"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuTexttoHTML 
         Caption         =   "Text to &HTML..."
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEDFile 
         Caption         =   "Encrypt/Decrypt &Text..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSpellChecker 
         Caption         =   "&Spell Checker..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuCompareFiles 
         Caption         =   "Co&mpare Files..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuWordCount 
         Caption         =   "&Word Count..."
      End
      Begin VB.Menu mnuNumFigurs 
         Caption         =   "&Numbers to Words"
      End
      Begin VB.Menu mnuNumberRoman 
         Caption         =   "Numbers to &Romans"
      End
      Begin VB.Menu hc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWinTools 
      Caption         =   "&WinTools"
      Begin VB.Menu mnuCalculator 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnuAddBook 
         Caption         =   "&Address Book"
      End
      Begin VB.Menu mnuSFC 
         Caption         =   "System &File Checker"
      End
      Begin VB.Menu mnuScanDisk 
         Caption         =   "Scan &Disk"
      End
      Begin VB.Menu mnuScanRegistry 
         Caption         =   "&Scan Registry"
      End
      Begin VB.Menu mnuOptimizeDrives 
         Caption         =   "&Optimize Drives"
      End
      Begin VB.Menu mnuCompatible 
         Caption         =   "&Make Compatible"
      End
      Begin VB.Menu mnuPrograman 
         Caption         =   "&Program Manager"
      End
      Begin VB.Menu mnuFileManager 
         Caption         =   "Fi&le Manager"
      End
      Begin VB.Menu mnuPackager 
         Caption         =   "Packa&ger"
      End
      Begin VB.Menu mnupopup 
         Caption         =   "Dr.&Watson"
      End
      Begin VB.Menu mnuStartupDisk 
         Caption         =   "Create Startup Dis&k"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuReadMe 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHomePage 
         Caption         =   "Home&Page"
      End
      Begin VB.Menu mnuContact 
         Caption         =   "&Contact"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'                 **********************
'                 STRING PROCESSOR V1.00
'                 **********************
'
'                     Developed by:
'                "SARFRAZ AHMED CHANDIO"


'                      CONTACT ME
'Email address:- sarfrazahmed_pk@yahoo.com
'Website:- http://www.angelfire.com/ultra/sarfraz

'Warning:-You are prohibted as per copyright laws to
're-use this piece of code as long as you don't
'distribute this code with YOUR name.

'Dated:- 1/1/04

Option Explicit
Dim Temp As Integer
Dim History(4) As String
Dim UseSound As String
Public CancelClicked As Boolean
Private Const WM_SETTEXT = &HC

Private Sub Form_Activate()
WatchDog.Enabled = True
End Sub

Private Sub Form_Load()

Jimmy = GetSetting(App.EXEName, "LastFile", "File", "")
If Jimmy = "" Or Not FileExists(GetSetting(App.EXEName, "LastFile", "File", "")) Then
Jimmy = "Untitled"
Else
Jimmy = GetSetting(App.EXEName, "LastFile", "File", "")
End If

On Error GoTo RunNext
Open Jimmy For Input As #1
Text1.Text = Input$(LOF(1), #1)
On Error Resume Next
Text1.Tag = Text1.Text
frmText.Caption = Jimmy & " - String Processor"
StatusBar1.Panels(1).Text = Jimmy
Sarfraz = False
mnuUndo.Enabled = False
Close #1

cd.filename = Jimmy

On Error Resume Next
AddToHis


RunNext:
Sarfraz = False
'Jimmy = "Untitled"
frmText.Caption = Jimmy & " - String Processor"
mnuUndo.Enabled = False


Dim Values(10)

Values(0) = GetSetting(App.EXEName, "StartupPos", "Position", "")

If Values(0) = 1 Then Me.WindowState = 1
If Values(0) = 2 Then Me.WindowState = 2
If Values(0) = 0 Then Me.WindowState = 0

Values(1) = GetSetting(App.EXEName, "FontNameSize", "FontName", "")
If Values(1) = "" Then
Text1.FontName = "Verdana"
Else
Text1.FontName = GetSetting(App.EXEName, "FontNameSize", "FontName", "")
End If

Values(2) = GetSetting(App.EXEName, "FontNameSize", "FontSize", "")
If Values(2) = "" Then
Text1.FontSize = 11
Else
Text1.FontSize = GetSetting(App.EXEName, "FontNameSize", "FontSize", "")
End If

Values(3) = GetSetting(App.EXEName, "FontStyle", "FontBold", "")
If Values(3) = "" Then
Text1.FontBold = False
Else
Text1.FontBold = GetSetting(App.EXEName, "FontStyle", "FontBold", "")
End If

Values(4) = GetSetting(App.EXEName, "FontStyle", "FontItalic", "")
If Values(4) = "" Then
Text1.FontItalic = False
Else
Text1.FontItalic = GetSetting(App.EXEName, "FontStyle", "FontItalic", "")
End If

Values(5) = GetSetting(App.EXEName, "FontStyle", "FontUnderline", "")
If Values(5) = "" Then
Text1.FontUnderline = False
Else
Text1.FontUnderline = GetSetting(App.EXEName, "FontStyle", "FontUnderline", "")
End If

Values(6) = GetSetting(App.EXEName, "FontStyle", "FontStrikethru", "")
If Values(6) = "" Then
Text1.FontStrikethru = False
Else
Text1.FontStrikethru = GetSetting(App.EXEName, "FontStyle", "FontStrikethru", "")
End If

Values(7) = GetSetting(App.EXEName, "Toolbar", "Visible", "")
If Values(7) = "" Or IsError(Values(7)) = True Or IsError(GetSetting(App.EXEName, "Toolbar", "Visible", "")) = True Then
Me.Toolbar1.Visible = True
Else
Me.Toolbar1.Visible = GetSetting(App.EXEName, "Toolbar", "Visible", "")
End If

Dim a

a = GetSetting(App.EXEName, "TypingSound", "Play", "")
If a = 1 Then Call mnuTypingSound_Click

Dim z
z = GetSetting(App.EXEName, "save", "his1", "")
If z = "" Then
Sep.Visible = False
Else
Sep.Visible = True
End If

If frmOptions.chkToolbar.Value = 1 Then
mnuToolbar.Checked = True
Else
mnuToolbar.Checked = False
End If


Values(8) = GetSetting(App.EXEName, "Color", "TextForeColor", "")
If Values(8) = "" Then
Text1.ForeColor = vbBlack
Else
Text1.ForeColor = GetSetting(App.EXEName, "Color", "TextForeColor", "")
End If

Values(9) = GetSetting(App.EXEName, "Color", "TextBackColor", "")
If Values(9) = "" Then
Text1.BackColor = vbWhite
Else
Text1.BackColor = GetSetting(App.EXEName, "Color", "TextBackColor", "")
End If

Values(10) = GetSetting(App.EXEName, "HtmlMenu", "Visible", "")
If Values(10) = 0 Or Values(10) = "" Then
Me.mnuHtml.Visible = False
Else
Me.mnuHtml.Visible = True
End If


'Show UserName
Dim sBuffer As String
Dim lSize As Long
sBuffer = Space$(255)
lSize = Len(sBuffer)
On Error Resume Next
Call GetUserName(sBuffer, lSize)
If lSize > 0 Then
StatusBar1.Panels(1).Text = "Welcome! " & Left$(sBuffer, lSize)
Else
StatusBar1.Panels(1).Text = "String Processor v1.00"
End If

On Error Resume Next
Call SetHook(Text1.hWnd, True)

On Error Resume Next
cd.filename = GetSetting(App.EXEName, "file", "file_pattern", "*.txt")
History(0) = GetSetting(App.EXEName, "save", "his1", "")
History(1) = GetSetting(App.EXEName, "save", "his2", "")
History(2) = GetSetting(App.EXEName, "save", "his3", "")
History(3) = GetSetting(App.EXEName, "save", "his4", "")
History(4) = GetSetting(App.EXEName, "save", "his5", "")

For i = 0 To 4
If History(i) <> "" Then
mnuHis(i).Caption = History(i)
On Error Resume Next
mnuHis(i).Visible = True
Sep.Visible = True
Else
mnuHis(i).Visible = False
End If
Next i

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Sarfraz = True Then
Dim Chandio As Integer
Chandio = MsgBox("Do you want to save the changes to: " & vbCr & Jimmy, vbYesNoCancel + vbQuestion, "Save File?")

If Chandio = vbCancel Then
Cancel = True
ElseIf Chandio = vbNo Then

On Error Resume Next
SaveSetting App.EXEName, "save", "his1", History(0)
SaveSetting App.EXEName, "save", "his2", History(1)
SaveSetting App.EXEName, "save", "his3", History(2)
SaveSetting App.EXEName, "save", "his4", History(3)
SaveSetting App.EXEName, "save", "his5", History(4)
SaveSetting App.EXEName, "file", "file_pattern", cd.filename
SaveSetting App.EXEName, "LastFile", "File", Jimmy

frmText.cd.filename = Jimmy
'Jimmy = frmText.cd.filename

UnloadForms

Exit Sub
Else
SaveText

' If cancel is clicked at the FileSaveAs box then don't
' unload.
If CancelClicked = True Then
 Cancel = True
 CancelClicked = False
Else
 Cancel = False
End If

End If
End If

End Sub

Private Sub Form_Resize()

If Toolbar1.Visible = True Then
 Text1.Width = frmText.Width - 140
 On Error Resume Next
 Text1.Height = frmText.Height - 1480
 Text1.Top = 445
Else
 On Error Resume Next
 Text1.Height = frmText.Height - 1000
 Text1.Width = frmText.Width - 140
 Text1.Top = 1
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveSetting App.EXEName, "save", "his1", History(0)
SaveSetting App.EXEName, "save", "his2", History(1)
SaveSetting App.EXEName, "save", "his3", History(2)
SaveSetting App.EXEName, "save", "his4", History(3)
SaveSetting App.EXEName, "save", "his5", History(4)
SaveSetting App.EXEName, "file", "file_pattern", cd.filename
SaveSetting App.EXEName, "LastFile", "File", Jimmy

frmText.cd.filename = Jimmy

UnloadForms

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuAction_Click()
On Error GoTo Hell
Text1.SelText = "ACTION="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuAddBook_Click()
On Error Resume Next
Shell "C:\Program Files\Outlook Express\wab.exe", vbNormalFocus
End Sub

Private Sub mnuAlign_Click()
On Error GoTo Hell
Text1.SelText = "ALIGN="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuAlink_Click()
On Error GoTo Hell
Text1.SelText = "ALINK="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuAllowChars_Click()
If TextNotSelected Then Exit Sub

Dim Str$, txt$
Str = Text1.SelText

Screen.MousePointer = 11
txt = OnlyChars(Str, True, True, True)

On Error GoTo Hell
Text1.SelText = txt
mnuUndo.Enabled = False
Screen.MousePointer = 0

Exit Sub
Hell:
HellError
Screen.MousePointer = 0
End Sub

Private Sub mnuAlt_Click()
On Error GoTo Hell
Text1.SelText = "ALT="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuAnimation_Click()
On Error GoTo Hell
Text1.SelText = "<MARQUEE>      </MARQUEE>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuAqua_Click()
On Error GoTo Hell
Text1.SelText = "AQUA"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBackColor_Click()
On Error GoTo Handler
frmText.cd.CancelError = True
frmText.cd.ShowColor
frmText.Text1.BackColor = frmText.cd.color

Exit Sub
Handler:
End Sub

Private Sub mnuBackGround_Click()
On Error GoTo Hell
Text1.SelText = "BACKGROUND="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBgColor_Click()
On Error GoTo Hell
Text1.SelText = "BGCOLOR="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBlack_Click()
On Error GoTo Hell
Text1.SelText = "BLACK"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBlue_Click()
On Error GoTo Hell
Text1.SelText = "BlUE"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBold_Click()
On Error GoTo Hell
Text1.SelText = "<B>" & "   " & "</B>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBorder_Click()
On Error GoTo Hell
Text1.SelText = "BORDER="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBorderColor_Click()
On Error GoTo Hell
Text1.SelText = "BORDERCOLOR="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBorderColorDark_Click()
On Error GoTo Hell
Text1.SelText = "BORDERCOLORDARK"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBorderColorHeight_Click()
On Error GoTo Hell
Text1.SelText = "BORDERCOLORHEIGHT="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuBulletList_Click()
On Error GoTo Hell
Text1.SelText = "<UL>" & vbCrLf & "<LI> ItemOne" & vbCrLf & "<LI>ItemTwo" & vbCrLf & "<LI>ItemThree" & vbCrLf & "</UL>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuButtons_Click()
On Error GoTo Hell
Text1.SelText = "<INPUT TYPE=SUBMIT   NAME=""" + "SubmitButton" + """   VALUE=""" + "SubmitButton" + """><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCalculator_Click()
On Error Resume Next
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mnuCellPadding_Click()
On Error GoTo Hell
Text1.SelText = "CELLPADDING="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCellSpacing_Click()
On Error GoTo Hell
Text1.SelText = "CELLSPACING="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCenter_Click()
On Error GoTo Hell
Text1.SelText = "CENTER"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCenterItem_Click()
On Error GoTo Hell
Text1.SelText = "<CENTER>" & vbCrLf & vbCrLf & "</CENTER>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCheckb_Click()
On Error GoTo Hell
Text1.SelText = "<INPUT TYPE=CHECKBOX   NAME=""" + "CheckBox" + """   VALUE=""" + "CheckBox" + """   CHECKED><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuChecked_Click()
On Error GoTo Hell
Text1.SelText = "CHECKED"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCitation_Click()
On Error GoTo Hell
Text1.SelText = "<BLOCKQUOTE>     </BLOCKQUOTE>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCite_Click()
On Error GoTo Hell
Text1.SelText = "<CITE>     </CITE>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCLASS_Click()
On Error GoTo Hell
Text1.SelText = "CLASS="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuClear_Click()
On Error GoTo Hell
Text1.SelText = ""
mnuUndo.Enabled = False

Exit Sub
Hell:
mnuUndo.Enabled = False
HellError
End Sub

Private Sub mnuClearAll_Click()
On Error GoTo Hell
Text1.Text = ""
Sarfraz = True
mnuUndo.Enabled = False

Exit Sub
Hell:
mnuUndo.Enabled = False
HellError
End Sub

Private Sub mnuClose_Click()

If Sarfraz = True Then
Dim Chandio As Integer
Chandio = MsgBox("Do you want to save the changes to: " & vbCr & Jimmy, vbYesNoCancel + vbQuestion, "Save File?")

If Chandio = vbYes Then
SaveText
ElseIf Chandio = vbNo Then
Sarfraz = False
Else
Exit Sub
End If
End If

If Chandio = 6 Then If Sarfraz = True Then Exit Sub

Unload Me
End Sub

Private Sub mnuCodeBase_Click()
On Error GoTo Hell
Text1.SelText = "CODEBASE="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCOLORText_Click()
On Error GoTo Hell
Text1.SelText = "COLOR="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCols_Click()
On Error GoTo Hell
Text1.SelText = "COLS="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuColSpan_Click()
On Error GoTo Hell
Text1.SelText = "COLSPAN="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCompareFiles_Click()
With frmText.cd
On Error GoTo Handler
.CancelError = True
.DefaultExt = "txt"
.filename = ""
.DialogTitle = "Select File Tobe Compated With Opened File"
.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
.ShowOpen


Open cd.filename For Input As #1
Text3.Text = Input(LOF(1), #1)

cd.filename = Jimmy

If Text3.Text = Text1.Text Then
frmCompare.Text1.Text = String(41, "*") & vbNewLine & "THERE IS NO DIFFERENCE BETWEEN THE TWO FILES" & vbNewLine & String(41, "*")
frmCompare.Show 1
Else

On Error Resume Next
frmCompare.Text1.SelStart = 1
frmCompare.Text1.SelText = String(46, "*") & vbNewLine & "THE TWO FILES ARE DIFFERENT DUE TO FOLLOWING TEXT" & vbNewLine & String(46, "*")

frmCompare.Text1.SelText = vbNewLine & vbNewLine & TrackChanges(Text1.Text, Text3.Text)

frmCompare.Text1.SelStart = Len(frmCompare.Text1.Text)
frmCompare.Text1.SelText = vbNewLine & vbNewLine & vbNewLine & String(43, "*") & vbNewLine & "THE REST OF THE TEXT (IF ANY) IS SAME IN BOTH FILES" & vbNewLine & String(43, "*") & vbNewLine & "NOTE:The different text encompassing the similar text is shown as different." & vbNewLine & "Note also that SPACES,CR,LF also cause the text tobe shown as" & vbNewLine & "different."
frmCompare.Show 1

Close #1
End If
End With
cd.filename = Jimmy
Close #1

Exit Sub
Handler:
cd.filename = Jimmy
Close #1
End Sub

Private Sub mnuCompatible_Click()
On Error Resume Next
Shell "MKCOMPAT.EXE", vbNormalFocus
End Sub

Private Sub mnuContact_Click()
On Error Resume Next
Shell "start mailto:sarfrazahmed_pk@yahoo.com", vbHide
End Sub

Private Sub mnuCopy_Click()
On Error GoTo Hell
Clipboard.SetText Text1.SelText

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuCut_Click()
On Error GoTo Hell
Clipboard.SetText Text1.SelText
Text1.SelText = ""
mnuUndo.Enabled = False

Exit Sub
Hell:
mnuUndo.Enabled = False
HellError
End Sub

Private Sub mnuDefination_Click()
On Error GoTo Hell
Text1.SelText = "<DFN>     </DFN>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuDIV_Click()
On Error GoTo Hell
Text1.SelText = "<DIV>" & vbCrLf & vbCrLf & "</DIV>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuDownbutton_Click()
On Error GoTo Hell
Text1.SelText = "<SELECT  NAME=""PutNameHere""  SIZE=1  TABINDEX=3>" & vbCrLf & "   " & vbCrLf & "<OPTION   VALUE=PutValueHere>ItemOne</OPTION>" & vbCrLf & "<OPTION   VALUE=PutValueHere>ItemTwo</OPTION>" & vbCrLf & "<OPTION   VALUE=PutValueHere>ItemThree</OPTION>" & vbCrLf & vbCrLf & "</SELECT><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuDynSrc_Click()
On Error GoTo Hell
Text1.SelText = "DYNSRC="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuEDFile_Click()
frmEDFile.Show vbModal
End Sub

Private Sub mnuEmailLink_Click()
On Error GoTo Hell
Text1.SelText = "<A HREF=""MAILTO:EmailAddressHere"">Description</A><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuEmphasis_Click()
On Error GoTo Hell
Text1.SelText = "<EM>     </EM>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuFace_Click()
On Error GoTo Hell
Text1.SelText = "FACE="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuFameSpacing_Click()
On Error GoTo Hell
Text1.SelText = "FRAMESPACING="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuFileManagement_Click()
frmFileManagement.Show 1
End Sub

Private Sub mnuFileManager_Click()
On Error Resume Next
Shell "WINFILE.EXE", vbNormalFocus
End Sub

Private Sub mnuFileS_Click(index As Integer)
If Sarfraz = True Then
mnuSaveText(3).Enabled = True
mnuSaveExit.Enabled = True
Else
mnuSaveText(3).Enabled = False
mnuSaveExit.Enabled = False
End If

For i = 0 To 4
If History(i) <> "" Then
mnuHis(i).Caption = History(i)
On Error GoTo theError
mnuHis(i).Visible = True
Sep.Visible = True
Else
On Error Resume Next
mnuHis(i).Visible = False
End If
Next i

Exit Sub
theError:
MsgBox Err.Description, vbInformation, "Error"
End Sub

Private Sub mnuFileSaveAs_Click(index As Integer)
SaveTextAs
End Sub

Private Sub mnuFind_Click()
Load frmFind
frmFind.Show 0, Me
frmFind.txtFind.SetFocus

If Text1.Text = frmFind.txtFind.Text Or frmFind.cmdFindAgain.Enabled = True Then
frmFind.txtReplace.Enabled = True
frmFind.Label2.Enabled = True
Else
frmFind.txtReplace.Enabled = False
frmFind.Label2.Enabled = False
End If

If frmFind.txtFind = "" Then
mnuFindAgain.Enabled = False
End If

End Sub

Private Sub mnuFindAgain_Click()
frmFind.cmdFindAgain.Value = True
End Sub

Private Sub mnuFindinFiles_Click()
frmSearchinFiles.Show , Me
End Sub

Private Sub mnuFlatLine_Click()
On Error GoTo Hell
Text1.SelText = "<HR SIZE=1  COLOR=BLACK>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuFloatingFrame_Click()
On Error GoTo Hell
Text1.SelText = "<IFRAME     >" & vbCrLf & vbCrLf & "</IFRAME>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuFont_Click()
On Error GoTo Hell
cd.CancelError = True
cd.flags = cdlCFBoth + cdlCFEffects
cd.FontName = Text1.FontName
cd.FontSize = Text1.FontSize
cd.FontBold = Text1.FontBold
cd.FontItalic = Text1.FontItalic
cd.color = Text1.ForeColor
cd.FontStrikethru = Text1.FontStrikethru
cd.FontUnderline = Text1.FontUnderline

cd.ShowFont

Text1.FontName = cd.FontName
Text1.FontBold = cd.FontBold
Text1.FontItalic = cd.FontItalic
Text1.FontSize = cd.FontSize
Text1.ForeColor = cd.color
Text1.FontStrikethru = cd.FontStrikethru
Text1.FontUnderline = cd.FontUnderline

StatusBar1.Panels(1).Text = cd.FontName & " " & cd.FontSize & "pt" & " Selected"
mnuUndo.Enabled = False

Exit Sub
Hell:
End Sub

Private Sub mnuFontDecrease_Click()
On Error GoTo Hell
Text1.FontSize = Text1.FontSize - 1

Exit Sub
Hell:
If Err.Number = 380 Then
MsgBox "Can't reduce the Font anymore!", vbInformation
Else
MsgBox Err.Description, vbInformation
End If
End Sub

Private Sub mnuFontIncrease_Click()
On Error GoTo Hell
Text1.FontSize = Text1.FontSize + 1

Exit Sub
Hell:
If Err.Number = 380 Then
MsgBox "Can't reduce the Font anymore!", vbInformation
Else
MsgBox Err.Description, vbInformation
End If
End Sub

Private Sub mnuFontTag_Click()
On Error GoTo Hell
Text1.SelText = "<FONT FACE=""VERDANA""   SIZE=2   COLOR=BLACK>" & vbCrLf & vbCrLf & "</FONT>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuForeColor_Click()
On Error GoTo Handler
frmText.cd.CancelError = True
frmText.cd.ShowColor
frmText.Text1.ForeColor = frmText.cd.color

Exit Sub
Handler:
End Sub

Private Sub mnuForm_Click()
On Error GoTo Hell
Text1.SelText = "<FORM   NAME=""" + "FormName" + """   ACTION=""" + "Action" + """   METHOD=""" + "Method" + """><P>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & " </FORM>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuFrameBorder_Click()
On Error GoTo Hell
Text1.SelText = "FRAMEBORDER="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuFrames_Click()
On Error GoTo Hell
Text1.SelText = "<FRAMESET  COLS=10>" & vbCrLf & vbCrLf & "</FRAMESET>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuFuchsia_Click()
On Error GoTo Hell
Text1.SelText = "FUCHSIA"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuFullScreen_Click()

On Error GoTo FullScreenError
frmScreen.txtScreen = Text1
frmScreen.Show 1

Exit Sub
FullScreenError:
If Err.Number = 7 Or Err.Description = "Out of memory" Then
MsgBox "File is too large!!", vbExclamation
Else
MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub mnuGeneralButton_Click()
On Error GoTo Hell
Text1.SelText = "<INPUT TYPE=BUTTON   NAME=""" + "GeneralButton" + """   VALUE=""" + "GeneralButton" + """><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuGoto_Click()
frmGoTo.Show 0, Me
End Sub

Private Sub mnuGray_Click()
On Error GoTo Hell
Text1.SelText = "GRAY"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuGreen_Click()
On Error GoTo Hell
Text1.SelText = "GREEN"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuHeadings_Click()
On Error GoTo Hell
Text1.SelText = "<H6>" & "   " & "</H6>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuHeight_Click()
On Error GoTo Hell
Text1.SelText = "HEIGHT="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuHis_Click(index As Integer)
If Sarfraz = True Then
Temp = MsgBox("Do you want to save the changes to: " & vbCr & Jimmy, vbYesNoCancel + vbQuestion, "Save File?")

If Temp = vbYes Then
SaveText
ElseIf Temp = vbNo Then
Sarfraz = False
Else
Exit Sub
End If
End If

If Temp = 6 Then If Sarfraz = True Then Exit Sub

cd.filename = mnuHis(index).Caption

Close #1
On Error GoTo NoFile
Open mnuHis(index).Caption For Input As #1
Text1.Text = Input$(LOF(1), #1)
On Error GoTo Handler
Text1.Tag = Text1.Text
Jimmy = mnuHis(index).Caption
frmText.Caption = Jimmy & " - String Processor"
StatusBar1.Panels(1).Text = Jimmy
Sarfraz = False
mnuUndo.Enabled = False
Close #1

cd.filename = Jimmy

Exit Sub
NoFile:
If Not FileExists(cd.filename) Then
MsgBox "File was not found!", vbInformation, "File not found"
cd.filename = Jimmy
mnuHis(index).Caption = ""
History(index) = ""
mnuHis(index).Visible = False
ElseIf Err.Description = "Out of memory" Or Err.Number = 7 Then
MsgBox "File is too large!!", vbExclamation
Else
MsgBox Err.Description, vbInformation, "Error"
'mnuHis(Index).Caption = ""
'History(Index) = ""
'mnuHis(Index).Visible = False
End If

Exit Sub
Handler:
If Err.Number = 7 Or Err.Description = "Out of Memory" Then
Close #1
MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub mnuHomePage_Click()
Dim Go As Long

On Error GoTo Hell
Go = ShellExecute(Me.hWnd, "open", App.path & "\StringPro Help.html", "", App.path, 1)

Hell:
If Not FileExists(App.path & "\StringPro Help.html") Then
On Error Resume Next
Shell "start http://www.angelfire.com/ultra/sarfraz", vbHide
End If
End Sub

Private Sub mnuHSpace_Click()
On Error GoTo Hell
Text1.SelText = "HSPACE="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuHtmlEditing_Click()
mnuHtml.Visible = Not mnuHtml.Visible

If mnuHtml.Visible = True Then
On Error GoTo Hell
Text1.SelText = "<HTML>" & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & "My WebPage</TITLE>" & vbCrLf & "</HEAD>" & vbCrLf & "<BODY BGCOLOR=WHITE   TEXT=BLACK>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & "</HTML>"
mnuUndo.Enabled = False
Else
Exit Sub
End If

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuImages_Click()
On Error GoTo Hell
Text1.SelText = "<IMG ALIGN=LEFT WIDTH=200 HEIGHT=200  BORDER=0  SRC=""EnterPathHere"">"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuInsert_Click()
If mnuSampleDate.count = 1 Then
For i = 1 To 16
'Notice the use of the "Load" used for menu commands too.
Load mnuSampleDate(i)
Next
End If

mnuSampleDate(0).Caption = Format(Now, "General Date")
mnuSampleDate(1).Caption = Format(Now, "m/d/yyyy")
mnuSampleDate(2).Caption = Format(Now, "d/m/yyyy")
mnuSampleDate(3).Caption = Format(Now, "d mmmm, yyyy")
mnuSampleDate(4).Caption = Format(Now, "d mmmm")
mnuSampleDate(5).Caption = Format(Now, "mmm dd, yyyy")
mnuSampleDate(6).Caption = Format(Now, "d,dddd")
mnuSampleDate(7).Caption = Format(Now, "mm ddd, yyyy")
mnuSampleDate(8).Caption = Format(Now, "Long Date")
mnuSampleDate(9).Caption = "-"
mnuSampleDate(10).Caption = "12 hr Clock = " & Format(Now, "h:mm:ss am/pm")
mnuSampleDate(11).Caption = "12 hr Clock = " & Format(Now, "h:mm AM/PM")
mnuSampleDate(12).Caption = "24 hr Clock = " & Format(Now, "h:mm:ss")
mnuSampleDate(13).Caption = "24 hr Clock = " & Format(Now, "h:mm")
mnuSampleDate(14).Caption = "-"
mnuSampleDate(15).Caption = "Day of the year = " & Format(Now, "y")
mnuSampleDate(16).Caption = "Week of the year = " & Format(Now, "ww")

End Sub

Private Sub mnuInsertTextFile_Click()
Dim choice As Integer
Dim filenum As Integer

On Error GoTo InsertErrorTrap
frmText.cd.DialogTitle = "Insert File"
frmText.cd.filename = ""
frmText.cd.Filter = "Text files (*.txt)|*.txt|All Files (*.*)|*.*"
frmText.cd.ShowOpen

If frmText.cd.filename <> "" Then
filenum = FreeFile
Open frmText.cd.filename For Input As filenum
frmText.Text1.SelStart = Len(Text1.Text)
frmText.Text1.SelText = Input(LOF(filenum), filenum)
Close (filenum)
frmText.cd.filename = Jimmy
frmText.cd.filename = Jimmy
mnuUndo.Enabled = False
End If

Pakistan:
frmText.cd.filename = Jimmy
Jimmy = frmText.cd.filename
Close FreeFile

Exit Sub
InsertErrorTrap:
mnuUndo.Enabled = False
If Err.Number = 32755 Then
GoTo Pakistan
frmText.cd.filename = Jimmy
Jimmy = frmText.cd.filename
mnuUndo.Enabled = False
End If
End Sub

Private Sub mnuItalic_Click()
On Error GoTo Hell
Text1.SelText = "<I>" & "   " & "</I>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuJava_Click()
On Error GoTo Hell
Text1.SelText = "<SCRIPT LANGUAGE =""" + "JaveScript" + """>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</SCRIPT>"
mnuUndo.Enabled = False
Exit Sub
Hell:
HellError
End Sub

Private Sub mnuLeft_Click()
On Error GoTo Hell
Text1.SelText = "LEFT"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuLime_Click()
On Error GoTo Hell
Text1.SelText = "LIME"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuLineBreak_Click()
On Error GoTo Hell
Text1.SelText = "<BR>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuLink_Click()
On Error GoTo Hell
Text1.SelText = "LINK="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuListBox_Click()
On Error GoTo Hell
Text1.SelText = "<SELECT  NAME=""PutNameHere""  SIZE=3  MULTIPLE>" & vbCrLf & "   " & vbCrLf & "<OPTION   VALUE=PutValueHere>ItemOne</OPTION>" & vbCrLf & "<OPTION   VALUE=PutValueHere>ItemTwo</OPTION>" & vbCrLf & "<OPTION   VALUE=PutValueHere>ItemThree</OPTION>" & vbCrLf & vbCrLf & "</SELECT><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuLoopInfinite_Click()
On Error GoTo Hell
Text1.SelText = "LOOP=INFINITE"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuLowerCaps_Click()
Dim StartPoint As Integer, SelectedLength As Integer

If TextNotSelected Then Exit Sub
StartPoint = Text1.SelStart
On Error GoTo Hell
SelectedLength = Text1.SelLength
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbLowerCase)
Text1.SelStart = StartPoint
Text1.SelLength = 1
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbUpperCase)
Text1.SelStart = Text1.SelStart + SelectedLength
mnuUndo.Enabled = False

Exit Sub
Hell:
If Err.Number = 6 Then
MsgBox Err.Description, vbCritical, "Error"
Else
HellError
End If
End Sub

Private Sub mnuLowerCase_Click()
If TextNotSelected Then Exit Sub
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbLowerCase)
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuMarginHeight_Click()
On Error GoTo Hell
Text1.SelText = "MARGINHEIGHT="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuMarginWidth_Click()
On Error GoTo Hell
Text1.SelText = "MARGINWIDTH="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuMaroon_Click()
On Error GoTo Hell
Text1.SelText = "MAROON"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuMethod_Click()
On Error GoTo Hell
Text1.SelText = "METHOD="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuMonoSpace_Click()
On Error GoTo Hell
Text1.SelText = "<KBD>     </KBD>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuMonoSpaceType_Click()
On Error GoTo Hell
Text1.SelText = "<CODE>     </CODE>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuName_Click()
On Error GoTo Hell
Text1.SelText = "NAME="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuNavy_Click()
On Error GoTo Hell
Text1.SelText = "NAVY"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuNew_Click(index As Integer)

If Sarfraz = True Then
Dim Chandio As Integer
Chandio = MsgBox("Do you want to save the changes to: " & vbCr & Jimmy, vbYesNoCancel + vbQuestion, "Save File?")

If Chandio = vbYes Then
SaveText
ElseIf Chandio = vbNo Then
Sarfraz = False
Else
Exit Sub
End If
End If

If Chandio = 6 Then If Sarfraz = True Then Exit Sub

Jimmy = "Untitled"
Sarfraz = False
Text1.Text = ""
Text1.Tag = ""
mnuUndo.Enabled = False
StatusBar1.Panels(1).Text = "New Text File"
frmText.Caption = Jimmy & " - String Processor"
cd.filename = ""

End Sub

Private Sub mnuNewInstance_Click()
On Error Resume Next
Call Shell(VB.App.path + "\" + VB.App.EXEName)
End Sub

Private Sub mnuNoShade_Click()
On Error GoTo Hell
Text1.SelText = "NOSHADE"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuNumberedList_Click()
On Error GoTo Hell
Text1.SelText = "<OL>" & vbCrLf & "<LI>ItemOne" & vbCrLf & "<LI>ItemTwo" & vbCrLf & "<LI>ItemThree" & vbCrLf & "</OL>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuNumberRoman_Click()

If TextNotSelected Then Exit Sub
If Not IsNumeric(Text1.SelText) Then
MsgBox "Select numeric text only !", vbExclamation
Exit Sub
End If

On Error GoTo Hell
Temp = MsgBox("Proceed with conversion?" & vbCr & "This action can't be undone!", vbYesNoCancel + vbQuestion)

If Text1.SelText > 10000000 Then
MsgBox "Sorry! " & vbCr & "The number is too large!!", vbCritical
Exit Sub
Else
If Temp = vbYes Then
On Error GoTo Hell
Screen.MousePointer = 11
Text1.SelText = NumericToRoman(Text1.SelText)
Screen.MousePointer = vbDefault
mnuUndo.Enabled = False
Else
Exit Sub
End If
End If


Exit Sub
Hell:
If Err.Number = 6 Then
Screen.MousePointer = vbDefault
MsgBox "Sorry! " & vbCr & "The number is too large!!", vbCritical
mnuUndo.Enabled = False
ElseIf Err.Number = 7 Then
Screen.MousePointer = vbDefault
HellError
Else
Screen.MousePointer = vbDefault
HellError
End If

End Sub

Private Sub mnuNumFigurs_Click()

If TextNotSelected Then Exit Sub
If Not IsNumeric(Text1.SelText) Then
MsgBox "Select numeric text only !", vbExclamation
Exit Sub
End If

If Len(Text1.SelText) > 66 Then 'Checks if they pass the 10^66 barrier
    MsgBox "Sorry!! currently conversion upto 9.99 * 10^66 is supported!", vbInformation
Exit Sub

Else

Temp = MsgBox("Proceed with conversion?" & vbCr & "This action can't be undone!", vbYesNoCancel + vbQuestion)

If Temp = vbYes Then
Convert Text1.SelText
Else
Exit Sub

End If
End If
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuOccurrences_Click()
Dim Ask$
On Error GoTo Hell
Ask = InputBox("Type the string you want to know the occurrences of in the current file." & vbCr & vbCr & "Note:-The Search is Case-Sensitive")
If Ask = "" Then
Exit Sub
Else
MsgBox "The number of OCCURRENCES is:" & vbCrLf & Format(getCountOf(Text1.Text, Ask), "###,###,###,###,###"), vbInformation, "String Occurrences"
End If

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuOlive_Click()
On Error GoTo Hell
Text1.SelText = "OLIVE"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuOpen_Click(index As Integer)

If Sarfraz = True Then
Dim Chandio As Integer
Chandio = MsgBox("Do you want to save the changes to: " & vbCr & Jimmy, vbYesNoCancel + vbQuestion, "Save File?")

If Chandio = vbYes Then
 SaveText
ElseIf Chandio = vbNo Then
Else
 Exit Sub
End If
End If

If Chandio = 6 Then If Sarfraz = True Then Exit Sub

Dim Directory As String

cd.flags = cdlOFNFileMustExist
On Error GoTo Hell
cd.CancelError = True
cd.DialogTitle = "Open"
cd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
On Error GoTo Hell
cd.ShowOpen

Directory = GetDirectory(cd.filename)


Close
Open cd.filename For Input As #1
Text1.Text = Input$(LOF(1), #1)
Text1.Tag = Text1.Text
Jimmy = cd.filename
frmText.Caption = Jimmy & " - String Processor"
StatusBar1.Panels(1).Text = Jimmy
Sarfraz = False
mnuUndo.Enabled = False
Close #1

cd.filename = Jimmy

On Error Resume Next
AddToHis
Exit Sub
Hell:
Close #1

End Sub

Private Sub mnuOptimizeDrives_Click()
On Error Resume Next
Shell "DEFRAG.EXE", vbNormalFocus
End Sub

Private Sub mnuOptions_Click()
frmOptions.Show , Me
End Sub

Private Sub mnuPackager_Click()
On Error Resume Next
Shell "PACKAGER.EXE", vbNormalFocus
End Sub

Private Sub mnuPageSetup_Click()
On Error Resume Next
With cd
.DialogTitle = "Page Setup"
.CancelError = True
.ShowPrinter
End With
End Sub

Private Sub mnuParagraph_Click()
On Error GoTo Hell
Text1.SelText = "<P ALIGN=""LEFT"">"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuPasswordBox_Click()
On Error GoTo Hell
Text1.SelText = "<INPUT TYPE=PASSWORD   NAME=""" + "Password" + """   VALUE=""" + "PasswordString" + """   SIZE=20   MAXLENGTH=50> <P> "
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuPaste_Click()
On Error GoTo Hell
Text1.SelText = Clipboard.GetText
mnuUndo.Enabled = False

Exit Sub
Hell:
If Err.Number = 7 Or Err.Description = "Out of Memory" Then
Close #1
HellError
End If
End Sub

Private Sub mnuPdf_Click()
frmConvertToPDF.Show
End Sub

Private Sub mnupopup_Click()
On Error Resume Next
Shell "drwatson.exe", vbNormalFocus
End Sub

Private Sub mnuPredefined_Click()
On Error GoTo Hell
Text1.SelText = "<PRE>" & vbCrLf & vbCrLf & "</PRE>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuPreview_Click()
Dim Go As Long

On Error GoTo Hell
Open App.path & "\Preview.html" For Output As #1
Print #1, Text1.Text
Close #1
On Error GoTo Hell
Go = ShellExecute(Me.hWnd, "open", App.path & "\Preview.html", "", App.path, 1)

Exit Sub
Hell:
MsgBox Err.Description, vbCritical
End Sub

Private Sub mnuPrint_Click()
Dim Hheight, Hwidth
On Error Resume Next

With cd
.PrinterDefault = True
'Disable printing to file and individual page printing.
.flags = cdlPDDisablePrintToFile Or cdlPDNoPageNums

If Text1.SelLength = 0 Then
'Hide Selection button if there is no selected text.
.flags = .flags Or cdlPDNoSelection
Else
'Else enable the Selection button and make it the default
'choice.
.flags = .flags Or cdlPDSelection
End If

'We need to know whether the user decided to print.
.CancelError = True
.ShowPrinter

If Err = 0 Then
If .flags And cdlPDSelection Then
Printer.Print Text1.SelText

Else

On Error GoTo Hell
Hheight = Printer.TextHeight(Text1.Text)
Hwidth = Printer.TextWidth(Text1.Text)
Printer.CurrentX = 10
Printer.CurrentY = 10
Printer.Print Text1.Text

End If
End If
Printer.EndDoc
End With

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuPrograman_Click()
On Error Resume Next
Shell "PROGMAN.EXE", vbNormalFocus
End Sub

Private Sub mnuProperCapitalize_Click()
If TextNotSelected Then Exit Sub
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbProperCase)
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuProperties_Click()
frmProperties.Show , Me
End Sub

Private Sub mnuPurple_Click()
On Error GoTo Hell
Text1.SelText = "PURPLE"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuQuickRepalce_Click()
If TextNotSelected Then Exit Sub
Dim a$
Dim strText As String
strText = Text1.Text

a = InputBox("Enter the Text you want to replace the selected text with to replace all occurrences." & vbCr & vbCr & "Note:-The text-selection is Case-Sensitive")
If a = "" Then
Exit Sub
Else
On Error GoTo Handler
Screen.MousePointer = 11
Text1.Text = ReplaceText(strText, Text1.SelText, a)
Screen.MousePointer = 0
mnuUndo.Enabled = False

Handler:
If Err.Number = 0 Then
Screen.MousePointer = 0
Exit Sub
Else
HellError
Screen.MousePointer = 0
End If
Exit Sub
End If
End Sub

Private Sub mnuRadio_Click()
On Error GoTo Hell
Text1.SelText = "<INPUT TYPE=RADIO   NAME=""" + "RadioButton" + """   VALUE=""" + "RadioButton" + """   CHECKED><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuReadMe_Click()
If FileExists(App.path & "\StringPro.hlp") Then
cd.HelpFile = App.path & "\StringPro.hlp"
cd.HelpCommand = cdlHelpContents
cd.ShowHelp
Else

Dim Go As Long
On Error GoTo Hell
Go = ShellExecute(Me.hWnd, "open", App.path & "\StringPro Help.html", "", App.path, 1)
End If

Hell:
If Not FileExists(App.path & "\StringPro Help.html") Then
On Error Resume Next
Shell "start http://www.angelfire.com/ultra/sarfraz", vbHide
End If
End Sub

Private Sub mnuRed_Click()
On Error GoTo Hell
Text1.SelText = "RED"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuRemoveExtraSpaces_Click()
If TextNotSelected Then Exit Sub

Screen.MousePointer = 11
On Error Resume Next
Text1.SelText = RemoveExtraSpaces(Text1.SelText)
mnuUndo.Enabled = False
Screen.MousePointer = 0
End Sub

Private Sub mnuRepetitive_Click()
Dim NumChar As String
Dim HowMany As String
On Error GoTo Hell
NumChar = InputBox("Type the character you want repeated : ")
If NumChar = "" Then
Exit Sub
Else
HowMany = InputBox("How many times the character should be repeated? ")
Text1.SelText = String(Val(HowMany), NumChar)
mnuUndo.Enabled = False
End If

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuResetButton_Click()
On Error GoTo Hell
Text1.SelText = "<INPUT TYPE=RESET   NAME=""" + "ResetButton" + """   VALUE=""" + "ResetButton" + """><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuRevertSaved_Click()
If Jimmy = "Untitled" Then Exit Sub
Temp = MsgBox("Are you sure you want to revert the file to original and loose the changes made?", vbQuestion + vbYesNo, "Revert File?")

If Temp = vbYes Then
On Error GoTo Handler
Text1.Text = Text1.Tag
'Sarfraz = False
Me.mnuUndo.Enabled = False
Else
Exit Sub
End If

Exit Sub
Handler:
MsgBox Err.Description, vbCritical
mnuUndo.Enabled = False
End Sub

Private Sub mnuRight_Click()
On Error GoTo Hell
Text1.SelText = "RIGHT"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuRows_Click()
On Error GoTo Hell
Text1.SelText = "ROWS="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuRowSpan_Click()
On Error GoTo Hell
Text1.SelText = "ROWSPAN="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuSampleDate_Click(index As Integer)
Dim txt As String

Select Case index
Case 0 To 8: txt = mnuSampleDate(index).Caption
Case 10: txt = Format(Now, "h:mm:ss am/pm")
Case 11: txt = Format(Now, "h:mm AM/PM")
Case 12: txt = Format(Now, "h:mm:ss")
Case 13: txt = Format(Now, "h:mm")
Case 14 To 16: txt = mnuSampleDate(index).Caption
End Select

On Error GoTo Hell:
Text1.SelText = txt
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuSaveExit_Click()
SaveText
Unload Me
End Sub

Private Sub mnuSaveSelectionAs_Click()

On Error GoTo DOWN
cd.CancelError = True
cd.flags = cdlOFNOverwritePrompt
cd.DialogTitle = "Save Selection As"
cd.DefaultExt = "txt"
cd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
On Error GoTo Hell
cd.ShowSave

Open cd.filename For Output As #1
Print #1, Text1.SelText
StatusBar1.Panels(1).Text = "Selected Text Saved!"
Close #1

cd.filename = Jimmy

Exit Sub
Hell:
cd.filename = Jimmy
DOWN:
If Err.Number = 32755 Then
Exit Sub
Else
HellError
cd.filename = Jimmy
End If
End Sub

Private Sub mnuSaveText_Click(index As Integer)
SaveText
End Sub

Private Sub mnuScanDisk_Click()
On Error Resume Next
Shell "SCANDSKW.EXE ", vbNormalFocus
End Sub

Private Sub mnuScanRegistry_Click()
On Error Resume Next
Shell "C:\WINDOWS\SCANREGW.EXE", vbNormalFocus
End Sub

Private Sub mnuScrolling_Click()
On Error GoTo Hell
Text1.SelText = "SCROLLING=YES"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuSelectAll_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub mnuSelection_Click()
If TextNotSelected Then Exit Sub

Dim lNotepadHwnd As Long
Dim lNotepadEdit As Long
Dim sText As String
    
On Error GoTo Handler
Shell "notepad.exe", vbNormalFocus
    
lNotepadHwnd = FindWindow("Notepad", vbNullString)
lNotepadEdit = FindWindowEx(lNotepadHwnd, 0&, "Edit", vbNullString)

sText = Text1.SelText

SendMessageSTRING lNotepadEdit, WM_SETTEXT, 256, sText

Exit Sub
Handler:
MsgBox "Unknown Error!!" & vbCrLf & "Can't transfer!", vbCritical, "Error"
End Sub

Private Sub mnuSelectLine_Click()
Dim LineNumber As Integer
Dim GetLineText As String

On Error Resume Next
'get line number
LineNumber = SendMessage(frmText.Text1.hWnd, _
EM_LINEFROMCHAR, -1, ByVal 0) + 1

'get current line text
 GetLineText = GetLine(Text1, LineNumber - 1)
 
If GetLineText = "" Then
 Exit Sub
Else
 Clipboard.SetText (GetLineText)
End If

End Sub

Private Sub mnuSFC_Click()
On Error Resume Next
Shell "sfc.exe", vbNormalFocus
End Sub

Private Sub mnuSignature_Click()
On Error GoTo Hell
Text1.SelText = "<ADDRESS>     </ADDRESS>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuSilver_Click()
On Error GoTo Hell
Text1.SelText = "SILVER"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuSize_Click()
On Error GoTo Hell
Text1.SelText = "SIZE="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuSpellChecker_Click()
Dim m_SpellCheck As clsSpellCheck

If m_SpellCheck Is Nothing Then
 Set m_SpellCheck = New clsSpellCheck

 m_SpellCheck.LoadDict App.path & "\SpellCheck.dat"

 m_SpellCheck.CheckTextBox Me.Text1

End If

 Set m_SpellCheck = Nothing

End Sub

Private Sub mnuStart_Click()
On Error GoTo Hell
Text1.SelText = "START"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuStartupDisk_Click()
On Error Resume Next
Shell "C:\Program Files\Plus!\SYSTEM\NOCOMP.EXE", vbNormalFocus
End Sub

Private Sub mnuSTYLE_Click()
On Error GoTo Hell
Text1.SelText = "STYLE="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuSTYLING_Click()
On Error GoTo Hell
Text1.SelText = "<STYLE>" & vbCrLf & vbCrLf & "</STYLE>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuSymbols_Click()
frmSymbols.Show 0, Me
End Sub

Private Sub mnuTable_Click()
On Error GoTo Hell
Text1.SelText = "<TABLE  ALIGN=CENTER  BORDER=1  BGCOLOR=WHITE>" & vbCrLf & "<TR>" & vbCrLf & "   " & "<TD>Row1,Column1</TD>" & vbCrLf & "   " & "<TD>Row1,Column2</TD>" & vbCrLf & "   " & "<TD>Row1,Column3</TD>" & vbCrLf & "</TR>" & vbCrLf & "<TR>" & vbCrLf & "   " & "<TD>Row2,Column1</TD>" & vbCrLf & "   " & "<TD>Row2,Column2</TD>" & vbCrLf & "   " & "<TD>Row2,Column3</TD>" & vbCrLf & "</TR>" & vbCrLf & "   " & "<TD>Row3,Column1</TD>" & vbCrLf & "   " & "<TD>Row3,Column2</TD>" & vbCrLf & "   " & "<TD>Row3,Column3</TD>" & vbCrLf & "</TR>" & vbCrLf & "</TABLE>"
mnuUndo.Enabled = False
Exit Sub
Hell:
HellError
End Sub

Private Sub mnuTeal_Click()
On Error GoTo Hell
Text1.SelText = "TEAL"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuText_Click()
On Error GoTo Hell
Text1.SelText = "<INPUT TYPE=TEXT   NAME=""" + "TextBox" + """   SIZE=20    VALUE=""" + "TextBox" + """<P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuTextArea_Click()
On Error GoTo Hell
Text1.SelText = "<TEXTAREA   NAME=""" + "TextAreaName" + """   COLS=20>   ROWS=2   MAXLENGTH=1000" & vbCrLf & vbCrLf & "</TEXTAREA><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuTextLink_Click()
On Error GoTo Hell
Text1.SelText = "TEXT="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Function HiByte(ByVal wParam As Long)
HiByte = wParam \ &H100 And &HFF&
End Function

Function LoByte(ByVal wParam As Long)
LoByte = wParam And &HFF&
End Function

Private Sub mnuTexttoExe_Click()
Dim a(14) As Byte
Dim i As Long

On Error GoTo Handler
a(0) = 190
a(1) = 15
a(2) = 1
a(3) = 185
a(4) = 0
a(5) = 0
a(6) = 252
a(7) = 172
a(8) = 205
a(9) = 41
a(10) = 73
a(11) = 117
a(12) = 250
a(13) = 205
a(14) = 32

Dim sourcelen

Open Jimmy For Input As #1
sourcelen = LOF(1)
Close #1

a(4) = LoByte(sourcelen)
a(5) = HiByte(sourcelen)

Dim newfilename$

Dim t

newfilename = Left(Jimmy, Len(Jimmy) - 4) & ".exe"

Open Jimmy For Input As #1
Open newfilename For Output As #2
t = Input(LOF(1), 1)

Dim k, St
For k = 0 To 14
St = St & Chr(a(k))
Next k
St = St & t
Print #2, St
Close #1
Close #2

MsgBox "The MS-DOS EXE of the current file was successfully created!!", vbInformation, "Executable file created!!"

Exit Sub
Handler:
Close #1
If Err.Number = 7 Then
MsgBox "File too large!", vbExclamation
Else
MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub mnuTexttoHTML_Click()
Dim a
On Error Resume Next
Text4.Text = Text1.Text

a = Text4.SelStart
Text4.SelStart = Trim(0)

Dim strTitle
Dim strFC$
Dim strBC$
Dim intSize

strTitle = InputBox("Enter the TitleName for the webpage.")
strBC = InputBox("Enter the BackColor name for the webpage." & vbNewLine & vbNewLine & "Please type the correct Spelling!")
strFC = InputBox("Enter the ForeColor name for the webpage." & vbNewLine & vbNewLine & "Please type the correct Spelling!")
intSize = InputBox("Enter the FontSize of the text for webpage.")

If Trim(strBC) = "" And Trim(strFC) = "" And Trim(intSize) = "" And Trim(strTitle) = "" Then
On Error GoTo BigSize
Text4.SelText = "<PRE>" & vbNewLine & "<BODY BGCOLOR = White >" & vbNewLine & "<FONT color = Black  face = " & Text1.FontName & "  " & "Size = 2" & "</FONT>"
Else
On Error GoTo BigSize
Text4.SelText = "<PRE>" & vbNewLine & "<Title>" & strTitle & "</Title>" & "<BODY BGCOLOR=" & strBC & ">" & vbNewLine & "<FONT color = " & strFC & "  " & "face = " & Text1.FontName & "  " & "Size = " & intSize & "</FONT>"
End If

With frmText.cd
On Error GoTo Handler
.CancelError = True
.DefaultExt = "htm"
.filename = ""
.flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
.DialogTitle = "Save As"
.Filter = "HTML Document (*.htm;*.html)|*.htm;*.html|"
.ShowSave

Open cd.filename For Output As #1
Print #1, Text4.Text

cd.filename = Jimmy

Close #1
End With


cd.filename = Jimmy


Exit Sub
Handler:
Close #1
cd.filename = Jimmy

Exit Sub
BigSize:
If Err.Number = 7 Or Err.Description = "Out of Memory" Then
Close #1
MsgBox "Text file is too large!! ", vbExclamation, "Can't convert!"
Else
Close #1
MsgBox Err.Description, vbCritical, "Error"
End If

End Sub

Private Sub mnuTFoot_Click()
On Error GoTo Hell
Text1.SelText = "TFOOT="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuToolbar_Click()
Toolbar1.Visible = Not Toolbar1.Visible

If Toolbar1.Visible Then
Text1.Top = 445
Text1.Height = 7575
mnuToolbar.Checked = True
Else
Text1.Top = 1
Text1.Height = 7800
mnuToolbar.Checked = False
End If

End Sub

Private Sub mnuTypeDisc_Click()
On Error GoTo Hell
Text1.SelText = "TYPE=DISC"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuTypeWriterText_Click()
On Error GoTo Hell
Text1.SelText = "<TT>     </TT>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuTypingSound_Click()
mnuTypingSound.Checked = Not mnuTypingSound.Checked

If mnuTypingSound.Checked = True Then
On Error Resume Next
UseSound = "Yes"
Else
UseSound = ""
End If

End Sub

Private Sub mnuUnderline_Click()
On Error GoTo Hell
Text1.SelText = "<U>" & "   " & "</U>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuUndo_Click()
SendMessage frmText.Text1.hWnd, EM_UNDO, 0, 0&
End Sub

Private Sub mnuUpperCase_Click()
If TextNotSelected Then Exit Sub
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbUpperCase)
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuUseMap_Click()
On Error GoTo Hell
Text1.SelText = "USEMAP="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuVAlign_Click()
On Error GoTo Hell
Text1.SelText = "VALIGN="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuVb_Click()
On Error GoTo Hell
Text1.SelText = "<SCRIPT LANGUAGE =""" + "VBScript" + """>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</SCRIPT>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuVLINK_Click()
On Error GoTo Hell
Text1.SelText = "VLINK="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuVSpace_Click()
On Error GoTo Hell
Text1.SelText = "VSPACE="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuWebLink_Click()
On Error GoTo Hell
Text1.SelText = "<A HREF=""WebsiteName"">Description</A><P>"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuWhite_Click()
On Error GoTo Hell
Text1.SelText = "WHITE"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuWidth_Click()
On Error GoTo Hell
Text1.SelText = "WIDTH="
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub mnuWordCount_Click()
frmWordCount.Show 1
End Sub

Private Sub mnuYellow_Click()
On Error GoTo Hell
Text1.SelText = "YELLOW"
mnuUndo.Enabled = False

Exit Sub
Hell:
HellError
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
Sarfraz = True
End If

If Sarfraz = True Then
mnuUndo.Enabled = True
Else
mnuUndo.Enabled = False
End If

If UseSound = "Yes" Then
Dim Play As String
On Error Resume Next
Play = sndPlaySound(App.path + "\TypingSound.wav", SND_ASYNC)
End If
End Sub

Private Sub Timer1_Timer()
Static count As Integer

If count = 0 Then
StatusBar1.Panels(1).Text = "String Processor v1.00"
count = 1
Exit Sub

ElseIf count = 1 Then
StatusBar1.Panels(1).Text = "Process The Text In A Professional Style..."
count = 2
Exit Sub

ElseIf count = 2 Then
StatusBar1.Panels(1).Text = "Convert Text To PDF Format!!"
count = 3
Exit Sub

ElseIf count = 3 Then
StatusBar1.Panels(1).Text = "Encrypt/Decrypt Text..."
count = 4
Exit Sub

ElseIf count = 4 Then
StatusBar1.Panels(1).Text = "On-Screen File-Properties!!"
count = 5
Exit Sub

ElseIf count = 5 Then
StatusBar1.Panels(1).Text = "Superb WinTools A Click Away..."
count = 6
Exit Sub

ElseIf count = 6 Then
StatusBar1.Panels(1).Text = "Spell-Checking to produce error-free files!!"
count = 7
Exit Sub

ElseIf count = 7 Then
StatusBar1.Panels(1).Text = "Find and Replace the Text!!"
count = 8
Exit Sub

ElseIf count = 8 Then
StatusBar1.Panels(1).Text = "Count Words,Characters,Spaces and Lines..."
count = 9
Exit Sub

ElseIf count = 9 Then
StatusBar1.Panels(1).Text = "Convert Numbers to Words upto 9.99 * 10^66"
count = 10
Exit Sub

ElseIf count = 10 Then
StatusBar1.Panels(1).Text = "Case-Select The Text..."
count = 11
Exit Sub

ElseIf count = 11 Then
StatusBar1.Panels(1).Text = "Insert Repetitive Characters!!"
count = 12
Exit Sub

ElseIf count = 12 Then
StatusBar1.Panels(1).Text = "Insert Symbols Easily!!"
count = 13
Exit Sub

ElseIf count = 13 Then
StatusBar1.Panels(1).Text = "Choose from various Time and Date formats!!"
count = 14
Exit Sub

ElseIf count = 14 Then
StatusBar1.Panels(1).Text = "Convert Numbers to Romans upto Ten Million!!"
count = 15
Exit Sub

ElseIf count = 15 Then
StatusBar1.Panels(1).Text = "Use """ + "HTML Editing" + """ feature to make WebPages!!"""
count = 16
Exit Sub

ElseIf count = 16 Then
StatusBar1.Panels(1).Text = "Play ""TypingSound"" with typewriter effect!"
count = 17
Exit Sub

ElseIf count = 17 Then
StatusBar1.Panels(1).Text = "There is more to StringPro..."
count = 18
Exit Sub

ElseIf count = 18 Then
StatusBar1.Panels(1).Text = "Please Feel Free To Distribute StringPro..."
count = 19
Exit Sub

ElseIf count = 19 Then
StatusBar1.Panels(1).Text = "Thank You For Using StringPro..."
count = 20
Exit Sub

ElseIf count = 20 Then
StatusBar1.Panels(1).Text = "For more Info,please visite: www.angelfire.com/ultra/sarfraz"
count = 21
Exit Sub

ElseIf count = 21 Then
StatusBar1.Panels(1).Text = "Developed By:- Sarfraz Ahmed Chandio"
count = 22
Exit Sub

ElseIf count = 22 Then
StatusBar1.Panels(1).Text = "Manipulate the Text Easily and Quickly..."
count = 23
Exit Sub

ElseIf count = 23 Then
StatusBar1.Panels(1).Text = "Revert to original file without any effort!!"
count = 24
Exit Sub

ElseIf count = 24 Then
StatusBar1.Panels(1).Text = "Manage the opened file easily..."
count = 25
Exit Sub

ElseIf count = 25 Then
StatusBar1.Panels(1).Text = "Use the ""Compare Text"" option to compare text files!"
count = 26
Exit Sub

ElseIf count = 26 Then
StatusBar1.Panels(1).Text = "Use ""Find in Files"" to find text in other files!"
count = 27
Exit Sub

ElseIf count = 27 Then
StatusBar1.Panels(1).Text = "Use amazing feature ""Text to EXE"" to convert text to MS-DOS EXE!"
count = 28
Exit Sub

ElseIf count = 28 Then
StatusBar1.Panels(1).Text = "Create WebPages in no time with ""Text to HTML"" feature!!"
count = 0
Exit Sub

End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Toolbar1.AllowCustomize = False

On Error Resume Next
Select Case Button.Key
Case "Symbols"
    Call mnuSymbols_Click
Case "Help"
    Call mnuReadMe_Click
Case "About"
    Call mnuAbout_Click
Case "Convert To PDF"
    Call mnuPdf_Click
Case "Encrypt/Decrypt"
    Call mnuEDFile_Click
Case "Spell Check"
    Call mnuSpellChecker_Click
Case "Delete"
    Call mnuClear_Click
Case "Print"
    Call mnuPrint_Click
Case "New"
    Call mnuNew_Click(2)
Case "Open"
    Call mnuOpen_Click(1)
Case "Save"
    Call SaveText
Case "Copy"
    Call mnuCopy_Click
Case "Cut"
    Call mnuCut_Click
Case "Paste"
    Call mnuPaste_Click
Case "Find"
    Call mnuFind_Click
End Select
End Sub

Private Sub WatchDog_Timer()
Dim Nawaz As String

If TextSelected Then
Toolbar1.Buttons.Item(6).Enabled = True
Toolbar1.Buttons.Item(9).Enabled = True
Toolbar1.Buttons.Item(10).Enabled = True
mnuSaveSelectionAs.Enabled = True
mnuCut.Enabled = True
mnuCopy.Enabled = True
mnuClear.Enabled = True
Else
Toolbar1.Buttons.Item(6).Enabled = False
Toolbar1.Buttons.Item(9).Enabled = False
Toolbar1.Buttons.Item(10).Enabled = False
mnuSaveSelectionAs.Enabled = False
mnuCut.Enabled = False
mnuCopy.Enabled = False
mnuClear.Enabled = False
End If

If Text1.Text <> "" Then
Toolbar1.Buttons.Item(7).Enabled = True
mnuGoto.Enabled = True
mnuFind.Enabled = True
mnuQuickRepalce.Enabled = True
mnuSelectAll.Enabled = True
mnuOccurrences.Enabled = True
mnuFullScreen.Enabled = True
mnuClearAll.Enabled = True
mnuSelectLine.Enabled = True
Else
Toolbar1.Buttons.Item(7).Enabled = False
mnuGoto.Enabled = False
mnuFind.Enabled = False
mnuQuickRepalce.Enabled = False
mnuSelectAll.Enabled = False
mnuOccurrences.Enabled = False
mnuFullScreen.Enabled = False
mnuClearAll.Enabled = False
mnuSelectLine.Enabled = False
End If

If Text1.Text <> Text1.Tag And Jimmy <> "Untitled" And Text1.Tag <> "" Then
mnuRevertSaved.Enabled = True
Else
mnuRevertSaved.Enabled = False
End If

If Jimmy <> "Untitled" And FileExists(cd.filename) Then
mnuProperties.Enabled = True
mnuTexttoExe.Enabled = True
mnuFileManagement.Enabled = True
mnuEDFile.Enabled = True
mnuPdf.Enabled = True
Toolbar1.Buttons(15).Enabled = True
Toolbar1.Buttons(16).Enabled = True
Else
mnuEDFile.Enabled = False
mnuPdf.Enabled = False
Toolbar1.Buttons(15).Enabled = False
Toolbar1.Buttons(16).Enabled = False
mnuProperties.Enabled = False
mnuTexttoExe.Enabled = False
mnuFileManagement.Enabled = False
End If


If Jimmy <> "Untitled" Then
mnuCompareFiles.Enabled = True

Else
mnuCompareFiles.Enabled = False
End If

On Error Resume Next
Nawaz = Clipboard.GetText
If Nawaz = "" Then
Toolbar1.Buttons.Item(11).Enabled = False
mnuPaste.Enabled = False
Else
Toolbar1.Buttons.Item(11).Enabled = True
mnuPaste.Enabled = True
End If

On Error Resume Next
StatusBar1.Panels(2).Text = "Line#" & Format(SendMessage(frmText.Text1.hWnd, EM_LINEFROMCHAR, -1, _
ByVal 0) + 1, "###,###,###,###")


On Error Resume Next
StatusBar1.Panels(3).Text = "Lines:" & Format(SendMessage(Text1.hWnd, EM_GETLINECOUNT, 0, 0&), "###,###,###,###")

If mnuHtml.Visible = True Then
mnuHtmlEditing.Checked = True
Else
mnuHtmlEditing.Checked = False
End If

End Sub

'Counts the number of occurrences of a given string
Function getCountOf(OriginalString As String, StringToLookFor As String) As Long
On Error GoTo Hell
Dim i As Long
getCountOf = 0  'initilaise the return value as 0

i = 1 ' set it as the first
On Error GoTo Hell
While i <> 0
i = InStr(i, OriginalString, StringToLookFor) ' set it to the next
If i <> 0 Then ' if found then
i = i + 1   ' set it to the found place +1
getCountOf = getCountOf + 1 'increment the count by one
End If
Wend
Exit Function

Hell:
MsgBox Err.Description, vbCritical
End Function

Public Sub AddToHis()
If Jimmy = mnuHis(0).Caption Or Jimmy = mnuHis(1).Caption Or Jimmy = mnuHis(2).Caption Or Jimmy = mnuHis(3).Caption Or Jimmy = mnuHis(4).Caption Then
Exit Sub
Else
For i = 0 To 4
If History(i) = "" Then
History(i) = cd.filename
mnuHis(i).Caption = History(i)
On Error GoTo theError
mnuHis(i).Visible = True
Sep.Visible = True
Exit Sub
End If
Next
On Error GoTo theError
i = GetSetting(App.EXEName, "options", "add", 0)
History(i) = cd.filename
mnuHis(i).Caption = History(i)
mnuHis(i).Visible = True
Sep.Visible = True

i = i + 1
If i >= 5 Then i = 0
SaveSetting App.EXEName, "options", "add", i
End If

Exit Sub
theError:
MsgBox Err.Description, vbInformation, "Error"
End Sub

'Removes extra spaces from textbox
Function RemoveExtraSpaces(TheString As String) As String
Dim LastChar As String
Dim NextChar As String

LastChar = Left(TheString, 1)
RemoveExtraSpaces = LastChar

For i = 2 To Len(TheString)
NextChar = Mid(TheString, i, 1)

If NextChar = " " And LastChar = " " Then
Else
RemoveExtraSpaces = RemoveExtraSpaces & NextChar
End If
LastChar = NextChar
Next i

End Function

'Replaces Text
Function ReplaceText(Text As String, TextToReplace As String, NewText As String) As String
Dim mtext As String, SpacePos As Long
mtext = Text
SpacePos = InStr(mtext, TextToReplace)
Do While SpacePos
mtext = Left(mtext, SpacePos - 1) & NewText & Mid(mtext, SpacePos + Len(TextToReplace))
SpacePos = InStr(SpacePos + Len(NewText), mtext, TextToReplace)
Loop
ReplaceText = mtext
End Function

'Convert numbers to words upto "9.99 * 10^66"
Sub Convert(sStr As String)
Dim x As Integer
Dim sText As String
Dim T1 As Integer
Dim Bot, Top As Integer
Dim Neg, Dol As String
Dim TempChars As String
Dim LenChars As Integer
Dim Lenght As Integer

If Left(sStr, 1) = "$" Then 'Checks if it is in dollars
sStr = Right(sStr, Len(sStr) - 1)   'Removes the dollar sign
Dol = " Dollars"                    'Adds the 'Dollars' Flag
End If

If Int(sStr) < 0 Then   'Checks if the number is negative
sStr = Right(sStr, Len(sStr) - 1) 'Turns number positive
Neg = "Negative "   'Adds the 'Negative' Flag
End If

TempChars = Flip(sStr) 'Takes the number and flips it so that the ones come first

LenChars = Len(sStr) 'Finds how long the Number is
Lenght = Int(LenChars / 3 + 2 / 3) 'Calulates how many places (powers of 10)
'Returns 1 if < 1000 2 if less than 1000000...

For x = 1 To Lenght
Bot = 3 * x - 2 'Sets the bottom barrier
Top = 3 * x     'Sets the top barrier
If Top > LenChars Then Top = LenChars 'Checks that the top does not exceed the amount of charachters
'Cuts the 3 numbers,flips then so that tehy are in correct order, convers them to decimal
T1 = Int(Flip(Mid(TempChars, Bot, Top - Bot + 1))) 'Derives numbers
sText = Places(Int(T1), x) & " " & sText 'Calls function to turn nums to text
Next
sText = Trim(sText) 'Removes unnessesary spaces
sText = Neg & sText 'If negative flag, then adde 'Negative'
sText = sText & Dol 'If dollars flag, then adds 'Dollars'
On Error GoTo Hell
Text1.SelText = sText 'Cuts of unneccesary spaces
If Int(sStr) = 0 Then Text1.SelText = "Zero" 'If number was 0

Exit Sub
Hell:
HellError
End Sub
Function Flip(St As String) As String 'Funtion flips string 'hello' into 'olleh'
Dim x As Long
For x = Len(St) To 1 Step -1
Flip = Flip & Mid(St, x, 1)
Next
End Function
Function Places(Num As Integer, Ln As Integer) As String
Dim D1 As Integer
Dim D2 As Integer
Dim D3 As Integer
Dim Labels(23) As String
Labels(1) = ""              'Declare place identifiers
Labels(2) = "Thousand"
Labels(3) = "Million"
Labels(4) = "Billion"
Labels(5) = "Trillion"
Labels(6) = "Quadrillion"  'From this one down, I found these at a website
Labels(7) = "Quintillion"
Labels(8) = "Sextillion"
Labels(9) = "Septillion"
Labels(10) = "Octillion"
Labels(11) = "Nonillion"
Labels(12) = "Decillion"
Labels(13) = "Undecillion"
Labels(14) = "Duodecillion"
Labels(15) = "Tredecillion"
Labels(16) = "Quatuordecillion"
Labels(17) = "Quindecillion"
Labels(18) = "Sexdecillion"
Labels(19) = "Septendecillion"
Labels(20) = "Octodecillion"
Labels(21) = "Novemdecillion"
Labels(22) = "Vigintillion"
D3 = Num Mod 10                 'Calculates the ones place
D2 = ((Num Mod 100) - D3) / 10  'Calculates the tens place
D1 = (Num - D2 * 10 - D3) / 100 'Calculates the hundreds place

If D1 <> 0 Then Places = Places & TT(D1) & " Hundred"  'Convers hudreds place to text
If D2 <> 0 Then
If D2 = 1 Then  'If the number is between 10-19
Places = Places & " " & TT(D2 * 10 + D3) 'So that it is 'Nineteen' instead of 'Ten Nine'
D3 = 0  'Turns ones place into zero so its doesnt print twice 'Nineteen Nine'
Else  'If the number is not 10-19
Places = Places & " " & TT(D2 * 10) 'Does tens seperately 'Twenty'
Places = Places & " " & TT(D3)      'Does ones seperately 'Nine'
D3 = 0
End If
End If
If D3 <> 0 Then Places = Places & " " & TT(D3) 'If Tens were 0 it prints the ones place
Places = Places & " " & Labels(Ln) 'Add the label (Thousand, Million, ect..)
End Function
Function TT(Num As Integer) As String
Select Case Num 'A string for every special number (1-19, and 20, 30, ..., 90 but not 0)
Case 1
TT = "One"
Case 2
TT = "Two"
Case 3
TT = "Three"
Case 4
TT = "Four"
Case 5
TT = "Five"
Case 6
TT = "Six"
Case 7
TT = "Seven"
Case 8
TT = "Eight"
Case 9
TT = "Nine"
Case 10
TT = "Ten"
Case 11
TT = "Eleven"
Case 12
TT = "Twelve"
Case 13
TT = "Thirteen"
Case 14
TT = "Fourteen"
Case 15
TT = "Fifteen"
Case 16
TT = "Sixteen"
Case 17
TT = "Seventeen"
Case 18
TT = "Eighteen"
Case 19
TT = "Nineteen"
Case 20
TT = "Twenty"
Case 30
TT = "Thirty"
Case 40
TT = "Forty"
Case 50
TT = "Fifty"
Case 60
TT = "Sixty"
Case 70
TT = "Seventy"
Case 80
TT = "Eighty"
Case 90
TT = "Ninety"
End Select

End Function

Public Sub SaveText()
If Sarfraz = True Then
If Jimmy = "Untitled" Then
SaveTextAs
Else
On Error GoTo Hell
Open Jimmy For Output As #1
Print #1, Text1.Text
StatusBar1.Panels(1).Text = "File Saved!"
Sarfraz = False
Close #1
End If
End If

Exit Sub
Hell:
Temp = GetAttr(cd.filename)
If (Temp And vbReadOnly) <> 0 Then
SaveTextAs
Else
SaveText
End If

End Sub

Public Sub SaveTextAs()
Dim Directory As String

On Error GoTo DOWN
cd.CancelError = True
cd.flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
cd.DialogTitle = "Save As"
cd.DefaultExt = "txt"
cd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
On Error GoTo Hell
cd.ShowSave
Directory = GetDirectory(cd.filename)

Open cd.filename For Output As #1
Print #1, Text1.Text
Jimmy = cd.filename
On Error Resume Next
AddToHis
frmText.Caption = Jimmy & " - String Processor"
StatusBar1.Panels(1).Text = "File Saved!"
Sarfraz = False
Close #1

Exit Sub
Hell:
DOWN:
If Err.Number = 32755 Then
CancelClicked = True
Exit Sub
Else
MsgBox Err.Description, vbCritical
End If

End Sub

Function GetLine(TB As TextBox, ByVal lineNum As Long) As String
Dim charOffset As Long, linelen As Long
    
 ' Retrieve the character offset of the first character of the line.
 charOffset = SendMessageByVal(TB.hWnd, EM_LINEINDEX, lineNum, 0)
 ' Now it's possible to retrieve the length of the line.
 linelen = SendMessageByVal(TB.hWnd, EM_LINELENGTH, charOffset, 0)
 ' Extract the line text.
 GetLine = Mid$(TB.Text, charOffset + 1, linelen)
    
End Function


