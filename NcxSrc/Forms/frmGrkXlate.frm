VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGrkXlate 
   Caption         =   "New Covenant Bible Greek Translator"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   945
   ClientWidth     =   11085
   Icon            =   "frmGrkXlate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   11085
   Begin VB.ListBox lstImputBox 
      Height          =   255
      Left            =   7080
      Sorted          =   -1  'True
      TabIndex        =   93
      Top             =   7380
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "Redisplay Greek word definition (ESC)"
      Top             =   660
      Width           =   855
   End
   Begin VB.PictureBox picVbar1 
      Height          =   6615
      Left            =   1660
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6555
      ScaleWidth      =   60
      TabIndex        =   36
      ToolTipText     =   "Drag to resize, double-click to reset"
      Top             =   780
      Width           =   120
   End
   Begin VB.PictureBox picVbar2 
      Height          =   6195
      Left            =   6780
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6135
      ScaleWidth      =   60
      TabIndex        =   41
      ToolTipText     =   "Drag to resize, double-click to reset"
      Top             =   840
      Width           =   120
   End
   Begin VB.PictureBox picHbar1 
      Height          =   120
      Left            =   1860
      MousePointer    =   7  'Size N S
      ScaleHeight     =   60
      ScaleWidth      =   8775
      TabIndex        =   37
      ToolTipText     =   "Drag to resize, double-click to reset"
      Top             =   3720
      Width           =   8835
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   780
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":151A
            Key             =   "goto"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":1F2C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":7B4E
            Key             =   "backup"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":A300
            Key             =   "search"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":A45A
            Key             =   "vine"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":A8AC
            Key             =   "words"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":ACFE
            Key             =   "rebuild"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":AE58
            Key             =   "fav"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":1235A
            Key             =   "help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":12C34
            Key             =   "keyboard"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":12F4E
            Key             =   "prevnote"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":13DA0
            Key             =   "nextnote"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":14BF2
            Key             =   "prevtheo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":15044
            Key             =   "nexttheo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":15496
            Key             =   "prevgreek"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":158E8
            Key             =   "nextgreek"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":15D3A
            Key             =   "kjvdict"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":1618C
            Key             =   "wordpad"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrkXlate.frx":16FDE
            Key             =   "notepad"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   14.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Find next match in List"
      Top             =   660
      Width           =   375
   End
   Begin VB.CommandButton cmdFindInText 
      Height          =   255
      Left            =   9720
      Picture         =   "frmGrkXlate.frx":17E30
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "Search for a word or phrase in the text below (Ctrl-S)"
      Top             =   660
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   960
      Index           =   10
      Left            =   10860
      ScaleHeight     =   900
      ScaleWidth      =   870
      TabIndex        =   76
      Top             =   5700
      Width           =   930
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "goto"
            Object.ToolTipText     =   "Go to a specific book, chapter, and verse"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save the current data image to the main database (Ctrl-S)"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "backup"
            Object.ToolTipText     =   "Make a Backup copy of the current data image (F9)"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "wordpad"
            Object.ToolTipText     =   "Lauch Wordpad text editor"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "notepad"
            Object.ToolTipText     =   "Launch Notepad text editor"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            Object.ToolTipText     =   "Search for words in the Bible (Ctrl-F)"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "vine"
            Object.ToolTipText     =   "View the Vine Dictionary list"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "words"
            Object.ToolTipText     =   "View the complete Synonym list"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kjvdict"
            Object.ToolTipText     =   "View the KJV Dictionary"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fav"
            Object.ToolTipText     =   "Add the current verse to the favorites list"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "rebuild"
            Object.ToolTipText     =   "Resort all synonym lists without disrupting indexing"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "prevnote"
            Object.ToolTipText     =   "Find previous Personal Note"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nextnote"
            Object.ToolTipText     =   "Find next Personal Note"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "prevtheo"
            Object.ToolTipText     =   "Find previous verse with Theological Notes"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nexttheo"
            Object.ToolTipText     =   "Find next verse with Theological Notes"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "prevgreek"
            Object.ToolTipText     =   "Find previous verse without Greek text"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nextgreek"
            Object.ToolTipText     =   "Find next verse without Greek text"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyboard"
            Object.ToolTipText     =   "Keyboard navigation help"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Get help on how to use this program (F1)"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "go"
            Description     =   "GotoBCV"
            Style           =   4
            Object.Width           =   5300
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox picBCV 
         Height          =   315
         Left            =   6420
         ScaleHeight     =   255
         ScaleWidth      =   2955
         TabIndex        =   85
         Top             =   0
         Width           =   3015
         Begin VB.CommandButton cmdGo 
            Caption         =   "Go"
            Height          =   255
            Left            =   2505
            TabIndex        =   89
            ToolTipText     =   "Go to the selected Book, Chapter, and Verse"
            Top             =   0
            Width           =   435
         End
         Begin VB.ComboBox cboVrs 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   88
            ToolTipText     =   "Select Verse"
            Top             =   0
            Width           =   615
         End
         Begin VB.ComboBox cboChp 
            Height          =   315
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   87
            ToolTipText     =   "Select Chapter"
            Top             =   0
            Width           =   615
         End
         Begin VB.ComboBox cboBk 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   86
            ToolTipText     =   "Select Book"
            Top             =   0
            Width           =   1275
         End
      End
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   960
      Index           =   9
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":183BA
      ScaleHeight     =   900
      ScaleWidth      =   870
      TabIndex        =   74
      Top             =   5100
      Width           =   930
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   8055
      Index           =   8
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":18753
      ScaleHeight     =   7995
      ScaleWidth      =   4605
      TabIndex        =   72
      Top             =   4680
      Width           =   4665
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   1380
      Index           =   7
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":1DD07
      ScaleHeight     =   1320
      ScaleWidth      =   1335
      TabIndex        =   71
      Top             =   4140
      Width           =   1395
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   2055
      Index           =   6
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":1E243
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   70
      Top             =   3540
      Width           =   2055
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   4845
      Index           =   5
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":1ECC8
      ScaleHeight     =   4785
      ScaleWidth      =   2265
      TabIndex        =   69
      Top             =   3000
      Width           =   2325
   End
   Begin VB.ListBox lstVineNum 
      Height          =   450
      Left            =   6480
      TabIndex        =   68
      Top             =   7080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ListBox lstVine 
      Height          =   450
      Left            =   5685
      TabIndex        =   67
      Top             =   7080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ListBox lstSort 
      Height          =   450
      Left            =   4890
      Sorted          =   -1  'True
      TabIndex        =   66
      Top             =   7080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Timer tmrAutoBackup 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1860
      Top             =   7080
   End
   Begin VB.TextBox txtMerge 
      Height          =   450
      Left            =   8160
      TabIndex        =   61
      Text            =   "txtMerge"
      Top             =   7080
      Visible         =   0   'False
      Width           =   795
   End
   Begin ComctlLib.ProgressBar pgrWrite 
      Height          =   255
      Left            =   2340
      TabIndex        =   60
      Top             =   7200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   1260
      Index           =   4
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":20134
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   58
      Top             =   2520
      Width           =   1260
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   1965
      Index           =   3
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":20E10
      ScaleHeight     =   1905
      ScaleWidth      =   1905
      TabIndex        =   57
      Top             =   1860
      Width           =   1965
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   1980
      Index           =   2
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":21812
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   56
      Top             =   1200
      Width           =   1980
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   1995
      Index           =   1
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":21EF5
      ScaleHeight     =   1935
      ScaleWidth      =   1935
      TabIndex        =   55
      Top             =   540
      Width           =   1995
   End
   Begin RichTextLib.RichTextBox rtbMerge 
      Height          =   450
      Left            =   7320
      TabIndex        =   49
      Top             =   7080
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   794
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmGrkXlate.frx":2324C
   End
   Begin VB.PictureBox picTop 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   12795
      TabIndex        =   48
      Top             =   420
      Width           =   12855
      Begin VB.Label lblVineRef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Double-click Words, Strong #'s, or underlined verses to access them"
         Height          =   195
         Left            =   7020
         TabIndex        =   73
         ToolTipText     =   "Double-click Words, Strong #'s, or underlined verses to access them"
         Top             =   0
         Width           =   4845
      End
      Begin VB.Label lblTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip: Learning the Shortcut keys will make processing a WHOLE LOT FASTER."
         Height          =   195
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   5580
      End
   End
   Begin VB.PictureBox picTile 
      AutoSize        =   -1  'True
      Height          =   2580
      Index           =   0
      Left            =   10860
      Picture         =   "frmGrkXlate.frx":232C2
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   47
      Top             =   60
      Width           =   2580
   End
   Begin VB.ListBox lstVNotes 
      Height          =   450
      Left            =   4065
      TabIndex        =   46
      Top             =   7080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ListBox lstGrk 
      Height          =   450
      Left            =   3240
      TabIndex        =   44
      Top             =   7080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox picNotes 
      BackColor       =   &H80000010&
      Height          =   2715
      Left            =   6960
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   40
      Top             =   900
      Width           =   3735
      Begin VB.CommandButton cmdAnalysis 
         Caption         =   "Anal&ysis"
         Height          =   375
         Left            =   2820
         TabIndex        =   22
         ToolTipText     =   "Fill Definition panel with complete breakdown of all words in the verse"
         Top             =   2100
         Width           =   795
      End
      Begin VB.CommandButton cmdCopyDef 
         Caption         =   "Copy"
         Height          =   375
         Left            =   0
         TabIndex        =   21
         ToolTipText     =   "Copy Definition to the clipboard"
         Top             =   2040
         Width           =   795
      End
      Begin RichTextLib.RichTextBox rtbNotes 
         Height          =   1935
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3413
         _Version        =   393217
         BackColor       =   -2147483626
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         Appearance      =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmGrkXlate.frx":24122
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblNoteInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notice that often the root word may not look much like the inflected rendition."
         Height          =   195
         Left            =   0
         TabIndex        =   53
         Top             =   2400
         Width           =   5430
      End
   End
   Begin VB.PictureBox picEditor 
      BackColor       =   &H80000010&
      Height          =   3195
      Left            =   1860
      ScaleHeight     =   3135
      ScaleWidth      =   4815
      TabIndex        =   39
      Top             =   3900
      Width           =   4875
      Begin VB.PictureBox picVbar4 
         Height          =   2595
         Left            =   2100
         MousePointer    =   9  'Size W E
         ScaleHeight     =   2535
         ScaleWidth      =   60
         TabIndex        =   65
         ToolTipText     =   "Drag to resize, double-click to reset"
         Top             =   0
         Width           =   120
      End
      Begin VB.CommandButton cmdFind 
         Height          =   375
         Left            =   2280
         Picture         =   "frmGrkXlate.frx":241A2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Search for  a referenced English word (or Ref #) in the Vine database"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdVine 
         Height          =   375
         Left            =   4260
         Picture         =   "frmGrkXlate.frx":26944
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "View Vine Reference regarding this word"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdCopyDT 
         Caption         =   "Copy"
         Height          =   375
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "Copy Direct Translation to the clipboard"
         Top             =   1260
         Width           =   795
      End
      Begin VB.CommandButton cmdCopympv 
         Caption         =   "Copy &MVP"
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         ToolTipText     =   "Copy from your personal version entry to the edit field"
         Top             =   2700
         Width           =   915
      End
      Begin VB.CommandButton cmdCpyXlt 
         Caption         =   "Copy &Xlt."
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         ToolTipText     =   "Copy the Direct Translation to the edit field"
         Top             =   2700
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdateMPV 
         Caption         =   "&Update MPV"
         Height          =   375
         Left            =   60
         TabIndex        =   12
         ToolTipText     =   "Update this verse to My Personal Version"
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Del."
         Height          =   375
         Left            =   3780
         TabIndex        =   18
         ToolTipText     =   "Remove a selected  un-used synonym entry"
         Top             =   2460
         Width           =   495
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         ToolTipText     =   "Edit the selected synonym entry to define a more useful synonym, or fix a goofed added word"
         Top             =   2460
         Width           =   495
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         ToolTipText     =   "Add a synonym to the word list (Add * to ignore minor words)"
         Top             =   2460
         Width           =   495
      End
      Begin VB.ListBox lstWords 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   2280
         MouseIcon       =   "frmGrkXlate.frx":2720E
         MousePointer    =   99  'Custom
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Top             =   360
         Width           =   2535
      End
      Begin RichTextLib.RichTextBox rtbTranslate 
         Height          =   1275
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2249
         _Version        =   393217
         BackColor       =   -2147483626
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   99
         Appearance      =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmGrkXlate.frx":27360
         MouseIcon       =   "frmGrkXlate.frx":27415
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbUser 
         Height          =   975
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   "Edit Personal Version Text Here"
         Top             =   1620
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1720
         _Version        =   393217
         BackColor       =   -2147483626
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmGrkXlate.frx":27577
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblEditPersonal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edit transliteration  below"
         Height          =   195
         Left            =   3060
         TabIndex        =   77
         ToolTipText     =   "Edit transliteration  below"
         Top             =   2880
         Width           =   1755
      End
      Begin VB.Label lblFeminine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Fem?)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   63
         ToolTipText     =   "Possibly FEMININE. Click this caption for more information on Gender Checking..."
         Top             =   1320
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblPlural 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(PL?)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1500
         TabIndex        =   62
         ToolTipText     =   "Possibly PLURAL. Click this caption for more information on Plurality Checking..."
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblSynonyms 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Synonyms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2220
         TabIndex        =   50
         ToolTipText     =   "English synonym(s) for the selected Greek word"
         Top             =   0
         Width           =   2760
      End
   End
   Begin VB.PictureBox picVerse 
      BackColor       =   &H80000010&
      Height          =   3135
      Left            =   6960
      ScaleHeight     =   3075
      ScaleWidth      =   3675
      TabIndex        =   38
      Top             =   3900
      Width           =   3735
      Begin VB.CommandButton cmdWBS 
         Caption         =   "W&BS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2370
         TabIndex        =   91
         ToolTipText     =   "View Webster's Translation (1833)"
         Top             =   2040
         Width           =   435
      End
      Begin VB.CommandButton cmdDBY 
         Caption         =   "&DBY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   435
         TabIndex        =   90
         ToolTipText     =   "View Dary's Translation (1884)"
         Top             =   2040
         Width           =   435
      End
      Begin VB.CommandButton cmdASV 
         Caption         =   "A&SV"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   82
         ToolTipText     =   "View American Standard Version (1901)"
         Top             =   2040
         Width           =   435
      End
      Begin VB.CommandButton cmdWEB 
         Caption         =   "&WEB"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2745
         TabIndex        =   81
         ToolTipText     =   "View World English Bible (1907)"
         Top             =   2040
         Width           =   435
      End
      Begin VB.CommandButton cmdMKJV 
         Caption         =   "&MKJV"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1185
         TabIndex        =   80
         ToolTipText     =   "View Modern King James Version (1962)"
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdKJV 
         Caption         =   "&KJV"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   810
         TabIndex        =   26
         ToolTipText     =   "View King James Version (1611)"
         Top             =   2040
         Width           =   435
      End
      Begin VB.CommandButton cmdYLT 
         Caption         =   "&YLT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         ToolTipText     =   "View Young's Literal Translation (1898)"
         Top             =   2040
         Width           =   435
      End
      Begin VB.CommandButton cmdRSV 
         Caption         =   "&RSV"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1995
         TabIndex        =   28
         ToolTipText     =   "View Revised Standard Version (1971)"
         Top             =   2040
         Width           =   435
      End
      Begin VB.CommandButton cmdMPV 
         Caption         =   "M&PV"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         TabIndex        =   29
         ToolTipText     =   "View your own Personal Verson"
         Top             =   2040
         Width           =   435
      End
      Begin VB.CommandButton cmdAddNote 
         Caption         =   "Add/Edit No&tes"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   32
         ToolTipText     =   "Add/Edit Personal Notes on this verse"
         Top             =   2460
         Width           =   1335
      End
      Begin VB.CommandButton cmdCopyVerse 
         Caption         =   "Copy"
         Height          =   375
         Left            =   60
         TabIndex        =   24
         ToolTipText     =   "Copy verse to the clipboard"
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton cmdCopyNotes 
         Caption         =   "Copy"
         Height          =   375
         Left            =   0
         TabIndex        =   31
         ToolTipText     =   "Copy verse notes to the clipboard"
         Top             =   2460
         Width           =   795
      End
      Begin RichTextLib.RichTextBox rtbVerse 
         Height          =   1095
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1931
         _Version        =   393217
         BackColor       =   -2147483626
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         Appearance      =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmGrkXlate.frx":27670
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbVerseNotes 
         Height          =   735
         Left            =   60
         TabIndex        =   30
         Top             =   1440
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1296
         _Version        =   393217
         BackColor       =   -2147483626
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         Appearance      =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmGrkXlate.frx":27759
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPersonal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Notes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   960
         TabIndex        =   25
         ToolTipText     =   "User personal notes are present (click to scroll up and see them)"
         Top             =   1140
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblTheoNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notice that the actual Greek text sometimes will disagree with these notes."
         Height          =   195
         Left            =   780
         TabIndex        =   52
         Top             =   2580
         Width           =   5235
      End
   End
   Begin VB.PictureBox picGreek 
      BackColor       =   &H80000010&
      FillColor       =   &H80000012&
      Height          =   2775
      Left            =   1860
      ScaleHeight     =   2715
      ScaleWidth      =   4695
      TabIndex        =   35
      Top             =   840
      Width           =   4755
      Begin VB.PictureBox picVbar3 
         Height          =   1875
         Left            =   1920
         MousePointer    =   9  'Size W E
         ScaleHeight     =   1815
         ScaleWidth      =   60
         TabIndex        =   64
         ToolTipText     =   "Drag to resize, double-click to reset"
         Top             =   0
         Width           =   120
      End
      Begin VB.ListBox lstGrkWords 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   2100
         MouseIcon       =   "frmGrkXlate.frx":277D4
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.PictureBox picGreekControl 
         BackColor       =   &H80000000&
         Height          =   735
         Left            =   0
         ScaleHeight     =   675
         ScaleWidth      =   4575
         TabIndex        =   42
         Top             =   1920
         Width           =   4635
         Begin VB.CommandButton cmdHView 
            Height          =   255
            Left            =   3975
            OLEDropMode     =   1  'Manual
            Picture         =   "frmGrkXlate.frx":27926
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Verse viewing history..."
            Top             =   420
            Width           =   315
         End
         Begin VB.CommandButton cmdHBack 
            Height          =   255
            Left            =   3660
            OLEDropMode     =   1  'Manual
            Picture         =   "frmGrkXlate.frx":27EB0
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Previous verse in history"
            Top             =   420
            Width           =   315
         End
         Begin VB.CommandButton cmdHNext 
            Height          =   255
            Left            =   4290
            Picture         =   "frmGrkXlate.frx":2843A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Next verse in history"
            Top             =   420
            Width           =   315
         End
         Begin VB.HScrollBar hsGreek 
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   420
            Width           =   3495
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Height          =   375
            Left            =   0
            TabIndex        =   4
            ToolTipText     =   "Copy Greek text to the clipboard (uses Symbol font)"
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton cmdCopyAll 
            Caption         =   "Copy A&ll"
            Height          =   375
            Left            =   3780
            TabIndex        =   3
            ToolTipText     =   "Copy Greek, Bible Verse, Transliteration (and possible user edit) text to the clipboard (Symbol font used for the Greek text)"
            Top             =   0
            Width           =   795
         End
         Begin VB.Label lblVerseIndex 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   43
            ToolTipText     =   "Click this field to type in verse to view in this chapter, or use the slider below"
            Top             =   0
            Width           =   2895
         End
      End
      Begin RichTextLib.RichTextBox rtbGreek 
         Height          =   1875
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   3307
         _Version        =   393217
         BackColor       =   -2147483626
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   99
         Appearance      =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmGrkXlate.frx":289C4
         MouseIcon       =   "frmGrkXlate.frx":28A46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblGrkWords 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Greek words"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   51
         Top             =   0
         Width           =   2760
      End
   End
   Begin VB.PictureBox picTree 
      Height          =   6210
      Left            =   60
      ScaleHeight     =   6150
      ScaleWidth      =   1725
      TabIndex        =   34
      Top             =   1200
      Width           =   1785
      Begin ComctlLib.TreeView tvBooks 
         Height          =   5835
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   10292
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   33
      Top             =   7725
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Version"
            TextSave        =   "Version"
            Key             =   "version"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   16457
            Text            =   $"frmGrkXlate.frx":28BA8
            TextSave        =   $"frmGrkXlate.frx":28C3A
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGrkXlate.frx":28CCC
            Key             =   "bkOpen"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGrkXlate.frx":28FE6
            Key             =   "bkClosed"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGrkXlate.frx":29300
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblWidth2 
      AutoSize        =   -1  'True
      Caption         =   "lblWidth2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   84
      Top             =   7380
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblPlurality 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "'Quick' plurality checks on word endings are indicated in the Direct Translation with a trailing (s)."
      Height          =   195
      Left            =   60
      TabIndex        =   54
      Top             =   7500
      Width           =   6735
   End
   Begin VB.Label lblwidth 
      AutoSize        =   -1  'True
      Caption         =   "lblWidth"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   45
      Top             =   7080
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileGoto 
         Caption         =   "Go to book chapter:verse"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCreate 
         Caption         =   "Initialize a &Personal version of the Bible..."
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save current chages to master database"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBackup 
         Caption         =   "Save backups of modified files..."
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFileSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSearch 
         Caption         =   "&Search..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBk 
         Caption         =   "Back&grounds"
         Begin VB.Menu mnuBKParch1 
            Caption         =   "Parchment &1"
         End
         Begin VB.Menu mnuBKParch2 
            Caption         =   "Parchment &2"
         End
         Begin VB.Menu mnuBKParch3 
            Caption         =   "Parchment &3"
         End
         Begin VB.Menu mnuBKIce 
            Caption         =   "&Ice"
         End
         Begin VB.Menu mnuBKCloth 
            Caption         =   "Smooth &cloth"
         End
         Begin VB.Menu mnuBKRumpled 
            Caption         =   "&Rumpled cloth"
         End
         Begin VB.Menu mnuBKStucco 
            Caption         =   "&Stucco"
         End
         Begin VB.Menu mnuBKMarble 
            Caption         =   "&Dark marble"
         End
         Begin VB.Menu mnuBKMarbleTx 
            Caption         =   "&Texured marble"
         End
         Begin VB.Menu mnuBKWood 
            Caption         =   "Red &wood"
         End
         Begin VB.Menu mnuBkSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBkCustom 
            Caption         =   "Custom Background..."
         End
      End
      Begin VB.Menu mnuFileSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileWordpad 
         Caption         =   "Launch &Wordpad text editor"
      End
      Begin VB.Menu mnuFileNotepad 
         Caption         =   "Launch &Notepad text editor"
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "Fo&nt"
      Begin VB.Menu mnuFont8 
         Caption         =   "&8 point"
      End
      Begin VB.Menu mnuFont10 
         Caption         =   "1&0 point"
      End
      Begin VB.Menu mnuFont12 
         Caption         =   "1&2 point"
      End
      Begin VB.Menu mnuFont14 
         Caption         =   "1&4 point"
      End
      Begin VB.Menu mnuFont16 
         Caption         =   "1&6 point"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      WindowList      =   -1  'True
      Begin VB.Menu mnuBBLOrgKJV 
         Caption         =   "Explore the original KJV translation strategy..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuBibleTranslateKJV 
         Caption         =   "To&ggle Extract KJV words and translate modernly"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewVine 
         Caption         =   "View Vine &database word list..."
      End
      Begin VB.Menu mnuViewSyn 
         Caption         =   "View &synonym word list..."
      End
      Begin VB.Menu mnuFileViewKJVDict 
         Caption         =   "View &King James Version Dictionary..."
      End
      Begin VB.Menu mnuFileSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileViewChapter 
         Caption         =   "View current chapter in the Definition panel"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileAnalysis 
         Caption         =   "View full verse breakdown (Analysis)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuFileViewTheo 
         Caption         =   "View &Theological Notes for this chapter"
      End
      Begin VB.Menu mnuFileTheoNext 
         Caption         =   "View next verse containing T&heological notes"
      End
      Begin VB.Menu mnuFileTheoPrev 
         Caption         =   "View previous verse containing The&ological notes"
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBBLFindNext 
         Caption         =   "View &next Personal Note"
      End
      Begin VB.Menu mnuBBLFindPrev 
         Caption         =   "View &previous Personal Note"
      End
      Begin VB.Menu mnuBBLViewPNotesChapter 
         Caption         =   "View Personal Notes for only this &chapter"
      End
      Begin VB.Menu mnuBBLViewAllPersonalNotes 
         Caption         =   "&View ALL Personal Notes..."
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBBLFindNextNoGreek 
         Caption         =   "View next verse without &Greek text"
      End
      Begin VB.Menu mnuBBLFindPrevNoGreek 
         Caption         =   "View previous verse without Gr&eek text"
      End
      Begin VB.Menu mnuBibleSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewReset 
         Caption         =   "&Reset frames to default sizing"
      End
   End
   Begin VB.Menu mnuBible 
      Caption         =   "&Bible"
      Begin VB.Menu mnuBibleASV 
         Caption         =   "&American Standard Version (1901)"
      End
      Begin VB.Menu mnuBibleDarby 
         Caption         =   "&Darby's Translation (1884)"
      End
      Begin VB.Menu mnuBibleKJV 
         Caption         =   "&King James Version (1611)"
      End
      Begin VB.Menu mnuBibleMKJV 
         Caption         =   "Mo&dern King James Version (1962)"
      End
      Begin VB.Menu mnuBibleMPV 
         Caption         =   "&My Personal Version"
      End
      Begin VB.Menu mnuBibleRSV 
         Caption         =   "&Revised Standard Version (1971)"
      End
      Begin VB.Menu mnuBibleWebster 
         Caption         =   "Web&ster's Translation (1833)"
      End
      Begin VB.Menu mnuBibleWeb 
         Caption         =   "&World English Bible (1907)"
      End
      Begin VB.Menu mnuBibleYLT 
         Caption         =   "&Young's Literal Translation (1898)"
      End
      Begin VB.Menu mnuBBLSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBBLFindGrkInst 
         Caption         =   "&Find all instances of current Greek word"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuBibleSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBBLCompare 
         Caption         =   "Compare all bibles to this verse"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuBibleSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCreateBible 
         Caption         =   "Write complete Bible file..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFileViewSavedBible 
         Caption         =   "View currently saved &Bible..."
      End
      Begin VB.Menu mnuBBLSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSort 
         Caption         =   "S&ort all synonyms and maintain word mapping"
      End
   End
   Begin VB.Menu mnuFav 
      Caption         =   "F&avorites"
      Begin VB.Menu mnuFavAdd 
         Caption         =   "&Add current verse to favorites"
      End
      Begin VB.Menu mnuFavDel 
         Caption         =   "&Edit Favorites list..."
      End
      Begin VB.Menu mnuFavSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFavList 
         Caption         =   "Favorite"
         Index           =   0
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Windows"
      Begin VB.Menu mnuWinVine 
         Caption         =   "&Vine database word list"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWinStrong 
         Caption         =   "&Synonym word list"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWinKJV 
         Caption         =   "&King James Version dictionary"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWinBible 
         Caption         =   "&Bible viewer"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWinSearch 
         Caption         =   "&Search Bibles"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHLPUsing 
         Caption         =   "Using the New Covenant Translator..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHlpCC 
         Caption         =   "Crash course in Greek..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuHLPQR 
         Caption         =   "Quick References for plurality..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuHLPKeyboard 
         Caption         =   "&Keyboard navigation help..."
      End
      Begin VB.Menu mnuHlpPlurality 
         Caption         =   "About &Plurality checking..."
      End
      Begin VB.Menu mnuHlpMascFem 
         Caption         =   "About &Gender checking..."
      End
      Begin VB.Menu mnuHLPViewReadme 
         Caption         =   "View the &README file..."
      End
      Begin VB.Menu mnuHLPSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHLPViewDemo 
         Caption         =   "View an auto-running &demonstration of this application..."
      End
      Begin VB.Menu mnuHLPTipoftheDay 
         Caption         =   "&Tip of the day..."
      End
      Begin VB.Menu mnuHlpSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHLPVisitSponsor 
         Caption         =   "&View the companion book at the AuthorHouse website..."
      End
      Begin VB.Menu mnuHLPSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHlpAbout 
         Caption         =   "&About this program..."
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Begin VB.Menu mnuPopupVine 
         Caption         =   "Search for &Vine Reference for the selection"
      End
      Begin VB.Menu mnuPopUpBible 
         Caption         =   "Search for &Bible reference for the selection..."
      End
      Begin VB.Menu mnuPopUpCopy 
         Caption         =   "&Copy selected text to the clipboard"
      End
   End
   Begin VB.Menu mnuPopUp2 
      Caption         =   "mnuPopUp2"
      Begin VB.Menu mnuPopUpCopy2 
         Caption         =   "&Copy selected text to the clipboard"
      End
   End
   Begin VB.Menu mnuPopUp3 
      Caption         =   "mnuPopUp3"
      Begin VB.Menu mnuPopUpCopy3 
         Caption         =   "&Copy selected text to the clipboard"
      End
   End
End
Attribute VB_Name = "frmGrkXlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vBarDown1 As Boolean          'true when vBar1 has mouse down over it
Private vBarDown2 As Boolean
Private vBarDown3 As Boolean
Private vBarDown4 As Boolean
Private hBarDown1 As Boolean

Private WinState As Integer           'keep track of window state

Private orgLstGrkWordsWidth As Long
Private orgLstWordsWidth As Long
Private orgPicGreekWidth As Long
Private orgPicGreekHeight As Long
Private orgTvWidth As Long

Private GreekWidth As Long            'top-left panel width and height
Private GreekHeight As Long

Private NotFirstTimeIn As Boolean
Private ManualResize As Boolean       'when forcing a resize due to moving bars
Private SetGoto As Boolean

Private LastSch As String             'last text search text

Public Favorz As String              'Favoring translation'

Private Const NoGreekText As String = "No Greek text for "

Private MyToolTips As clsToolTip

Private BookDropHandler As clscboFullDrop   'handle fulldrop on book list
Private ChapDropHandler As clscboFullDrop   'handle fulldrop on chapter list
Private VerseDropHandler As clscboFullDrop  'handle fulldrop on verse list
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_CONTROL = &H11
Private Const VK_SHIFT = &H10
Private Const VeryLight As Long = &HF0F0F0

'*******************************************************************************
' Subroutine Name   : cboBk_Click
' Purpose           : build the chapter combo list
'*******************************************************************************
Private Sub cboBk_Click()
  Dim Idx As Long, I As Long
  Dim Ary() As String
  Me.cmdGo.Enabled = Bk <> Me.cboBk.ListIndex + 1 Or Chp <> Me.cboChp.ListIndex + 1 Or Vrs <> Me.cboVrs.ListIndex + 1
  If SetGoto Then Exit Sub
  Ary = Split(Books(Me.cboBk.ListIndex + 1), ",")
  I = CLng(Ary(4))              'get # of chapters in the book
'
' build chapter list
'
  With Me.cboChp
    .Clear
    For Idx = 1 To I
      .AddItem CStr(Idx) 'add chapter list
    Next Idx
    .ToolTipText = "Select chapter 1 - " & CStr(I)
   .ListIndex = 0               'select chapter 1
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cboChp_Click
' Purpose           : build the verse combo list
'*******************************************************************************
Private Sub cboChp_Click()
  Me.cmdGo.Enabled = Bk <> Me.cboBk.ListIndex + 1 Or Chp <> Me.cboChp.ListIndex + 1 Or Vrs <> Me.cboVrs.ListIndex + 1
  If Not SetGoto Then SetVerseCount
End Sub

'*******************************************************************************
' Subroutine Name   : cboVrs_Click
' Purpose           : Verse selected
'*******************************************************************************
Private Sub cboVrs_Click()
  Me.cmdGo.Enabled = Bk <> Me.cboBk.ListIndex + 1 Or Chp <> Me.cboChp.ListIndex + 1 Or Vrs <> Me.cboVrs.ListIndex + 1
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAddNote_Click
' Purpose           : Add a personal note to a verse
'*******************************************************************************
Private Sub cmdAddNote_Click()
  frmPersonalNotes.Show vbModal, Me
  If bCancel Then Exit Sub            'user cancelled
  UpdateVerse                         'else update verse information
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAnalysis_Click
' Purpose           : Save a verse analys of a verse
'*******************************************************************************
Private Sub cmdAnalysis_Click()
  Dim S As String, Ary() As String, T As String, TT As String
  Dim Idx As Long, I As Long, J As Long, SS As Long, SL As Long
  
  S = Format$(Bk, "00") & Format$(Chp, "00") & Format$(Vrs, "00")
  I = FindExactMatch(Me.lstGrk, S)            'find the greek text
  S = Grk(I)                                  'grab it
  If Len(S) < 8 Then                          'only a header?
   MessageBox Me, "No Greek text exists for " & Ttl, vbOKOnly Or vbExclamation, "Verse Has No Legal Content"
    Exit Sub
  End If
  
'  Me.cmdAnalysis.Enabled = False              'disable button until data updates
'  Me.mnuFileAnalysis.Enabled = False
  
  InitMerge                                   'initialize the merge text data
  AddMergeCrLf2 Ttl, "Arial", 14, True        'add the book title with 2 CRLF
  AddMergeCrlf Mid$(S, 8), "Symbol", FntSize, True    'display the Greek text
  AddMergeSep                                 'add a line separator
  S = Me.rtbVerse.Text                        'get the English Bible verse
  I = InStr(1, S, ":")                        'find the separator
  AddMerge Left$(S, I), , FntSize, True, True 'stuff it, bold-talic
  I = InStr(1, S, vbCrLf)
  AddMergeCrlf Mid$(S, I + 2), , FntSize      'rest of the data is normal
  AddMergeSep
  AddMergeCrlf "Current Direct Translation:", , FntSize
  AddMergeCrlf Mid$(Me.rtbTranslate.Text, DTHeaderOffset), , FntSize, True
  AddMergeSep
  
  SS = Len(Me.rtbMerge.Text)                  'save the insert point for a later hanging indent
'
' now process each Greek word
'
  With Me.lstGrkWords
    For Idx = 0 To .ListCount - 1
      I = CLng(BBLLine(Idx + 1))
      Ary = Split(DefRef(CLng(BBLLine(Idx + 1))), vbTab)
      Ary = Split(DefRef(I), vbTab)
      AddMerge .List(Idx) & "    (" & Ary(1) & ")    ", "Symbol", FntSize, True
      AddMerge Ary(2), , FntSize, True, True
      S = "    (" & Ary(3) & ")    Strong's Reference # " & Ary(5) & vbCrLf
      If Len(Ary(4)) <> 0 Then S = S & Ary(4) & vbCrLf
      AddMergeCrlf S, , FntSize
      '
      ' display the description contents
      '
      S = Ary(6)
      Call CheckUsage(.List(Idx))
      If Len(Favorz) <> 0 Then
        S = S & "\\This particular usage may favor a sense of:  " & Favorz & "."
      End If
      I = InStr(1, S, "\")                          'conver "\" to vbCrLf
      Do While I
        S = Left$(S, I - 1) & vbCrLf & Mid$(S, I + 1)
        I = InStr(I + 2, S, "\")
      Loop
      AddMergeCrLf2 S, , FntSize
      AddMergeCrlf "Current Synonyms:", , FntSize   'prepart for synonyms
      Ary = Split(WordRef(CLng(Ary(5))), vbTab)     'grab list of words from word list
      Ary = Split(Ary(2), ",")                      'grab words
      I = CLng(MiniMap(Idx + 1))                    'get current selected words
      If I = -1 Then I = 0
      Ary(I) = "[" & Ary(I) & "]"                   'enbrace that one for show
      AddMergeCrlf Join(Ary, ", "), , FntSize, , True 'sent it to merge data
      '
      ' if we just processed the last entry, then add a separator
      '
      If Idx = .ListCount - 1 Then
        AddMergeCrlf Underline
      End If
      AddMergeSep
    Next Idx
    '
    ' add verse notes
    '
    S = Format$(Bk, "00") & Format$(Chp, "00") & Format$(Vrs, "00")
    T = "No Theological Verse Notes for " & Ttl           'init no theological verse notes
    I = FindExactMatch(Me.lstVNotes, S)       'find a match
    If I <> -1 Then                           'found something...
      J = I                                   'save last-found index
      T = "Theological Verse Update Notes for " & Ttl 'init header
      Do While I <> -1
        T = T & vbCrLf & vbCrLf & Mid$(VNotes(I), 8)  'add a note
        I = FindExactMatch(Me.lstVNotes, S, I)  'find another
        If J >= I Then Exit Do                 'ignore if index matches last
      Loop
    End If
    
    TT = Mid$(MyNotes(VrsIdx), 8)
    If Len(TT) <> 0 Then
      Me.lblPersonal.Visible = True
      J = InStr(1, TT, "\")
      Do While J <> 0
        TT = Left$(TT, J - 1) & vbCrLf & Mid$(TT, J + 1)
        J = InStr(J + 2, TT, "\")
      Loop
      AddMergeCrLf2 T, , FntSize
      AddMergeCrlf MyPersonalNotes, , FntSize, True, True, , PNotesColor
      T = TT
    Else
      Me.lblPersonal.Visible = False
    End If
      
    J = InStr(1, T, "{")
    Do While J <> 0
      AddMerge Left$(T, J - 1), , FntSize, , , , PNotesColor
      I = InStr(J + 1, T, "}")
      If I = 0 Then Exit Do                 'no match
      TT = Mid$(T, J + 1, I - J - 1)
      AddMerge TT, "Symbol", FntSize, True, , , PNotesColor
      T = Mid$(T, I + 1)
      J = InStr(1, T, "{")
    Loop
    If Len(T) <> 0 Then
      AddMerge T, , FntSize, , , , PNotesColor
    End If
  End With
  
  SL = Len(Me.rtbMerge.Text) - SS                   'get the leangth of the data after headers
  With Me.lblwidth
    .FontSize = FntSize
    .Caption = String$(7, 32)
    SetIndent SS, SL, .Width                        'set hanging indent
  End With
  
  With Me.rtbNotes
    LockWindowUpdate .hwnd
    .BackColor = VeryLight
    .Text = vbNullString                            'ensure we are at the top of the textbox
    .TextRTF = Me.rtbMerge.TextRTF                  'stuff data
  End With
  LockWindowUpdate 0
  InitMerge                                         'then clear it
  Me.cmdVine.Enabled = True                         'enable Vine data button if visisble
  Me.lblVineRef.Visible = False
  Me.cmdCopyDef.Enabled = True                      'ensure the user can copy this text
  lblNoteInfo.Visible = True
  Me.cmdBack.Visible = True
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBack_Click (RESET button)
' Purpose           : Redisplay definition for Greek word
'*******************************************************************************
Private Sub cmdBack_Click()
  Me.rtbNotes.ToolTipText = vbNullString
  Call lstGrkWords_Click
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopyDef_Click
' Purpose           : Copy definitions list to clickboard
'*******************************************************************************
Private Sub cmdCopyDef_Click()
  Clipboard.Clear
  Clipboard.SetText Me.rtbNotes.Text
  Clipboard.SetText Me.rtbNotes.TextRTF, vbCFRTF
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopyDT_Click
' Purpose           : Copy Direct Translation list to clickboard
'*******************************************************************************
Private Sub cmdCopyDT_Click()
  Clipboard.Clear
  Clipboard.SetText Me.rtbTranslate.Text              'save base text version
  Clipboard.SetText Me.rtbTranslate.TextRTF, vbCFRTF  'save rich text version
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopyNotes_Click
' Purpose           : Copy notes list to clickboard
'*******************************************************************************
Private Sub cmdCopyNotes_Click()
  Clipboard.Clear
  Clipboard.SetText Me.rtbVerseNotes.Text
  Clipboard.SetText Me.rtbVerseNotes.TextRTF, vbCFRTF
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopyVerse_Click
' Purpose           : Copy verse notes list to clickboard
'*******************************************************************************
Private Sub cmdCopyVerse_Click()
  Clipboard.Clear
  Clipboard.SetText Me.rtbVerse.Text
  With Me.rtbMerge                        'copy to merge to ensure forecolor is black
    .TextRTF = Me.rtbVerse.TextRTF        '(can be white on personal bible)
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelColor = vbBlack
    Clipboard.SetText .TextRTF, vbCFRTF
    .Text = vbNullString
  End With
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdFindInText_Click
' Purpose           : Find text in the Notes panel
'*******************************************************************************
Private Sub cmdFindInText_Click()
  Dim Sch As String, Text As String
  Dim Idx As Long
  
  If Me.rtbNotes.SelLength <> 0 Then
    Text = Me.rtbNotes.SelText
  Else
    Text = LastSch
  End If
  Sch = InputMsgBox(Me, "Enter word or phrase to find in the Definition Panel:", "Search For Text", Text)
  If Len(Sch) = 0 Then Exit Sub
  LastSch = Sch
  With Me.rtbNotes
    Text = .Text
    Idx = InStr(1, Text, Sch, vbTextCompare)
    If Idx <> 0 Then
      .SelStart = Len(Text)
      .SelStart = Idx - 1
      .SelLength = Len(Sch)
      Me.cmdFindNext.Enabled = True
    Else
      Me.cmdFindNext.Enabled = False
      MessageBox Me, "Search text not found: " & Sch, vbOKOnly Or vbExclamation, "Text Not Found"
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdFindNext_Click
' Purpose           : Find next match of same search text
'*******************************************************************************
Private Sub cmdFindNext_Click()
  Dim Idx As Long
  Dim Text As String
  With Me.rtbNotes
    Text = .Text
    Idx = .SelStart + .SelLength + 1
    Idx = InStr(Idx, LCase$(Text), LCase$(LastSch))
    If Idx <> 0 Then
      .SelStart = Len(Text)
      .SelStart = Idx - 1
      .SelLength = Len(LastSch)
      Me.cmdFindNext.Enabled = True
    Else
      Me.cmdFindNext.Enabled = False
      MessageBox Me, "Search text not found: " & LastSch, vbOKOnly Or vbExclamation, "Text Not Found"
    End If
  End With
End Sub

Private Sub cmdGo_Click()
  Bk = Me.cboBk.ListIndex + 1
  Chp = Me.cboChp.ListIndex + 1
  Vrs = Me.cboVrs.ListIndex + 1
  ChpCnt = Me.cboChp.ListCount
  VrsCnt = Me.cboVrs.ListCount
  SetGoto = True
  Call UpdateVerse    'all ok, so display the user selection
  Me.cmdGo.Enabled = False
  SetGoto = False
End Sub

'*******************************************************************************
' Subroutine Name   : cmdHBack_Click
' Purpose           : Go to previous point to history
'*******************************************************************************
Private Sub cmdHBack_Click()
  Dim S As String, Ary() As String
  Dim Idx As Long
  
  ChgSCroll = True                        'prevent redundancy
  HistIdx = HistIdx - 1                   'back off history index
  If HistIdx < 1 Then HistIdx = 1
  Me.cmdHBack.Enabled = HistIdx > 1       'enable/disable buttons as needed
  If Me.cmdHBack.Enabled Then Me.cmdHBack.ToolTipText = "Previous verse in history: " & GetVerseData(HistIdx - 1)
  Me.cmdHNext.Enabled = HistIdx < colHistory.Count
  If Me.cmdHNext.Enabled Then Me.cmdHNext.ToolTipText = "Next verse in history: " & GetVerseData(HistIdx + 1)
  S = colHist(HistIdx)                    'get contents
  Bk = CLng(Left$(S, 2))                  'get book, chapter and verse
  Chp = CLng(Mid$(S, 3, 2))
  Vrs = CLng(Right$(S, 2))
  Ary = Split(Books(Bk), ",")
  ChpCnt = CLng(Ary(4))                   'get the chapter count
  Call GetVerseCount                      'get the verse count
  HistUpdt = True
  Call UpdateVerse                        'display the verse
  HistUpdt = False
  ChgSCroll = False                       'reset protections flags
End Sub

'*******************************************************************************
' Subroutine Name   : cmdHNext_Click
' Purpose           : Go to next point in history
'*******************************************************************************
Private Sub cmdHNext_Click()
  Dim S As String, Ary() As String
  Dim Idx As Long
  
  ChgSCroll = True                        'prevent redundancy
  HistIdx = HistIdx + 1                   'bump history index
  If HistIdx > colHist.Count Then HistIdx = colHist.Count
  Me.cmdHBack.Enabled = HistIdx > 1       'enable/disable buttons as needed
  If Me.cmdHBack.Enabled Then Me.cmdHBack.ToolTipText = "Previous verse in history: " & GetVerseData(HistIdx - 1)
  Me.cmdHNext.Enabled = HistIdx < colHist.Count
  If Me.cmdHNext.Enabled Then Me.cmdHNext.ToolTipText = "Next verse in history: " & GetVerseData(HistIdx + 1)
  S = colHist(HistIdx)                    'get contents
  Bk = CLng(Left$(S, 2))                  'get book, chapter and verse
  Chp = CLng(Mid$(S, 3, 2))
  Vrs = CLng(Right$(S, 2))
  Ary = Split(Books(Bk), ",")
  ChpCnt = CLng(Ary(4))                   'get the chapter count
  Call GetVerseCount                      'get the verse count
  HistUpdt = True
  Call UpdateVerse                        'display the verse
  HistUpdt = False
  ChgSCroll = False                       'reset protections flags
End Sub

'*******************************************************************************
' Subroutine Name   : cmdHView_Click
' Purpose           : View session history list
'*******************************************************************************
Private Sub cmdHView_Click()
  frmViewSessionHistory.Show vbModal, Me
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : mnuBBLCompare_Click
' Purpose           : Compare Bibles against each other
'*******************************************************************************
Private Sub mnuBBLCompare_Click()
  Dim SB As String
  Dim V As Long, I As Long
  
  Screen.MousePointer = vbHourglass                         'show that we are busy
  Me.Enabled = False
  DoEvents
  InitMerge                                                 'init merging
  AddMergeCrlf Ttl, "Arial", FntSize + 2, True, , vbCenter  'add the title for the verse
'
' display Greek text
'
  AddMergeCrlf "Greek:", , FntSize - 2, True
  SB = Mid$(Grk(VrsIdx), 8)
  V = CLng(Mid$(Grk(VrsIdx), 5, 2))
  If Len(Grk(VrsIdx)) < 8 Then V = -V                       'if a vapor verse
  If Len(SB) = 0 Then
    AddVerse NoVerseTextAvail & vbCrLf, -V, , FntSize - 2
  Else
    AddVerse SB & vbCrLf, V, "Symbol", FntSize - 2
  End If
'
' display direct translation
'
  AddMergeCrlf " ", , 6
  AddMergeCrlf "Direct Translation:", , FntSize - 2, True
  SB = Mid$(Me.rtbTranslate.Text, DTHeaderOffset)
  If Len(SB) = 0 Then
    AddVerse NoVerseTextAvail & vbCrLf, -V, , FntSize - 2
  Else
    AddVerse SB & vbCrLf, V, , FntSize - 2
  End If
  
  If ASVAvail Then GrabVerse 6  '"ASV"
  If DBYAvail Then GrabVerse 7  '"DBY"
  If KJVAvail Then GrabVerse 0  '"KJV"
  If MKJVAvail Then GrabVerse 4 '"MKJV"
  If MPVAvail Then GrabVerse 3  '"MPV"
  If RSVAvail Then GrabVerse 2  '"RSV"
  If WBSAvail Then GrabVerse 8  '"WBS"
  If WEBAvail Then GrabVerse 5  '"WEB"
  If YLTAvail Then GrabVerse 1  '"YLT"
  
  LockWindowUpdate Me.rtbNotes.hwnd           'avoid flashing
  Me.rtbNotes.Text = vbNullString             'ensure at top of window
  Me.rtbNotes.TextRTF = Me.rtbMerge.TextRTF   'stuff new data
  Me.rtbNotes.BackColor = VeryLight           'set background
  LockWindowUpdate 0                          'refresh display
  InitMerge                                   'erase merge data
  
  Me.cmdCopyDef.Enabled = True                'enable things that may be disabled...
  Me.mnuFileAnalysis.Enabled = True
  Me.cmdAnalysis.Enabled = True
  Me.lblVineRef.Enabled = True
  lblNoteInfo.Visible = False
  Me.cmdBack.Visible = True
  Screen.MousePointer = vbDefault             'no longer busy
  Me.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : GrabVerse
' Purpose           : Support routine. DIsplay various verses of the verse
'*******************************************************************************
Private Sub GrabVerse(sBible As Long)
  Dim Ary() As String, SB As String, sV As String, S As String
  Dim V As Long
  
  Select Case sBible
    Case 1
      SB = "YLT"
      sV = "Young's Literal Translation"
    Case 2
      SB = "RSV"
      sV = "Revised Standard Version"
    Case 3
      SB = "MPV"
      sV = "My Personal Version"
    Case 4
      SB = "MKJV"
      sV = "Modern King James Version"
    Case 5
      SB = "WEB"
      sV = "World English Bible"
    Case 6
      SB = "ASV"
      sV = "American Standard Version"
    Case 7
      SB = "DBY"
      sV = "Darby's Translation"
    Case 8
      SB = "WBS"
      sV = "Webster's Translation"
    Case Else
      SB = "KJV"
      sV = "King James Version"
  End Select

  AddMergeCrlf " ", , 6                                   'give a little space
  AddMergeCrlf "(" & sV & ")", , FntSize - 2, True, True  'show which version of bible
  If sBible = BblVersion Then
    S = Bible(VrsIdx)                                     'grab verse data
  Else
    If SB = "MPV" Then
      Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\" & SB & ".txt", ForReading, False)
    Else
      Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\" & SB & ".txt", ForReading, False)
    End If
    Ary = Split(ts.ReadAll, vbCrLf)                       'derive bible data
    ts.Close
    S = Ary(VrsIdx)                                       'grab verse data
    Erase Ary                                             'release resources
  End If
  
  V = InStr(1, S, "{")
  Do While V <> 0
    Mid$(S, V, 1) = "["
    V = InStr(1, S, "{")
  Loop
  V = InStr(1, S, "}")
  Do While V <> 0
    Mid$(S, V, 1) = "]"
    V = InStr(1, S, "}")
  Loop
  
  V = CLng(Mid$(S, 5, 2))                                 'grab verse data
  If Len(Grk(VrsIdx)) < 8 Then V = -V                     'vaport verse
  
  If SB = "MPV" And V < 0 Then
    SB = vbNullString
  Else
    SB = Mid$(S, 8)                                       'grab verse text
  End If
  If Len(SB) = 0 Then
    AddVerse NoVerseTextAvail & vbCrLf, -V, , FntSize - 2 'display verse data
  Else
    AddVerse SB & vbCrLf, V, , FntSize - 2                'display verse data
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuBBLFindGrkInst_Click
' Purpose           : Find all instances of current Greek word in bible
'*******************************************************************************
Private Sub mnuBBLFindGrkInst_Click()
  Screen.MousePointer = vbHourglass
  DoEvents
  If Me.cmdBack.Visible Then Me.cmdBack.Value = True
  frmFndGrkInst.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuBBLFindNext_Click
' Purpose           : Find Next personal Note
'*******************************************************************************
Private Sub mnuBBLFindNext_Click()
  Dim Idx As Long
  Dim S As String
  
  For Idx = VrsIdx + 1 To UBound(MyNotes) - 1
    S = MyNotes(Idx)
    If Len(S) > 7 Then Exit For
  Next Idx
  If Len(S) < 8 Then
    For Idx = 0 To VrsIdx - 1
      S = MyNotes(Idx)
      If Len(S) > 7 Then Exit For
    Next Idx
  End If
  If Len(S) > 7 Then
    ForceVerse S            'update display for the verse indicated by the text header
    Call lblPersonal_Click
    Me.lblVineRef.Visible = False
    Me.cmdCopyDef.Enabled = True
    Me.cmdAnalysis.Enabled = True
    Me.mnuFileAnalysis.Enabled = True
  Else
    MessageBox Me, "No (other) Personal Notes Found", vbOKOnly Or vbInformation, "No Personal Notes"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ForceVerse
' Purpose           : Force Display of non-sequential Book, Chapter, and Verse
'*******************************************************************************
Private Sub ForceVerse(BkChVs As String)
  Dim Ary() As String
  
  Bk = CLng(Left$(BkChVs, 2))     'grab the Book
  Chp = CLng(Mid$(BkChVs, 3, 2))  'grab the Chapter
  Vrs = CLng(Mid$(BkChVs, 5, 2))  'Grab the verse
  
  Ary = Split(Books(Bk), ",")     'get the chapter count for the book
  ChpCnt = CLng(Ary(4))
  Call GetVerseCount              'get the verse count for the chapter
  UpdateVerse                     'update everything
End Sub


'*******************************************************************************
' Subroutine Name   : mnuBBLFindPrev_Click
' Purpose           : Find Previous personal Note
'*******************************************************************************
Private Sub mnuBBLFindPrev_Click()
  Dim Idx As Long
  Dim S As String
  
  For Idx = VrsIdx - 1 To 0 Step -1
    S = MyNotes(Idx)
    If Len(S) > 7 Then Exit For
  Next Idx
  If Len(S) < 8 Then
    For Idx = UBound(MyNotes) - 1 To VrsIdx + 1 Step -1
      S = MyNotes(Idx)
      If Len(S) > 7 Then Exit For
    Next Idx
  End If
  If Len(S) > 7 Then
    ForceVerse S            'update display for the verse indicated by the text header
    Call lblPersonal_Click
    Me.lblVineRef.Visible = False
    Me.cmdCopyDef.Enabled = True
    Me.cmdAnalysis.Enabled = True
    Me.mnuFileAnalysis.Enabled = True
  Else
    MessageBox Me, "No (other) Personal Notes Found", vbOKOnly Or vbInformation, "No Personal Notes"
  End If
End Sub
'
' User selection of bible versions
'
Private Sub cmdKJV_Click()
  If BblVersion <> 0 Then UpdateBible 0
End Sub

Private Sub cmdYLT_Click()
  If BblVersion <> 1 Then UpdateBible 1
End Sub

Private Sub cmdRSV_Click()
  If BblVersion <> 2 Then UpdateBible 2
End Sub

Private Sub cmdMPV_Click()
  If BblVersion <> 3 Then UpdateBible 3
End Sub

Private Sub cmdMKJV_Click()
  If BblVersion <> 4 Then UpdateBible 4
End Sub

Private Sub cmdWEB_Click()
  If BblVersion <> 5 Then UpdateBible 5
End Sub

Private Sub cmdASV_Click()
  If BblVersion <> 6 Then UpdateBible 6
End Sub

Private Sub cmdDBY_Click()
  If BblVersion <> 7 Then UpdateBible 7
End Sub

Private Sub cmdWBS_Click()
  If BblVersion <> 8 Then UpdateBible 8
End Sub

'*******************************************************************************
' Subroutine Name   : cmdVine_Click
' Purpose           : Display Vine reference data, if button visible
'*******************************************************************************
Private Sub cmdVine_Click()
  DisplayVine VineIndex, LCase$(Me.lstWords.List(Me.lstWords.ListIndex))
  Me.cmdVine.Enabled = False                        'disable button
End Sub

'*******************************************************************************
' Subroutine Name   : DisplayVine
' Purpose           : Display the Vine Bible Dictionary
'*******************************************************************************
Public Sub DisplayVine(ByVal Index As Long, Txt As String)
  Dim S As String, Ary() As String, T As String
  Dim I As Long, J As Long, K As Long, oIndex As Long
  
  oIndex = -1
  S = vbNullString
  Do
    If oIndex = -1 Then
      oIndex = Index
      Me.rtbNotes.Text = vbNullString
      InitMerge                                     'init merge RTB
    Else
      I = FindExactMatch(Me.lstVine, Txt, oIndex)   'search for next match
      If oIndex >= I Then Exit Do                   'no next match, so done
      oIndex = I
      AddMergeSep
    End If
    S = Me.lstVineNum.List(oIndex)                  'get data for reference
    S = Vine(CLng(S))
    Ary = Split(S, vbTab)                           'break up data
    
    AddMerge Ary(1), , FntSize, True, True
    AddMergeCrlf "    (Vine Reference #" & Ary(0) & ")", , FntSize, , True
    AddMergeSep
    S = Ary(2)                                      'grab main data
    I = InStr(1, S, "[TT]")
    
    If I = 1 Then
      S = Chr$(34) & Txt & Chr$(34) & Mid$(S, I + 4)
    ElseIf I <> 0 Then
      S = Left$(S, I - 1) & Txt & Mid$(S, I + 4)
    End If
'
' add newline codes
'
    I = InStr(1, S, "\")
    Do While I <> 0
      S = Left$(S, I - 1) & vbCrLf & Mid$(S, I + 1)
      I = InStr(I + 2, S, "\")
    Loop
'
' now locate Greek data
'
    I = InStr(1, S, "#")
    Do While I <> 0
      I = InStrRev(S, ";", I - 1)                     'Range "; Strong #"
      J = InStrRev(S, ":", I - 1)                     'find beginning of data
      If I = 0 Or J = 0 Then Exit Do                  'if the line is simply a reference
      MergeCheck Left$(S, J)                          'process previous "normal" data
      T = Trim$(Mid$(S, J + 1, I - J - 1))            'grab Greek word
      AddMerge " " & T, "Symbol", FntSize, True       'display in Greek
      If Right$(T, 1) = "V" Then T = Left$(T, Len(T) - 1) & "s"
      AddMerge " [" & T & "]", , FntSize, , True
      S = Mid$(S, I)                                  'strip Greek word and previous
      J = InStr(1, S, ")")                            'find end of data
      AddMerge Left$(S, J), , FntSize                 'add the sub-data normal
      S = Mid$(S, J + 1)                              'strip processed data
      I = InStr(1, S, "#")                            'find next entry
    Loop
    If Len(S) <> 0 Then
      If Right$(S, 2) = vbCrLf Then
        MergeCheck S                                  'normal for any remainder
      Else
        MergeCheck S & vbCrLf                         'normal for any remainder
      End If
    End If
  Loop
  UnderlineBbl Me.rtbMerge
  SetIndent                                         'set indenting
  LockWindowUpdate Me.rtbNotes.hwnd
  Me.rtbNotes.BackColor = clBlue
  Me.rtbNotes.TextRTF = Me.rtbMerge.TextRTF         'stuff merged data to display
  LockWindowUpdate 0
  Me.rtbMerge.Text = vbNullString                   'flush pot
  Me.cmdAnalysis.Enabled = True                     'enable buttons
  Me.mnuFileAnalysis.Enabled = True
  Me.cmdCopyDef.Enabled = True
  Me.lblVineRef.Visible = True
  Me.cmdBack.Visible = True
End Sub

'*******************************************************************************
' Subroutine Name   : MergeCheck
' Purpose           : Convert any tagged Greek text to the Symbol typeface
'*******************************************************************************
Private Sub MergeCheck(Text As String)
  Dim S As String, T As String
  Dim I As Long, J As Long
  
  S = Text
  I = InStr(1, S, "{")                              'find start of any Greek data
  Do While I <> 0                                   'while we find it
    AddMerge Left$(S, I - 1), , FntSize + BumpFactor 'normal data
    J = InStr(I + 1, S, "}")                        'find end of Greek data
    If J = 0 Then Exit Do                           'no match
    T = Mid$(S, I + 1, J - I - 1)                   'grab Greek text
    AddMerge T, "Symbol", FntSize + BumpFactor, True 'stuff Greek data
    S = Mid$(S, J + 1)                              'strip Greek data and previous
    I = InStr(1, S, "{")                            'find start of any more Greek data
  Loop
  If Len(S) <> 0 Then                               'anything left?
    AddMerge S, , FntSize + BumpFactor              'yes, process as normal data
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdFind_Click
' Purpose           : Find a word or words in the vine index
'*******************************************************************************
Private Sub cmdFind_Click()
  Dim S As String, T As String, Ary() As String, TT As String, Find As String
  Dim I As Long, J As Long, Chk As Long, Idx As Long
  
  With Me.lstWords
    If .ListIndex >= 0 Then S = .List(.ListIndex)
  End With
  Find = Trim$(InputMsgBox(Me, "Enter an English word or Vine Reference # to find:", _
               "Find Word", S, _
               "Check the full Vine database text if the word is not found in the word list.", True))
  If Len(Find) = 0 Then Exit Sub
  Chk = CLng(GetSetting(App.Title, "Settings", "CheckVineAll", CStr(vbChecked)))
'
' if numeric, get the actual databased index and the search word
'
  If IsNumeric(Find) Then
    I = FindExactMatch(Me.lstVineNum, Find)
    If I <> -1 Then Find = Me.lstVine.List(I)    'get search word
  Else
    I = FindExactMatch(Me.lstVine, LCase$(Find)) 'find word
  End If
'
' if not found (I=-1) and the user wants to check the full database...
'
  If I = -1 Then
    Screen.MousePointer = vbHourglass         'show that we are busy
    Me.Enabled = False
    DoEvents
    
    If Chk = vbChecked Then
      LastSch = Find
      TT = " " & LCase$(Find) & " "           'test string
      For J = 1 To Len(TT)
        Select Case Mid$(TT, J, 1)
          Case "a" To "z", "-", "'", " ", "0" To "9"
          Case Else
            Mid$(TT, J, 1) = " "             'strip non-allowed characters
        End Select
      Next J
      
      For Idx = 1 To UBound(Vine) - 1         'scan each database line
        Ary = Split(Vine(Idx), vbTab)         'grab data
        S = LCase$(Ary(2)) & " "              'now process contents
        For J = 1 To Len(S)
          Select Case Mid$(S, J, 1)
            Case "a" To "z", "-", "'", " ", "0" To "9"
            Case Else
              Mid$(S, J, 1) = " "             'strip non-allowed characters
          End Select
        Next J
        J = InStr(1, S, "  ")                 'collapse large gaps
        Do While J <> 0
          S = Left$(S, J) & Mid$(S, J + 2)
          J = InStr(J, S, "  ")
        Loop
        If InStr(1, S, TT) Then               'now check for match
          I = FindExactMatch(Me.lstVineNum, CStr(Idx)) 'found it, get the ref index
          If I <> -1 Then                     'safety net (should never be -1)
            Find = Me.lstVine.List(I)         'get search word
            Exit For                          'exit scan
          End If
        End If
      Next Idx
    End If
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    If I <> -1 Then
      DisplayVine I, UCase$(Find)
      Me.cmdFindNext.Enabled = True
      Me.cmdFindNext.Value = True
      Me.cmdVine.Enabled = True
      Me.cmdAnalysis.Enabled = True
      Me.mnuFileAnalysis.Enabled = True
      Me.cmdCopyDef.Enabled = True
      Exit Sub
    End If
  End If
'
' if Chk not (or still) not set, indicate it weas not found
'
  If I = -1 Then
    MessageBox Me, "The word """ & Find & """ was not found in the Vine Database.", vbOKOnly Or vbExclamation, "Word not found"
    Exit Sub
  End If
'
' found the data, so display it
'
  DisplayVine I, UCase$(Find)
  Me.cmdVine.Enabled = True
  Me.cmdAnalysis.Enabled = True
  Me.mnuFileAnalysis.Enabled = True
  Me.cmdCopyDef.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Initialize
' Purpose           : Set up for XP buttons
'*******************************************************************************
Private Sub Form_Initialize()
  If App.PrevInstance Then Exit Sub
  FormInitialize
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : IShow splash screen, init databases
'*******************************************************************************
Private Sub Form_Load()
  Dim Nd As Node, cNd As Node
  Dim Idx As Long, I As Long
  Dim S As String, Ary() As String, T As String, SD As String, DD As String
  Dim Drv As Drive
  Dim Fil As File
  Dim Fnd As Boolean, Chg As Boolean, TryTip As Boolean
'
' see if a previous instance exists
'
  If App.PrevInstance Then
    ActivatePrevInstance
    Unload Me
    Exit Sub
  End If
  IsLoading = True
'
' hide certain popup menus
'
  Me.mnuPopUp.Visible = False
  Me.mnuPopUp2.Visible = False
  Me.mnuPopUp3.Visible = False
  Me.mnuWin.Visible = False
'
' clear text boxes
'
Me.rtbGreek.Text = vbNullString
Me.rtbNotes.Text = vbNullString
Me.rtbTranslate.Text = vbNullString
Me.rtbUser.Text = vbNullString
Me.rtbVerse.Text = vbNullString
Me.rtbVerseNotes.Text = vbNullString
'
' store original list sizes
'
  orgLstGrkWordsWidth = Me.lstGrkWords.Width
  orgLstWordsWidth = Me.lstWords.Width
  orgTvWidth = Me.picTree.Width
  WinState = -1
'
' get a copy of the local data path
'
  AppPath = App.Path
'
' set version
'
  Me.StatusBar1.Panels("version").Text = "Version " & GetAppVersion()
'
' init tiling images
'
  For I = 0 To Me.picTile.Count - 1
    InitTileFormBackground Me.picTile(I) 'init tiling images
  Next I
'
' set up toolbar
'
  AssignToolbarButtonImages Me.Toolbar1, Me.ImageList2
'
' Set some colors
'
  clBlue = RGB(223, 239, 255)
  cdGray = RGB(128, 128, 128)
  cMissing = RGB(255, 0, 255)
  cBurgandy = RGB(128, 0, 0)
  PNotesColor = CLng(GetSetting(App.Title, "Settings", "PNotesColor", CStr(vbBlue)))
'
' set color patterns
'
  Select Case GetScreenColorCount(Me.hdc)
    Case Is < 3
      custDark = CLng(GetSetting(App.Title, "Settings", "custDark", CStr(RGB(255, 255, 255))))
      custMedium = CLng(GetSetting(App.Title, "Settings", "custMedium", CStr(RGB(208, 208, 208))))
      custLight = CLng(GetSetting(App.Title, "Settings", "custLight", CStr(RGB(224, 224, 224))))
      CustVLight = CLng(GetSetting(App.Title, "Settings", "custVLight", CStr(RGB(255, 255, 255))))
      S = GetSetting(App.Title, "Settings", "CustPicture", AddSlash(App.Path) & "Resources\BkWhite.jpg")
      Idx = CLng(GetSetting(App.Title, "Settings", "Background", CStr(bkCustom)))
    Case Else
      custDark = CLng(GetSetting(App.Title, "Settings", "custDark", CStr(RGB(230, 198, 134))))
      custMedium = CLng(GetSetting(App.Title, "Settings", "custMedium", CStr(RGB(250, 225, 172))))
      custLight = CLng(GetSetting(App.Title, "Settings", "custLight", CStr(RGB(255, 235, 186))))
      CustVLight = CLng(GetSetting(App.Title, "Settings", "custVLight", CStr(RGB(255, 242, 186))))
      S = GetSetting(App.Title, "Settings", "CustPicture", vbNullString)
      Idx = CLng(GetSetting(App.Title, "Settings", "Background", CStr(bkParch1)))
  End Select
  If Len(S) <> 0 Then
    On Error Resume Next
    Me.picTile(bkCustom).Picture = LoadPicture(S)
    If Err.Number <> 0 Then
      Me.picTile(bkCustom).Picture = Me.picTile(bkParch1).Picture
    End If
    On Error GoTo 0
  Else
    Me.picTile(bkCustom).Picture = Me.picTile(bkParch1).Picture
  End If
  
  SetPattern Idx
'
' set custom tootip for treeview
'
  Set MyToolTips = New clsToolTip
  With MyToolTips
    .Create Me               'create object
    .MaxTipWidth = 1440 * 2  'width max = 2 inches
    .DelayTime(ttDelayShow) = 20 * 1000 'set to 20 seconds
    .SetFont , 8
    .AddTool Me.tvBooks
    .ToolText(Me.tvBooks) = vbNullString
    .AddTool Me.lstGrkWords
    .ToolText(Me.lstGrkWords) = vbNullString
    .AddTool Me.lstWords
    .ToolText(Me.lstWords) = vbNullString
    .AddTool Me.rtbGreek
    .ToolText(Me.rtbGreek) = vbNullString
    .AddTool Me.hsGreek     'trick to allow scrollbar to have a tooltip
    .ToolText(Me.hsGreek) = "Click or slide tab to move between verses"
  End With
'
' get flag for KJV word translation
'
    TranslateKJV = CBool(GetSetting(App.Title, "Settings", "TranslateKJV", "1"))
    Me.mnuBibleTranslateKJV.Checked = TranslateKJV
'
' init file I/O workhorse
'
  Set Fso = New FileSystemObject
'
' set up storage location. This nat be due to running the app from a CD, where
' its local storage location is read-only.  This code allows the user to run
' the application from the CD by copying requisite files from the CD to a
' writable hard-disc location.
'
  Set Drv = Fso.GetDrive(Fso.GetDriveName(AppPath))
  If (GetAttr(Drv.Path) And vbReadOnly) Then
    S = GetSetting(App.Title, "Settings", "RomStore", vbNullString)
    If Not Fso.FolderExists(S) Then
      S = vbNullString
    Else
      If Not Fso.FolderExists(AddSlash(S) & "DB") Then S = vbNullString
    End If
    If S = vbNullString Then
      Do
        If MessageBox(Me, "You appear to be running this program from a read-only location." & vbCrLf & _
                  "Would you like to define a recordable storage location for data?", _
                  vbYesNo Or vbQuestion, "Running from Read-Only Location") = vbNo Then
          Unload Me
          Exit Sub
        End If
        S = DirBrowser(Me.hwnd, ViewDirsOnly, "Define a Data Storage Location")
        If Len(S) = 0 Then
          Unload Me
          Exit Sub
        End If
        Set Drv = Fso.GetDrive(Fso.GetDriveName(S))
        If Not (GetAttr(Drv.Path) And vbReadOnly) Then Exit Do
      Loop
      T = AddSlash(S) & "DB"
      If Not Fso.FolderExists(T) Then
        On Error Resume Next
        Fso.CreateFolder T
        If Err.Number <> 0 Then
         MessageBox Me, "Cannot created path: " & T, vbOKOnly Or vbExclamation, "Aborting"
          Unload Me
          Exit Sub
        End If
        On Error GoTo 0
      End If
      AppPath = S
      SaveSetting App.Title, "Settings", "RomStore", AppPath
    Else
      AppPath = S
    End If
  End If
'
' reconstruct application as needed
'
' if Running from CD, or running from a folder without a DB folder, it is assumed that
' all file contents of the DB folder (and the UsingNCXlate_files folder) are stored
' in the current folder that NCXlate.exe is running from.
'
  If Not (GetAttr(AppPath) And vbReadOnly) Then
    SD = AddSlash(App.Path)
    DD = AddSlash(AppPath) & "DB"
    If Fso.FolderExists(DD) Then
      Fnd = CBool(Len(Dir$(DD & "\*.htm")))
    Else
      Fnd = False
    End If
    Screen.MousePointer = vbHourglass
    DoEvents
    If Not Fnd Then
      If Not Fso.FolderExists(DD) Then Fso.CreateFolder DD
      DD = DD & "\"
      T = Dir$(SD & "*.htm*")
      Do While Len(T)
        Fso.CopyFile SD & T, DD & T, True
        Set Fil = Fso.GetFile(DD & T)
        If Fil.Attributes And ReadOnly Then Fil.Attributes = Fil.Attributes - ReadOnly
        T = Dir$()
      Loop
      T = Dir$(SD & "*.pdf")
      Do While Len(T)
        Fso.CopyFile SD & T, DD & T, True
        Set Fil = Fso.GetFile(DD & T)
        If Fil.Attributes And ReadOnly Then Fil.Attributes = Fil.Attributes - ReadOnly
        T = Dir$()
      Loop
      T = Dir$(SD & "*.txt")
      Do While Len(T)
        If LCase$(T) <> "readme.txt" Then
          Fso.CopyFile SD & T, DD & T, True
        Set Fil = Fso.GetFile(DD & T)
        If Fil.Attributes And ReadOnly Then Fil.Attributes = Fil.Attributes - ReadOnly
        End If
        T = Dir$()
      Loop
      T = Dir$(SD & "*.vbs")
      Do While Len(T)
        Fso.CopyFile SD & T, DD & T, True
        Set Fil = Fso.GetFile(DD & T)
        If Fil.Attributes And ReadOnly Then Fil.Attributes = Fil.Attributes - ReadOnly
        T = Dir$()
      Loop
      T = Dir$(SD & "*.jar")
      Do While Len(T)
        Fso.CopyFile SD & T, DD & T, True
        Set Fil = Fso.GetFile(DD & T)
        If Fil.Attributes And ReadOnly Then Fil.Attributes = Fil.Attributes - ReadOnly
        T = Dir$()
      Loop
      T = Dir$(SD & "*.viewlet")
      Do While Len(T)
        Fso.CopyFile SD & T, DD & T, True
        Set Fil = Fso.GetFile(DD & T)
        If Fil.Attributes And ReadOnly Then Fil.Attributes = Fil.Attributes - ReadOnly
        T = Dir$()
      Loop
    End If
    
    DD = AddSlash(DD) & "UsingNCX_Files"
    If Not Fso.FolderExists(DD) Then
      Fso.CreateFolder DD
      DD = DD & "\"
      T = Dir$(SD & "*.png")
      Do While Len(T)
        Fso.CopyFile SD & T, DD & T, True
        Set Fil = Fso.GetFile(DD & T)
        If Fil.Attributes And ReadOnly Then Fil.Attributes = Fil.Attributes - ReadOnly
        T = Dir$()
      Loop
      T = Dir$(SD & "*.jpg")
      Do While Len(T)
        Fso.CopyFile SD & T, DD & T, True
        Set Fil = Fso.GetFile(DD & T)
        If Fil.Attributes And ReadOnly Then Fil.Attributes = Fil.Attributes - ReadOnly
        T = Dir$()
      Loop
      T = Dir$(SD & "*.xml")
      Do While Len(T)
        If LCase$(T) <> "readme.txt" Then
          Fso.CopyFile SD & T, DD & T, True
        Set Fil = Fso.GetFile(DD & T)
        If Fil.Attributes And ReadOnly Then Fil.Attributes = Fil.Attributes - ReadOnly
        End If
        T = Dir$()
      Loop
    End If
    On Error GoTo 0
    Screen.MousePointer = vbDefault
  End If
'
' continue now with normal startup operations
'
  Set colFavs = New Collection
  I = CLng(GetSetting(App.Title, "Settings", "FavCnt", "0"))
  Me.mnuFavSep.Visible = CBool(I)
  Me.mnuFavList(0).Visible = False
  For Idx = 0 To I - 1
    S = GetSetting(App.Title, "Settings", "Fav" & CStr(Idx + 1))
    If Idx > 0 Then Load Me.mnuFavList(Idx)
    Me.mnuFavList(Idx).Caption = S
    Me.mnuFavList(Idx).Visible = True
    colFavs.Add S, S
  Next Idx
  Me.mnuFavDel.Enabled = I > 0
'
' set up history collection
'
  Set colHist = New Collection
  HistIdx = 0
  Set colSrch = New Collection
'
' get View History
'
  Set colHistory = New Collection
  If Fso.FileExists(AddSlash(AppPath) & "DB\History.txt") Then
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\History.txt", ForReading, False)
    Ary = Split(ts.ReadAll, vbCrLf)
    ts.Close
    For Idx = 0 To UBound(Ary)
      S = Ary(Idx)
      If Len(S) <> 0 Then
        colHistory.Add S
      End If
    Next Idx
  End If
'
' disable menu items if they are not available
'
  If Not Fso.FileExists(AddSlash(App.Path) & "DB\UsingNCX.htm") Then Me.mnuHLPUsing.Enabled = False
  If Not Fso.FileExists(AddSlash(App.Path) & "DB\CrashCourse.htm") Then Me.mnuHlpCC.Enabled = False
  If Not Fso.FileExists(AddSlash(App.Path) & "DB\QRef.htm") Then Me.mnuHLPQR.Enabled = False
  If Not Fso.FileExists(AddSlash(App.Path) & "DB\NCXlateDemo_Viewlet.html") Then Me.mnuHLPViewDemo.Enabled = False
  If Not Fso.FileExists(AddSlash(App.Path) & "Readme.txt") Then Me.mnuHLPViewReadme.Enabled = False
  
  S = GetWindowsDir()
  I = InStrRev(S, "\")
  S = Left$(S, I) & "Program Files\"
  WordPadPath = S & "Accessories"
  If Len(Dir$(WordPadPath, vbDirectory)) Then
    WordPadPath = WordPadPath & "\WordPad.exe"
    If Len(Dir$(WordPadPath)) = 0 Then WordPadPath = vbNullString
  Else
    WordPadPath = vbNullString
  End If
  If Len(WordPadPath) = 0 Then
    WordPadPath = S & "Windows NT\Accessories"
    If Len(Dir$(WordPadPath, vbDirectory)) Then
      WordPadPath = WordPadPath & "\WordPad.exe"
      If Len(Dir$(WordPadPath)) = 0 Then WordPadPath = vbNullString
    End If
  End If
  If Len(WordPadPath) = 0 Then
    Me.mnuFileWordpad.Enabled = False
    Me.Toolbar1.Buttons("wordpad").Enabled = False
    WordPadPath = vbNullString
  End If
  
  S = GetWindowsDir() & "\"
  NotePadPath = S & "Notepad.exe"
  If Len(Dir$(NotePadPath)) = 0 Then
    NotePadPath = S & "System\Notepad.exe"
  End If
  If Len(Dir$(WordPadPath)) = 0 Then
    WordPadPath = vbNullString
    Me.mnuFileNotepad.Enabled = False
    Me.Toolbar1.Buttons("notepad").Enabled = False
    NotePadPath = vbNullString
  End If
  
  
  If Not Fso.FileExists(AddSlash(App.Path) & "DB\KJVDict.txt") Then
    Me.mnuFileViewKJVDict.Enabled = False
    Me.Toolbar1.Buttons("kjvdict").Enabled = False
  End If
  Me.mnuBBLFindNext.Enabled = False
  Me.mnuBBLFindPrev.Enabled = False
  Me.mnuBBLFindNextNoGreek.Enabled = False
  Me.mnuBBLFindPrevNoGreek.Enabled = False
  Me.mnuBBLViewAllPersonalNotes.Enabled = False
  Me.mnuBBLViewPNotesChapter.Enabled = False
  Me.mnuFileViewTheo.Enabled = False
  Me.mnuBBLOrgKJV.Enabled = False
  Me.mnuBibleTranslateKJV.Enabled = False
  With Me.Toolbar1
    .Buttons("prevnote").Enabled = False
    .Buttons("nextnote").Enabled = False
    .Buttons("prevtheo").Enabled = False
    .Buttons("nexttheo").Enabled = False
    .Buttons("prevgreek").Enabled = False
    .Buttons("nextgreek").Enabled = False
  End With
  
  Me.cmdEdit.Enabled = False        'initially disable the command buttons
  Me.cmdAdd.Enabled = False
  Me.cmdDel.Enabled = False
  Me.cmdUpdateMPV.Enabled = False
  Me.cmdCpyXlt.Enabled = False
  Me.cmdCopympv.Enabled = False
  Me.rtbUser.Visible = False
  Me.mnuFileCreateBible.Enabled = False
  Me.mnuFileViewChapter.Enabled = False
  Me.cmdFindInText.Enabled = False
  Me.cmdFindNext.Enabled = False
  Me.cmdBack.Visible = False
  Me.mnuFileTheoNext.Enabled = False
  Me.mnuFileTheoPrev.Enabled = False
  Me.mnuBBLCompare.Enabled = False
'
' show the splash screen
'
  Me.mnuHlpAbout.Enabled = False
  With frmSplash
    SplashIsOn = True               'indicate using form as a splash screen
    .Show vbModeless, Me            'bring it up
    Do While Not .Timer1.Enabled    'wait until the timer is activated after fade-in
      DoEvents                      'allow stuff to happen
    Loop
  End With
'
' get font point sizing information and set it
'
  Changefont CLng(GetSetting(App.Title, "settings", "FontSize", "1"))
'
' find out if a personal version is available
'
  PersonalVersion = Fso.FileExists(AddSlash(AppPath) & "DB\MPV.TXT")
'
' read the list of 27 bible looks
'
  On Error Resume Next
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\Books.txt", ForReading, False)
  If Err.Number <> 0 Then
    MessageBox Me, "Cannot find DB\Books.txt database.", vbOKOnly Or vbCritical, "Aborting"
    Unload frmSplash
    Unload Me
    Exit Sub
  End If
  On Error GoTo 0
  Books = Split(ts.ReadAll, vbCrLf)     'store content in an arroa
  ts.Close
'
' get last selected book/chapter/verse
'
  Bk = CLng(GetSetting(App.Title, "Settings", "Book", "0"))
  Chp = CLng(GetSetting(App.Title, "Settings", "Chapter", "0"))
  Vrs = CLng(GetSetting(App.Title, "Settings", "Verse", "0"))
'
' build the treeview lists
'
  Me.tvBooks.ImageList = Me.ImageList1  'set image list for the treeview
  For Idx = 1 To 27
    Ary = Split(Books(Idx), ",")
    If Idx = 1 Then
      Set RootNode = Me.tvBooks.Nodes.Add(, , Ary(1), Ary(3), 2, 2)
      Set Nd = RootNode
    Else
      Set Nd = Me.tvBooks.Nodes.Add(RootNode.Index, tvwLast, Ary(1), Ary(3), 2, 2)
    End If
    Nd.Tag = Ary(3) 'set book Title
    
    If Idx = Bk Then Set BkNode = Nd  'set book node if we are pre-loading data
    For I = 1 To CLng(Ary(4))
      Set cNd = Me.tvBooks.Nodes.Add(Nd.Index, tvwChild, Ary(1) & Format$(I, "00"), CStr(I), 3)
      cNd.Tag = Ary(3) & ", Chapter " & CStr(I)
      If Bk = Idx Then
        If Chp = I Then Set ChpNode = cNd 'set chapter node if it is being preloaded
      End If
    Next I
  Next Idx
'
' ensure book and chapter shown in treeview if we are pre-loading it
'
  If Bk <> 0 Then
    BkNode.Expanded = True
    ChpNode.Selected = True
    ChpNode.EnsureVisible
    BkNode.EnsureVisible
  End If
'
' read the word definition and comment database
'
  On Error Resume Next
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\GreekDefRef.txt", ForReading, False)
  If Err.Number <> 0 Then
    MessageBox Me, "Cannot find DB\GreekDefRef.txt database.", vbOKOnly Or vbCritical, "Aborting"
    Unload frmSplash
    Unload Me
    Exit Sub
  End If
  On Error GoTo 0
  DefRef = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' read the word list database
'
  On Error Resume Next
  Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\GreekWordRef.txt", ForReading, False)
  If Err.Number <> 0 Then
    On Error Resume Next
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\GreekWordRefNew.txt", ForReading, False)
  End If
  If Err.Number <> 0 Then
    MessageBox Me, "Cannot find DB\GreekWordRef.txt database.", vbOKOnly Or vbCritical, "Aborting"
    Unload frmSplash
    Unload Me
    Exit Sub
  End If
  On Error GoTo 0
  WordRef = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' grab the indexes for "favored" or last-used words for a particular synonym group
'
  ReDim BBLWIdx(UBound(WordRef)) As Long
  For Idx = 1 To UBound(BBLWIdx)
    S = WordRef(Idx)
    If Len(S) <> 0 Then
      Ary = Split(S, vbTab)
      BBLWIdx(Idx) = CLng(Ary(1))
      '
      ' if recreating a fresh map image...
      '
      If MakeVirgin Then
        BBLWIdx(Idx) = 0                 'used to reset to a 'virgin' reference table
        Ary(1) = "0"
        WordRef(Idx) = Join(Ary, vbTab)
      End If
    End If
  Next Idx
'
' read the word reference database to maintain personally-selected word usage integrity
'
  On Error Resume Next
  Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\GreekBBL.txt", ForReading, False)
  If Err.Number <> 0 Then
    MessageBox Me, "Cannot find DB\GreekBBL.txt database.", vbOKOnly Or vbCritical, "Aborting"
    Unload frmSplash
    Unload Me
    Exit Sub
  End If
  On Error GoTo 0
  GrkBBL = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' read the Vine word reference database
'
  On Error Resume Next
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\VineRef.txt", ForReading, False)
  If Err.Number <> 0 Then
    MessageBox Me, "Cannot find DB\VineRef.txt database.", vbOKOnly Or vbCritical, "Aborting"
    Unload frmSplash
    Unload Me
    Exit Sub
  End If
  On Error GoTo 0
  Vine = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' build a list of references to the Vine words
'
  For Idx = 1 To UBound(Vine)
    S = Vine(Idx)
    If Len(S) <> 0 Then
      Ary = Split(S, vbTab)
      S = LCase$(Ary(1))
      Ary = Split(S, ",")
      For I = 0 To UBound(Ary)
        Me.lstVine.AddItem Trim$(Ary(I))
        Me.lstVineNum.AddItem CStr(Idx)
      Next I
    End If
  Next Idx
'
' grab list of user-selected synonyms for each word in each verse
'
  If Not Fso.FileExists(AddSlash(AppPath) & "DB\WordMap.txt") Then
'
' create default list if one does not yet exist
'
    ReDim WordMap(UBound(GrkBBL)) As String
    For Idx = 0 To UBound(GrkBBL)
      S = GrkBBL(Idx)
      If Len(S) Then
        Ary = Split(S, " ")
        For I = 1 To UBound(Ary)
          Ary(I) = "-1"  'flag using default from BBLWIdx until user selects one
        Next I
        WordMap(Idx) = Join(Ary, " ")
      End If
    Next Idx
  Else
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\WordMap.txt", ForReading, False)
    WordMap = Split(ts.ReadAll, vbCrLf)
    ts.Close
'
' upgrade WordMap.txt data from older version, if needed
'
    If Left$(WordMap(5912), 6) <> "081314" Then     'ensure blank verse added if not there
      For I = UBound(WordMap) To 5913 Step -1
        WordMap(I) = WordMap(I - 1)
      Next I
      WordMap(5912) = "081314"
      ReDim Preserve WordMap(UBound(WordMap) + 1)
    End If
    
    If Left$(WordMap(7528), 6) = "250115" Then      'remove verse data incorporated in previous verse
      For I = 7528 To UBound(WordMap) - 1
        WordMap(I) = WordMap(I + 1)
      Next I
      ReDim Preserve WordMap(UBound(WordMap) - 1)
    End If
    
    If Left$(WordMap(7764), 6) = "271218" Then      'remove verse data incorporated in previous verse
      For I = 7764 To UBound(WordMap) - 1
        WordMap(I) = WordMap(I + 1)
      Next I
      ReDim Preserve WordMap(UBound(WordMap) - 1)
    End If
  End If
'
' read the Paragraph map. This contains a list of flags indicating if each verse
' (save the first verse of a chapter) begins a new paragraph
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\PMap.txt", ForReading, False)
  ParMap = ts.ReadAll
  ts.Close
'
' Read the KJV word counts
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\KJVCounts.txt", ForReading, False)
  KJVCount = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' read the KJV translation verse index
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\KJVVerseIndex.txt", ForReading, False)
  KJVidxAry = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' read the KJV translation word list
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\KJVVerseWords.txt", ForReading, False)
  KJVwrdAry = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' check all supported versions of the bible
'
  KJVAvail = Fso.FileExists(AddSlash(AppPath) & "DB\KJV.txt")
  Me.cmdKJV.Enabled = KJVAvail
  Me.mnuBibleKJV.Enabled = KJVAvail
  MKJVAvail = Fso.FileExists(AddSlash(AppPath) & "DB\MKJV.txt")
  Me.cmdMKJV.Enabled = MKJVAvail
  Me.mnuBibleMKJV.Enabled = MKJVAvail
  YLTAvail = Fso.FileExists(AddSlash(AppPath) & "DB\YLT.txt")
  Me.cmdYLT.Enabled = YLTAvail
  Me.mnuBibleYLT.Enabled = YLTAvail
  RSVAvail = Fso.FileExists(AddSlash(AppPath) & "DB\RSV.txt")
  Me.cmdRSV.Enabled = RSVAvail
  Me.mnuBibleRSV.Enabled = RSVAvail
  WEBAvail = Fso.FileExists(AddSlash(AppPath) & "DB\WEB.txt")
  Me.cmdWEB.Enabled = WEBAvail
  Me.mnuBibleWeb.Enabled = WEBAvail
  ASVAvail = Fso.FileExists(AddSlash(AppPath) & "DB\ASV.txt")
  Me.cmdASV.Enabled = ASVAvail
  Me.mnuBibleASV.Enabled = ASVAvail
  DBYAvail = Fso.FileExists(AddSlash(AppPath) & "DB\DBY.txt")
  Me.cmdDBY.Enabled = DBYAvail
  Me.mnuBibleDarby.Enabled = DBYAvail
  WBSAvail = Fso.FileExists(AddSlash(AppPath) & "DB\WBS.txt")
  Me.mnuBibleWebster.Enabled = WBSAvail
  Me.cmdWBS.Enabled = WBSAvail
  
  MPVAvail = PersonalVersion
  Me.cmdMPV.Enabled = MPVAvail
  Me.mnuBibleMPV.Enabled = MPVAvail
  Me.lblEditPersonal.Visible = False
'
' get the user's currently selected bible version
'
  BblVersion = CLng(GetSetting(App.Title, "Settings", "Bible", "0"))
  Select Case BblVersion
    Case 1
      S = "YLT"
      VersionText = "Young's Literal Translation"
    Case 2
      S = "RSV"
      VersionText = "Revised Standard Version"
    Case UserPVer
      S = "MPV"
      VersionText = "My Personal Version"
    Case 4
      S = "MKJV"
      VersionText = "Modern King James Version"
    Case 5
      S = "WEB"
      VersionText = "World English Bible"
    Case 6
      S = "ASV"
      VersionText = "American Standard Version"
    Case 7
      S = "DBY"
      VersionText = "Darby's Translation"
    Case 8
      S = "WBS"
      VersionText = "Webster's Translation"
    Case Else
      S = "KJV"
      VersionText = "King James Version"
  End Select
'
' if we cannot find selected bible, default to KJV
'
  If S = "MPV" Then
    Fnd = Fso.FileExists(AddSlash(AppPath) & "DB\" & S & ".txt")
  Else
    Fnd = Fso.FileExists(AddSlash(App.Path) & "DB\" & S & ".txt")
  End If
  If Not Fnd Then
    If S = "MPV" Then PersonalVersion = False
    S = "KJV"
    VersionText = "King James Version"
    BblVersion = 0
    If Not Fso.FileExists(AddSlash(App.Path) & "DB\KJV.txt") Then
      MessageBox Me, "Cannot find Bible database.", vbOKOnly Or vbCritical, "Aborting"
      Unload frmSplash
      Unload Me
      Exit Sub
    End If
  End If
  
  Me.mnuBibleMPV.Enabled = PersonalVersion
  Select Case BblVersion
    Case 1
      Me.mnuBibleYLT.Checked = True
      Me.cmdYLT.Enabled = False
    Case 2
      Me.mnuBibleRSV.Checked = True
      Me.cmdRSV.Enabled = False
    Case UserPVer
      Me.mnuBibleMPV.Checked = True
      Me.cmdMPV.Enabled = False
      Me.lblEditPersonal.Visible = True
    Case 4
      Me.mnuBibleMKJV.Checked = True
      Me.cmdMKJV.Enabled = False
    Case 5
      Me.mnuBibleWeb.Checked = True
      Me.cmdWEB.Enabled = False
    Case 6
      Me.mnuBibleASV.Checked = True
      Me.cmdASV.Enabled = False
    Case 7
      Me.mnuBibleDarby.Checked = True
      Me.cmdDBY.Enabled = False
    Case 8
      Me.mnuBibleWebster.Checked = True
      Me.cmdWBS.Enabled = False
    Case Else
      Me.mnuBibleKJV.Checked = True
      Me.cmdKJV.Enabled = False
  End Select
'
' load selected bible verse database
'
  If S = "MPV" Then
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\" & S & ".txt", ForReading, False)
  Else
    Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\" & S & ".txt", ForReading, False)
  End If
  Bible = Split(ts.ReadAll, vbCrLf)
  ts.Close
  PVDirty = False
'
' load Greek verse database
'
  On Error Resume Next
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\Greek.txt", ForReading, False)
  If Err.Number <> 0 Then
    MessageBox Me, "Cannot find Greek Bible database.", vbOKOnly Or vbCritical, "Aborting"
    Unload frmSplash
    Unload Me
    Exit Sub
  End If
  On Error GoTo 0
  Grk = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' save reference list for verses
'
  For Idx = 0 To UBound(Grk) - 1
    Me.lstGrk.AddItem Left$(Grk(Idx), 6)
  Next Idx
'
' get Verse notes list
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\VNotes.txt", ForReading, False)
  VNotes = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' get Greek Word Count List
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\GrkWrdCnt.txt", ForReading, False)
  GrkWrdCnt = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' check for personal notes
'
  If Not Fso.FileExists(AddSlash(AppPath) & "DB\MyNotes.txt") Then
    ReDim MyNotes(UBound(Grk)) As String
    For Idx = 0 To UBound(Grk)
      MyNotes(Idx) = Left$(Grk(Idx), 6)
    Next Idx
    MyNotesDirty = True           'file must be updated to disc
    HavePersonalNotes = False     'ACTUAL note entries do not exist
  Else
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\MyNotes.txt", ForReading, False)
    MyNotes = Split(ts.ReadAll, vbCrLf)
    ts.Close
    For Idx = 0 To UBound(MyNotes)
      If Len(MyNotes(Idx)) > 7 Then
        HavePersonalNotes = True
        Exit For
      End If
    Next Idx
  End If
'
' save reference list for verses
'
  For Idx = 0 To UBound(VNotes)
    Me.lstVNotes.AddItem Left$(VNotes(Idx), 6)
  Next Idx
'
' get the form location
'
  Me.Top = CLng(GetSetting(App.Title, "Settings", "FormTop", "0"))
  Me.Left = CLng(GetSetting(App.Title, "Settings", "FormLeft", "0"))
  Me.Width = CLng(GetSetting(App.Title, "Settings", "FormWidth", CStr(Me.Width)))
  Me.Height = CLng(GetSetting(App.Title, "Settings", "FormHeight", CStr(Me.Height)))
  I = CLng(GetSetting(App.Title, "Settings", "FormState", "-1"))
  If I = -1 Then
    Me.WindowState = vbMaximized
  Else
    Me.WindowState = I
  End If
'
' build the book combo list
'
  With Me.picBCV
    .BorderStyle = 0
    .BackColor = cMedium
    Me.cmdGo.Height = .Height
   .Width = Me.cmdGo.Left + Me.cmdGo.Width
   .Left = Me.Toolbar1.Buttons("go").Left
  End With
  
  With Me.cboBk
    .Clear
    For Idx = 1 To 27
      Ary = Split(Books(Idx), ",")
      Me.cboBk.AddItem Ary(3)
    Next Idx
    If Bk = 0 Then
      .ListIndex = 0
    Else
      .ListIndex = Bk - 1
    End If
  End With
  Me.cmdGo.Enabled = Bk = 0
'
' build the verse combo list
'
  Set BookDropHandler = New clscboFullDrop
  Set ChapDropHandler = New clscboFullDrop
  Set VerseDropHandler = New clscboFullDrop
'''*** Comment out following 3 lines if you are debugging this form code,
'''*** as otherwise the WndProc handler will hang the VB IDE if you try to STOP
'''*** (not step through; this is OK) this code with this code active
  BookDropHandler.hwnd = Me.cboBk.hwnd
  ChapDropHandler.hwnd = Me.cboChp.hwnd
  VerseDropHandler.hwnd = Me.cboVrs.hwnd
'
' set up display
'
  Me.Show
  DoEvents
'
' if Books/verse/chapter indicated (not first time in) then load the selection
'
  TryTip = Bk <> 0
  If Not TryTip Then
    TryTip = False
    Me.mnuFav.Enabled = False
    Me.cmdAnalysis.Enabled = False
    Me.mnuFileAnalysis.Enabled = False
    If Me.mnuHLPViewDemo.Enabled Then
      Call mnuHLPViewDemo_Click
    ElseIf Me.mnuHLPUsing.Enabled Then
      Call mnuHLPUsing_Click
    End If
    Bk = 1    'default to Matthew 1:1
    Chp = 1
    Vrs = 1
  End If
  I = Vrs
  Vrs = 1
  HistUpdt = True
  ChgSCroll = True
  Call ShowVerse
  ChgSCroll = False
  HistUpdt = False
  Call GetVerseCount
  'S = "Verse " & CStr(Vrs) & " of " & CStr(VrsCnt)
  Vrs = I
  Ary = Split(Books(Bk), ",")
  ChpCnt = CLng(Ary(4))
  Call UpdateVerse
  Me.mnuFileExit.Caption = "E&xit" & vbTab & "Alt-F4"
'
' get the backup path
'
  BackupPath = GetSetting(App.Title, "Settings", "BackupPath", vbNullString)
  If Len(BackupPath) <> 0 Then
    If Not Fso.FolderExists(BackupPath) Then BackupPath = vbNullString
  End If
'
' if we should save a backup upon program start
'
  If CBool(GetSetting(App.Title, "Settings", "AutoSave", "0")) Then
'
' see if we should auto-save backup files
'
    On Error Resume Next
    If Len(BackupPath) <> 0 Then
'
' ALWAYS save word reference table (constantly updated)
'
      Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "GreekWordRef.txt", ForWriting, True)
      ts.Write Join(WordRef, vbCrLf)
      ts.Close
'
' and word map (user-selections for syninyms of words)
'
      Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "WordMap.txt", ForWriting, True)
      ts.Write Join(WordMap, vbCrLf)
      ts.Close
'
' and Greek word reference
'
      Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "GreekBBL.txt", ForWriting, True)
      ts.Write Join(GrkBBL, vbCrLf)
      ts.Close
'
' save personal verse notes
'
  Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "DB\MyNotes.txt", ForWriting, True)
  ts.Write Join(MyNotes, vbCrLf)
  ts.Close
'
' save user personal bible
'
      If PersonalVersion Then
        If BblVersion = UserPVer Then
          Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "MPV.txt", ForWriting, True)
          ts.Write Join(Bible, vbCrLf)
          ts.Close
        Else
          Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\MPV.txt", ForReading, False)
          Ary = Split(ts.ReadAll, vbCrLf)
          ts.Close
          Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "MPV.txt", ForWriting, True)
          ts.Write Join(Ary, vbCrLf)
          ts.Close
        End If
      End If
    End If
  End If
'
' set positions of bars
'
  I = CLng(GetSetting(App.Title, "Settings", "Hbar1", "0"))
  If I <> 0 Then
    If GreekHeight <> I Then
      Chg = True
      GreekHeight = I
    End If
    I = CLng(GetSetting(App.Title, "Settings", "Vbar1", "0"))
    If Me.picTree.Width <> I Then
      Chg = True
      Me.picVbar1.Left = I + Me.picTree.Left
    End If
    I = CLng(GetSetting(App.Title, "Settings", "Vbar2", "0"))
    If GreekWidth <> I Then
      Chg = True
      GreekWidth = I
    End If
    I = CLng(GetSetting(App.Title, "Settings", "Vbar3", "0"))
    If Me.lstGrkWords.Width <> I Then
      Chg = True
      Me.lstGrkWords.Width = I
    End If
    I = CLng(GetSetting(App.Title, "Settings", "Vbar4", "0"))
    If Me.lstWords.Width <> I Then
      Chg = True
      Me.lstWords.Width = I
    End If
    If Chg Then
      Call ForceResize    'force a repaint if something has changed
    End If
  End If
'
' get Save Bible Path
'
  SaveBible = GetSetting(App.Title, "Settings", "SaveBible", AddSlash(AppPath) & "MyBible.rtf")
  If Not Fso.FileExists(SaveBible) Then SaveBible = vbNullString
  Me.mnuFileViewSavedBible.Enabled = SaveBible <> vbNullString
'
' check for enabling the timer backup
'
  AutoTime = CLng(GetSetting(App.Title, "Settings", "TimeSet", "0"))
  AutoTimeUpd = AutoTime
  Me.tmrAutoBackup.Enabled = CBool(GetSetting(App.Title, "Settings", "AutoTimer", "0"))
  AutoDirty = False
'
' copy the view history to the standard history list
'
  With colHistory
    colHist.Remove 1                   'remove current verse from list
    For Idx = 1 To .Count
      colHist.Add .Item(Idx)
    Next Idx
    HistIdx = .Count
    Me.cmdHBack.Enabled = HistIdx > 1  'enable/disable buttons as needed
    If Me.cmdHBack.Enabled Then Me.cmdHBack.ToolTipText = "Previous verse in history: " & GetVerseData(HistIdx - 1)
    Me.cmdHNext.Enabled = False
  End With
'
' try displaying the tip of the day
'
  If TryTip Then                  'no demo or help up?
    Do While SplashIsOn           'if splash screen is still up
      DoEvents                    'pause
    Loop
    On Error Resume Next
    frmTip.Show vbModal, Me       'show tip of the day, if enabled
  End If
  IsLoading = False
End Sub

'*******************************************************************************
' Subroutine Name   : EnsureTVSet
' Purpose           : Ensure the Treeview is set to the proper Book and chapter
'*******************************************************************************
Private Sub EnsureTVSet()
  Dim bNd As Node, cNd As Node
  Dim S As String, Ary As String
  Dim Idx As Long
  
  Set bNd = RootNode
  With Me.tvBooks
    If Bk > 0 Then
      For Idx = 2 To Bk
        Set bNd = bNd.Next
      Next Idx
      Set cNd = bNd.Child
      If Chp > 1 Then
        For Idx = 2 To Chp
          Set cNd = cNd.Next
        Next Idx
      End If
      
      If Not bNd Is BkNode Then
        If Not BkNode Is Nothing Then BkNode.Selected = False
        Set BkNode = bNd
      End If
      
      If Not cNd Is ChpNode Then
        If Not ChpNode Is Nothing Then ChpNode.Selected = False
        Set ChpNode = cNd
      End If
      
      If Not BkNode.Expanded Then BkNode.Expanded = True
      If Not ChpNode.Selected Then ChpNode.Selected = True
      ChpNode.EnsureVisible
      BkNode.EnsureVisible
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Reset the background tiling
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, Me.picTile(Background)
End Sub

'*******************************************************************************
' Subroutine Name   : ForceResize
' Purpose           : Force a reformatting of the display
'*******************************************************************************
Private Sub ForceResize()
  ManualResize = True
  Call Form_Resize
  ManualResize = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Adjust everything when teh main form resizes
'*******************************************************************************
Private Sub Form_Resize()
  Dim I As Long
  If Me.WindowState = vbMinimized Then Exit Sub
  If Me.Visible = False Then Exit Sub 'in case if Me.Hide/Me.Show refresh
'
' hide/show certain user items based upon a personal bible being used or not
'
  Me.rtbUser.Visible = BblVersion = UserPVer
  Me.cmdUpdateMPV.Visible = Me.rtbUser.Visible
  Me.cmdCpyXlt.Visible = Me.rtbUser.Visible
  Me.cmdCopympv.Visible = Me.rtbUser.Visible
  
  If Me.WindowState = vbNormal Then
    If Me.Width < 11000 Then Me.Width = 11000 'do not go below minimum twips
    If Me.Height < 7000 Then Me.Height = 7000
  End If
  
  On Error Resume Next  'in case the screen resolution is REALLY small
  With Me.picTree
    .Left = 60
    .Top = Me.picTop.Height + Me.Toolbar1.Height
    .Height = Me.ScaleHeight - Me.StatusBar1.Height - Me.picTop.Height * 2 - Me.Toolbar1.Height
    If (WinState <> Me.WindowState And WinState <> -1) Or ManualResize = False Then
      .Width = orgTvWidth
    Else
     .Width = Me.picVbar1.Left - .Left
    End If
  End With
  
  With Me.picTop
    .Left = 0
    .Top = Me.Toolbar1.Height
    .Width = Me.ScaleWidth
    .BorderStyle = 0
  End With
  
  With Me.tvBooks
    .Top = 0
    .Left = 30
    .Height = Me.picTree.ScaleHeight
    .Width = Me.picTree.ScaleWidth
  End With
  
  With Me.picVbar1
    .Width = 60
    .Left = Me.picTree.Width + Me.picTree.Left
    .Top = Me.picTree.Top
    .Height = Me.picTree.Height
    .BorderStyle = 0
  End With
  
  With Me.picGreek
    .Top = Me.picTree.Top
    .Left = Me.picVbar1.Left + Me.picVbar1.Width
    If (WinState <> Me.WindowState And WinState <> -1) Or ManualResize = False Then
      orgPicGreekWidth = (Me.Width - .Left) / 2 - Me.picVbar2.Width - 80
      GreekWidth = orgPicGreekWidth
      orgPicGreekHeight = (Me.picTree.Height - Me.picHbar1.Height) / 2
      GreekHeight = orgPicGreekHeight
      If NotFirstTimeIn Then WinState = Me.WindowState
      NotFirstTimeIn = True
    End If
    .Width = GreekWidth
    .Height = GreekHeight
    Me.lblPlurality.Left = .Left
    Me.lblPlurality.Top = Me.ScaleHeight - Me.StatusBar1.Height - Me.lblPlurality.Height - 30
  End With
  
  With Me.picGreekControl
    .Left = 0
    .BorderStyle = 0
    .Top = Me.picGreek.ScaleHeight - .Height + 60
    .Width = Me.picGreek.ScaleWidth
    Me.cmdHNext.Top = .ScaleHeight - Me.cmdHNext.Height - 60
    Me.cmdHNext.Left = Me.picGreekControl.ScaleWidth - Me.cmdHNext.Width
    Me.cmdHView.Top = Me.cmdHNext.Top
    Me.cmdHView.Left = Me.cmdHNext.Left - Me.cmdHView.Width - 30
    Me.cmdHBack.Top = Me.cmdHView.Top
    Me.cmdHBack.Left = Me.cmdHView.Left - Me.cmdHBack.Width - 30
    Me.hsGreek.Left = Me.cmdHBack.Top
    Me.hsGreek.Left = .ScaleLeft
    Me.hsGreek.Width = Me.cmdHBack.Left - 120
  End With
  
  With Me.lblVerseIndex
    .Top = -15
    .Left = Me.cmdCopy.Width
    .Width = Me.picGreekControl.ScaleWidth - Me.cmdCopyAll.Width - Me.cmdCopy.Width
  End With
  
  With Me.cmdCopyAll
    .Left = Me.picGreek.ScaleWidth - .Width
    .Top = 0
  End With
  
  With Me.cmdCopy
    .Left = 0
    .Top = 0
  End With
  
  With Me.lstGrkWords
    .Top = Me.lblGrkWords.Height
    .Left = Me.picGreek.ScaleWidth - .Width
    .Height = Me.picGreekControl.Top - Me.lblGrkWords.Height
    I = Me.picGreek.ScaleWidth - Me.lstGrkWords.Width - 60
    If I < 0 Then
      With Me.lstGrkWords
        .Width = Me.picGreek.ScaleWidth / 2
        .Left = Me.picGreek.ScaleWidth - .Width
      End With
    End If
    Me.lblGrkWords.Top = 0
    Me.lblGrkWords.Left = .Left
    Me.lblGrkWords.Width = .Width
  End With
  
  With Me.picVbar3
    .Width = 120
    .Top = 0
    .Left = Me.lstGrkWords.Left - .Width
    .Height = Me.picGreekControl.Top
    .BorderStyle = 0
  End With
    
  With Me.rtbGreek
    .Top = 0
    .Left = 0
    .Height = Me.picGreek.ScaleHeight - Me.picGreekControl.Height + 45
    .Width = Me.picVbar3.Left
  End With

  With Me.picHbar1
    .Height = 60
    .Top = Me.picGreek.Top + Me.picGreek.Height
    .Left = Me.picGreek.Left
    .Left = Me.picGreek.Left
    .Width = Me.ScaleWidth - Me.picGreek.Left
    .BorderStyle = 0
  End With
  
  With Me.picEditor
    .Top = Me.picHbar1.Top + Me.picHbar1.Height
    .Left = Me.picGreek.Left
    .Width = Me.picGreek.Width
    .Height = Me.tvBooks.Height - GreekHeight
  End With
  
  With Me.cmdAdd
    .Top = Me.picEditor.ScaleHeight - .Height
    Me.cmdEdit.Top = .Top
    Me.cmdDel.Top = .Top
    Me.cmdUpdateMPV.Top = .Top
    Me.cmdCpyXlt.Top = .Top
    Me.cmdCopympv.Top = .Top
  End With
  
  With Me.lblSynonyms
    .Top = 0
    .Width = Me.lstWords.Width
    .Left = Me.picEditor.ScaleWidth - .Width
  End With
  
  With Me.lstWords
    .Top = Me.lblSynonyms.Height
    .Left = Me.picEditor.ScaleWidth - .Width
    .Height = Me.cmdAdd.Top - Me.lblSynonyms.Height
    I = Me.picEditor.ScaleWidth - Me.lstWords.Width - 60
    If I < 0 Then
      With Me.lstWords
        .Width = Me.picEditor.ScaleWidth / 2
        .Left = Me.picEditor.ScaleWidth - .Width
        Me.lblSynonyms.Width = .Width
        Me.lblSynonyms.Left = .Left
      End With
    End If
  End With
  
  With Me.cmdFind
    .Top = 0
    .Left = Me.lblSynonyms.Left
  End With
  
  Me.cmdEdit.Left = Me.lstWords.Left + Me.lstWords.Width - Me.cmdEdit.Width
  Me.cmdDel.Left = Me.cmdEdit.Left - Me.cmdDel.Width - 30
  Me.cmdAdd.Left = Me.cmdDel.Left - Me.cmdDel.Width - 30
  
  With Me.picVbar4
    .Width = 120
    .Top = 0
    .Left = Me.lstWords.Left - .Width
    .Height = Me.cmdAdd.Top - 30
    .BorderStyle = 0
  End With
  
  With Me.rtbTranslate
    .Top = 0
    .Left = 0
    .Width = Me.picVbar4.Left
    Me.cmdCpyXlt.Left = Me.cmdUpdateMPV.Left + Me.cmdUpdateMPV.Width + 60
    Me.cmdCopympv.Left = Me.cmdCpyXlt.Left + Me.cmdCpyXlt.Width + 60
    If Me.rtbUser.Visible Then
      .Height = Me.cmdUpdateMPV.Top \ 2
    Else
      .Height = Me.cmdUpdateMPV.Top - 30
    End If
    Me.cmdCopyDT.Left = 0
    Me.cmdCopyDT.Top = .Top + .Height + 15
  End With
    
  With Me.cmdVine
    .Top = 0
    .Left = Me.picEditor.ScaleWidth - .Width
  End With
  
  With Me.cmdCopyDT
    Me.lblFeminine.Top = .Top + .Height - Me.lblFeminine.Height - 30
    Me.lblFeminine.Left = .Left + .Width + 60
  End With
  
  With Me.lblPlural
    .Top = Me.lblFeminine.Top
    .Left = Me.lblFeminine.Left + Me.lblFeminine.Width + 60
    Me.lblEditPersonal.Top = .Top + .Height - Me.lblEditPersonal.Height
    Me.lblEditPersonal.Left = .Left + .Width + 60
  End With
  
  With Me.rtbUser
    If Me.rtbUser.Visible Then
      .Top = Me.cmdCopyDT.Top + Me.cmdCopyDT.Height + 15
    End If
    .Left = 0
    .Width = Me.picVbar4.Left
    .Height = Me.cmdUpdateMPV.Top - .Top - 15
  End With
  
  With Me.picVbar2
    .Width = 60
    .Top = Me.picVbar1.Top
    .Left = Me.picGreek.Left + Me.picGreek.Width
    .Height = Me.picTree.Height
    .BorderStyle = 0
  End With
  
  With Me.picVerse
    .Top = Me.picEditor.Top
    .Left = Me.picVbar2.Left + Me.picVbar2.Width
    .Width = Me.ScaleWidth - Me.picVbar2.Left - Me.picVbar2.Width - 60
    .Height = Me.picEditor.Height
  End With
  
  With Me.rtbVerseNotes
    .Left = 0
    .Top = Me.picVerse.ScaleHeight \ 2
    .Width = Me.picVerse.ScaleWidth - 15
    .Height = .Top - Me.cmdCopyNotes.Height
    Me.cmdCopyNotes.Left = 0
    Me.cmdCopyNotes.Top = .Top + .Height
    Me.cmdAddNote.Top = Me.cmdCopyNotes.Top
    Me.cmdAddNote.Left = .Width - Me.cmdAddNote.Width
  End With
  
  With Me.rtbVerse
    .Top = 0
    .Left = 0
    .Width = Me.picVerse.ScaleWidth - 15
    .Height = Me.picVerse.ScaleHeight \ 2 - Me.cmdCopyVerse.Height - 30
    Me.cmdCopyVerse.Top = .Top + .Height + 15
  End With
  
  With Me.cmdCopyVerse
    Me.lblPersonal.Top = .Top + 60
    Me.lblPersonal.Left = .Left + .Width + 120
  End With
  
  With Me.cmdCopyVerse
    .Left = 0
    Me.lblTheoNote.Left = .Width + 240
    Me.lblTheoNote.Top = Me.picVerse.ScaleHeight - Me.lblTheoNote.Height - 60
    Me.cmdKJV.Top = .Top
    Me.cmdMKJV.Top = .Top
    Me.cmdYLT.Top = .Top
    Me.cmdRSV.Top = .Top
    Me.cmdMPV.Top = .Top
    Me.cmdWEB.Top = .Top
    Me.cmdASV.Top = .Top
    Me.cmdDBY.Top = .Top
    Me.cmdWBS.Top = .Top
    Me.cmdYLT.Left = Me.rtbVerse.Width - Me.cmdYLT.Width
    Me.cmdWEB.Left = Me.cmdYLT.Left - Me.cmdWEB.Width
    Me.cmdWBS.Left = Me.cmdWEB.Left - Me.cmdWBS.Width
    Me.cmdRSV.Left = Me.cmdWBS.Left - Me.cmdRSV.Width
    Me.cmdMPV.Left = Me.cmdRSV.Left - Me.cmdMPV.Width
    Me.cmdMKJV.Left = Me.cmdMPV.Left - Me.cmdMKJV.Width
    Me.cmdKJV.Left = Me.cmdMKJV.Left - Me.cmdKJV.Width
    Me.cmdDBY.Left = Me.cmdKJV.Left - Me.cmdDBY.Width
    Me.cmdASV.Left = Me.cmdDBY.Left - Me.cmdASV.Width
  End With
  
  With Me.picNotes
    .Top = Me.picGreek.Top
    .Left = Me.picVerse.Left
    .Width = Me.picVerse.Width
    .Height = Me.picGreek.Height
  End With
  
  Me.lblTip.Top = 60
  With Me.lblVineRef
    .Top = 60
    .Left = Me.picNotes.Left
    .Visible = False
  End With
  
  With Me.rtbNotes
    .Top = 0
    .Left = 0
    .Width = Me.picNotes.ScaleWidth
    .Height = Me.picNotes.ScaleHeight - Me.cmdCopyNotes.Height
    Me.cmdCopyDef.Left = 0
    Me.cmdCopyDef.Top = .Height
    Me.lblNoteInfo.Left = Me.cmdCopyDef.Width + 240
    Me.lblNoteInfo.Top = Me.picNotes.ScaleHeight - Me.lblNoteInfo.Height - 60
    Me.cmdAnalysis.Top = .Height
    Me.cmdAnalysis.Left = .Width - Me.cmdAnalysis.Width
  End With
  
  With Me.cmdFindNext
    .Top = Me.Toolbar1.Height
    .Height = Me.picTop.Height
    .Left = Me.picNotes.Left + Me.picNotes.Width - .Width - 30
    Me.cmdFindInText.Left = .Left - Me.cmdFindInText.Width
    Me.cmdFindInText.Top = .Top
    Me.cmdFindInText.Height = .Height
    Me.cmdBack.Left = Me.cmdFindInText.Left - Me.cmdBack.Width
    Me.cmdBack.Top = .Top
    Me.cmdBack.Height = .Height
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Form_QueryUnload
' Purpose           : User hit the form "X" button
'*******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Call mnuFileExit_Click
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Remove allowcated resources
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set colHist = Nothing
  Set colHistory = Nothing
  Set colSrch = Nothing
  Set colFavs = Nothing
  Set Fso = Nothing
  Set VerseDropHandler = Nothing
  Set ChapDropHandler = Nothing
  Set BookDropHandler = Nothing
  Set MyToolTips = Nothing

  Unload frmMessageBox
  Unload frmInputBox
  Unload frmShowVineList
  Unload frmView
  Unload frmViewDemo
  Unload frmKJVDict
  
  If Me.WindowState = vbMinimized Then Exit Sub
  
  Call SaveSetting(App.Title, "Settings", "FormState", CStr(Me.WindowState))
  
  If Me.WindowState = vbNormal Then
    Call SaveSetting(App.Title, "Settings", "FormTop", CStr(Me.Top))
    Call SaveSetting(App.Title, "Settings", "FormLeft", CStr(Me.Left))
    Call SaveSetting(App.Title, "Settings", "FormWidth", CStr(Me.Width))
    Call SaveSetting(App.Title, "Settings", "FormHeight", CStr(Me.Height))
  End If
  
  Call SaveSetting(App.Title, "Settings", "Hbar1", CStr(Me.picGreek.Height))
  Call SaveSetting(App.Title, "Settings", "Vbar1", CStr(Me.picTree.Width))
  Call SaveSetting(App.Title, "Settings", "Vbar2", CStr(Me.picGreek.Width))
  Call SaveSetting(App.Title, "Settings", "Vbar3", CStr(Me.lstGrkWords.Width))
  Call SaveSetting(App.Title, "Settings", "Vbar4", CStr(Me.lstWords.Width))
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyDown
' Purpose           : Ctrl Key help down
'*******************************************************************************
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim I As Long
  Dim Ary() As String, S As String
  Dim cBol As Boolean, vBol As Boolean
  
'----------------
' Control pressed
'----------------
  If Shift = vbCtrlMask Then
    vBol = False
    cBol = False
    Select Case KeyCode
' Ctrl-Left arrow - previous verse
      Case 37
        KeyCode = 0
        Vrs = Vrs - 1           'back off a verse
        If Vrs = 0 Then         'bottom of chapter...
          vBol = True           'set verse breached flag
          Chp = Chp - 1         'back off a chapter
          If Chp = 0 Then       'if bottom of book...
            cBol = True         'book breach (must get last chapter of lower book
            Bk = Bk - 1         'back off a book
            If Bk = 0 Then      'bottom of New Covenant?
              Bk = 27           'yes, so set Revelation 22:21
            End If
          End If
          Ary = Split(Books(Bk), ",") 'get number of chapters in this book
          ChpCnt = CLng(Ary(4)) 'get number of chapters
          If cBol Then Chp = ChpCnt
          Call GetVerseCount    'get number of verses in this chapter
          If vBol Then Vrs = VrsCnt
        End If
' Ctrl-up arrow - previous chapter
      Case 38
        KeyCode = 0
        Chp = Chp - 1         'back off a chapter
        If Chp = 0 Then       'book breached?
          Bk = Bk - 1         'back off a book
          If Bk = 0 Then      'bottom of New Covenant?
            Bk = 27           'yes, so set to Revelation
          End If
          Ary = Split(Books(Bk), ",") 'get number of chapters in this book
          ChpCnt = CLng(Ary(4)) 'get number of chapters
          Chp = ChpCnt        'set chapter to last one in book
        End If
        Call GetVerseCount    'get number of verses in this chapter
        Vrs = 1               'set verse 1
' Ctrl-right arrow - next verse
      Case 39
        KeyCode = 0
        Vrs = Vrs + 1         'bump verse
        If Vrs > VrsCnt Then  'if chapter breached
          Call Form_KeyDown(40, vbCtrlMask) 'simulate Ctrl-DA
          Exit Sub
        End If
' Ctrl-down arrow - next chapter
      Case 40
        KeyCode = 0
        Chp = Chp + 1         'bump chapter
        If Chp > ChpCnt Then  'book breached?
          Call Form_KeyDown(40, vbAltMask)  'simulate ALT-DA
          Exit Sub
        End If
        Call GetVerseCount    'get number of verses in this chapter
        Vrs = 1               'set verse 1
      Case Else
        Exit Sub
    End Select
    Call UpdateVerse
    Exit Sub
  End If
'------------
' ALT pressed
'------------
  If Shift = vbAltMask Then
    Select Case KeyCode
' ALT-left arrow/up arrow - previous book
      Case 37, 38
        KeyCode = 0
        Bk = Bk - 1
        If Bk = 0 Then Bk = 27
' ALT-right arrow/down arrow - next book
      Case 39, 40
        KeyCode = 0
        Bk = Bk + 1
        If Bk > 27 Then Bk = 1
      Case Else
        Exit Sub
    End Select
    Chp = 1                       'set to start of book
    Vrs = 1
    Ary = Split(Books(Bk), ",") 'get number of chapters in this book
    ChpCnt = CLng(Ary(4))
    Call GetVerseCount    'get number of verses in this chapter
    Call UpdateVerse
    Exit Sub
  End If
'--------------------------------------
' non-Ctrl/Alt: Navigate words in verse
'--------------------------------------
  If Me.lstGrkWords.ListCount = 0 Then Exit Sub
  If Shift <> 0 Then Exit Sub
  With Me.lstGrkWords
    I = .ListIndex
    Select Case KeyCode
      Case 27                 'ESC
        If Me.cmdBack.Visible Then Me.cmdBack.Value = True
      Case 35                 'END
        KeyCode = 0
        .ListIndex = .ListCount - 1
      Case 36                 'HOME
        KeyCode = 0
        .ListIndex = 0
      Case 37                 'left arrow - previous Greek word
        KeyCode = 0
        If I > 0 Then
          I = I - 1
          .ListIndex = I
        Else
          Call Form_KeyDown(37, vbCtrlMask) 'prev chapter
          Call Form_KeyDown(35, 0)          'end
        End If
      Case 39                 'right arrow - next Greek word
        KeyCode = 0
        I = I + 1
        If I < .ListCount Then
          .ListIndex = I
        Else
          Call Form_KeyDown(39, vbCtrlMask)
        End If
      Case 38                 'up arrow - higher synonym
        KeyCode = 0
        I = Me.lstWords.ListIndex - 1
        If I >= 0 Then
          Me.lstWords.ListIndex = I
          AutoDirty = True    'something has changed
        End If
      Case 40                 'down arrow - lower synonym
        KeyCode = 0
        I = Me.lstWords.ListIndex + 1
        If I < Me.lstWords.ListCount Then
          Me.lstWords.ListIndex = I
          AutoDirty = True    'something has changed
        End If
    End Select
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : GetVerseCount
' Purpose           : Set VrsCnt to number of Verses in the current chapter
'*******************************************************************************
Public Sub GetVerseCount()
  Dim I As Long
  Dim S As String
  
  VrsCnt = 0                                'init to 0 verses
  S = Format$(Bk, "00") & Format$(Chp, "00")
  I = FindExactMatch(Me.lstGrk, S & "01")   'find verse 1
  Do
    VrsCnt = VrsCnt + 1                     'find consecutive verse
  Loop While Left$(Grk(I + VrsCnt), 4) = S
  ChgSCroll = True                          'prevent updates
  Me.hsGreek.Max = VrsCnt - 1               'set as max value for progress bar
  ChgSCroll = False
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAdd_Click
' Purpose           : Add a word to the list of synonyms
'*******************************************************************************
Private Sub cmdAdd_Click()
  Dim S As String
'
' get word to add from the user
'
  S = Me.lstWords.List(Me.lstWords.ListIndex) ' get a "default" to perhaps base word on
  S = Trim$(InputMsgBox(Me, "Enter synonym to add to list (be sure that it is a valid synonym):", "Add synonym", S))
  AddWord S
End Sub

Private Sub AddWord(Text As String)
  Dim Col As Collection
  Dim S As String, Ary() As String, UsrTxt As String, T As String
  Dim Idx As Long, I As Long, Hld As Long, NewIndex As Long
'
' get word to add from the user
'
  S = Trim$(Text)
  If Len(S) = 0 Then
    Me.lstGrkWords.SetFocus
    Exit Sub         'ignore if nothing
  End If
'
' save user text
'
  UsrTxt = Trim$(Me.rtbUser.TextRTF)
'
' use the collection to fleece an accidentally entered duplicate
'
  Set Col = New Collection
  With Me.lstWords
    On Error Resume Next
    Hld = .ListIndex
    For Idx = 0 To .ListCount - 1
      Col.Add .List(Idx), .List(Idx)
    Next Idx
    Err.Clear
    Col.Add S, S                      'add new word
    If Err.Number <> 0 Then           'if word already exists
      For NewIndex = 0 To .ListCount - 1
        If StrComp(S, .List(NewIndex), vbTextCompare) = 0 Then Exit For 'found match
      Next NewIndex
      If NewIndex = .ListIndex Then   'match found?
        MessageBox Me, "This entry already exists. Perhaps you should edit it?", vbOKOnly Or vbExclamation, "Entry Already Exists"
        Exit Sub
      End If
    Else
      NewIndex = .ListCount
    End If
    On Error GoTo 0
    ReDim Ary(Col.Count - 1)          'set new image for array
    .Clear                            'reset displayed list
    Do While Col.Count
      .AddItem Col(1)                 'rebuild list
      Ary(.ListCount - 1) = Col(1)    'stuff in the array
      Col.Remove 1
    Loop
    Set Col = Nothing
    
    S = Join(Ary, ",")                'update word reference database
    Ary = Split(WordRef(Strong), vbTab)
    Ary(2) = S
    WordRef(Strong) = Join(Ary, vbTab)
'
' now update the verse display for the direct translation
'
    With Me.lstGrkWords
      Idx = .ListIndex                        'save word index
      Call UpdateVerse                        'update verse
      If Idx <> .ListIndex Then .ListIndex = Idx  'reset index
    End With
    .ListIndex = NewIndex                     'set
    lstWordsClicked = True                    'prevent redundancy
    .ListIndex = NewIndex                     'reset
    lstWordsClicked = False
  End With
  AutoDirty = True                            'something has changed
  Me.rtbUser.TextRTF = UsrTxt
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDel_Click
' Purpose           : Delete an un-used synonym from the list
'*******************************************************************************
Private Sub cmdDel_Click()
  Dim S As String, Ary() As String, T As String, TT As String, BBLMap() As String, UsrTxt As String
  Dim Idx As Long, Hld As Long, I As Long, AdjustIndex As Long, J As Long
  Dim bChg As Boolean
  
  UsrTxt = Me.rtbUser.TextRTF
  With Me.lstWords
    If .ListCount = 1 Then  'do not allow delete when there is one item present
      MessageBox Me, "Sorry. You cannot delete lists containing one one item.", vbOKOnly Or vbExclamation, "Illegal Option"
      Me.lstGrkWords.SetFocus
      Exit Sub
    End If
    Hld = .ListIndex        'get index to item to delete
    AdjustIndex = Hld - 1   'keep copy of limit checker
    S = .List(Hld)          'get text
    If MessageBox(Me, "Verify deleting this synonym entry: '" & S & "'.", vbYesNo Or vbQuestion Or vbDefaultButton2, "Confirm Delete") = vbNo Then Exit Sub
    .RemoveItem Hld         'if user said yes, remove the word from the list
'
' scan the user-selected word map and adjust any verses that used the deleted word
'
    T = CStr(DefRefIdx)                         'string to check for
    TT = " " & T & " "                          'search mask
    For Idx = 0 To UBound(GrkBBL)
      S = GrkBBL(Idx)                           'grab a line
      If Len(S) <> 0 Then                       'if something there
        If InStr(1, S & " ", TT) <> 0 Then      'found DefRef index?
          Ary = Split(GrkBBL(Idx), " ")         'yes, itemize the table
          BBLMap = Split(WordMap(Idx), " ")     'tab offset index table
          For I = 1 To UBound(Ary)
            If T = Ary(I) Then                  'entry found?
              If BBLMap(I) <> "-1" Then         'was it mapped? (-1 = no)
                J = CLng(BBLMap(I))             'get numeric value
                If J > AdjustIndex Then
                  If J > 0 Then
                    BBLMap(I) = CStr(J - 1)     'yes, so set to new offset index
                  End If
                End If
              End If
            End If
          Next I
          WordMap(Idx) = Join(BBLMap, " ")      'put line back
        End If
      End If
    Next Idx                                    'process all entries
    If BBLWIdx(Strong) > AdjustIndex Then       'adjust main index
      BBLWIdx(Strong) = BBLWIdx(Strong) - 1
    End If
'
' now update the word list for the adjusted entry
'
    S = vbNullString
    For Idx = 0 To .ListCount - 1
      S = S & "," & .List(Idx)
    Next Idx
    Ary = Split(WordRef(Strong), vbTab)
    Ary(2) = Mid$(S, 2)
    Ary(1) = CStr(BBLWIdx(Strong))
    WordRef(Strong) = Join(Ary, vbTab)
'
' now update the verse display for the direct translation
'
    With Me.lstGrkWords
      I = .ListIndex                          'save word index
      Call UpdateVerse                        'update verse
      If I <> .ListIndex Then .ListIndex = I  'reset index
    End With
    lstWordsClicked = True                    'prevent redundancy
    If Hld = .ListCount Then Hld = Hld - 1
    .ListIndex = Hld                          'reset
    lstWordsClicked = False
  End With
  AutoDirty = True                            'something has changed
  Me.rtbUser.TextRTF = UsrTxt
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdEdit_Click
' Purpose           : Edit a synonym (correct spelling of an ADD, or modify to
'                   : a more-used form).
'*******************************************************************************
Private Sub cmdEdit_Click()
  Dim S As String, OldData As String, Ary() As String, UsrTxt As String
  Dim Idx As Long, Hld As Long
  Dim Col As Collection
  
  UsrTxt = Me.rtbUser.TextRTF
  With Me.lstWords
    Hld = .ListIndex                  'save index to selected word
    S = .List(Hld)                    'get the selected word
    OldData = S                       'save a copy
    S = Trim$(InputMsgBox(Me, "Modify the synonym entry: ", "Change Symnonym", S))
    If Len(S) = 0 Then Exit Sub       'if user cancelled
    If S = OldData Then Exit Sub      'if no change
'
' see if the new entry matches one already there
'
    Set Col = New Collection
    On Error Resume Next
    For Idx = 0 To .ListCount - 1
      Col.Add .List(Idx), .List(Idx)
    Next Idx
    Err.Clear
    Col.Add S, S
    If Err.Number <> 0 Then
      If LCase$(S) = LCase$(.List(Hld)) And S <> .List(Hld) Then
      '
      Else
        MessageBox Me, "The changed entry ('" & OldData & "' to '" & S & "') already exists in the list.", vbOKOnly Or vbExclamation, "Ignoring Edit"
        Me.lstGrkWords.SetFocus
        Exit Sub
      End If
    End If
    On Error GoTo 0
'
' does not mat other entries, so update the entry
'
    .List(Hld) = S                    'update entry
'
' update the word reference list
'
    S = vbNullString
    For Idx = 0 To .ListCount - 1     'build a comma-delimited string from the listbox
      S = S & "," & .List(Idx)
    Next Idx
    
    Ary = Split(WordRef(Strong), vbTab)
    Ary(2) = Mid$(S, 2)
    WordRef(Strong) = Join(Ary, vbTab)
    
    Set Col = Nothing                 'release resources
'
' now update the verse display for the direct translation
'
    With Me.lstGrkWords
      Idx = .ListIndex                            'save word index
      Call UpdateVerse                            'update verse
      If Idx <> .ListIndex Then .ListIndex = Idx  'reset index
    End With
    lstWordsClicked = True                        'prevent redundancy
    If Hld = .ListCount Then Hld = Hld - 1
    .ListIndex = Hld                              'reset
    lstWordsClicked = False
  End With
  AutoDirty = True                                'something has changed
  Me.rtbUser.TextRTF = UsrTxt
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopyAll_Click
' Purpose           : Copy Greek, Verse, direct translation (and user def) to clipboard
'*******************************************************************************
Private Sub cmdCopyAll_Click()
  Dim S As String
  
  Clipboard.Clear
  S = Me.rtbGreek.Text & vbCrLf & vbCrLf & _
  Me.rtbVerse.Text & vbCrLf & vbCrLf & _
  Me.rtbTranslate.Text
  If Len(Trim$(Me.rtbUser.Text)) <> 0 Then
    S = S & vbCrLf & vbCrLf & _
      "My Transliteration: " & Ttl & ":" & vbCrLf & vbCrLf & Trim$(Me.rtbUser.Text)
  End If
  Clipboard.SetText S               'save the straight text version
'
' build a common data block in a hidden rich textbox
'
  With Me.rtbMerge
    .TextRTF = Me.rtbGreek.TextRTF
    .SelStart = Len(.Text)
    .SelText = vbCrLf & vbCrLf
    .SelStart = Len(.Text)
    .SelRTF = Me.rtbVerse.TextRTF
    .SelStart = Len(.Text)
    .SelText = vbCrLf & vbCrLf
    .SelStart = Len(.Text)
    .SelRTF = Me.rtbTranslate.TextRTF
    If Len(Trim$(Me.rtbUser.Text)) <> 0 Then
      .SelStart = Len(.Text)
      .SelText = vbCrLf & vbCrLf
      .SelStart = Len(.Text)
      .SelText = "My Personal Transliteration: " & Ttl & ":"
      .SelStart = Len(.Text)
      .SelText = vbCrLf & vbCrLf
      .SelStart = Len(.Text)
      .SelText = Me.rtbUser.TextRTF
    End If
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelColor = vbBlack
    Clipboard.SetText .TextRTF, vbCFRTF
    .Text = vbNullString
  End With
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopy_Click
' Purpose           : Copy Greek to the clipboard
'*******************************************************************************
Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetText Me.rtbGreek.Text              'save base text version
  Clipboard.SetText Me.rtbGreek.TextRTF, vbCFRTF  'save rich text version
  Me.lstGrkWords.SetFocus
'  Me.cmdCopy.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : cmdUpdateMPV_Click
' Purpose           : Update personal version with edited text
'*******************************************************************************
Private Sub cmdUpdateMPV_Click()
  Dim S As String
  Dim I As Long
  
  S = Trim$(Me.rtbUser.Text)
  If Len(S) = 0 Then Exit Sub
  I = InStr(1, S, vbCrLf)
  Do While I <> 0
    S = Left$(S, I - 1) & "\" & Mid$(S, I + 2)
    I = InStr(I + 1, S, vbCrLf)
  Loop
  Bible(UserIndex) = Left$(Bible(UserIndex), 6) & "*" & S
  Me.cmdUpdateMPV.Enabled = False
  With Me.rtbUser
    S = Trim$(.Text)
    Call UpdateVerse
    LockWindowUpdate .hwnd
    .Text = S
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelLength = 0
    LockWindowUpdate 0
  End With
  PVDirty = True
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopympv_Click
' Purpose           : Copy the current personal version verse to the edit window
'*******************************************************************************
Private Sub cmdCopympv_Click()
  Dim S As String
  Dim I As Long
  
  S = Me.rtbVerse.Text
  I = InStr(1, S, "Common KJV")
  If I <> 0 Then S = Left$(S, I - 5)
  I = InStr(1, S, "[")
  Do While I <> 0
    S = Left$(S, I - 1) & Mid$(S, I + 1)
    I = InStr(1, S, "]")
    S = Left$(S, I - 1) & Mid$(S, I + 1)
    I = InStr(1, S, "[")
  Loop
  
  I = InStr(1, S, vbCr)
  With Me.rtbUser
    LockWindowUpdate .hwnd
    .Text = vbNullString
    .Text = Mid$(S, I + 4)
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelStart = 0
    LockWindowUpdate 0
  End With
  Me.cmdUpdateMPV.Enabled = True
'  Me.lstGrkWords.SetFocus
  Me.rtbUser.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCpyXlt_Click
' Purpose           : Copy the direct translation to the edit window
'*******************************************************************************
Private Sub cmdCpyXlt_Click()
  Dim S As String
  Dim I As Long
  
  S = UserText
  I = InStr(1, S, "(s)")
  Do While I <> 0
    S = Left$(S, I - 1) & Mid$(S, I + 3)
    I = InStr(1, S, "(s)")
  Loop

  With Me.rtbUser
    LockWindowUpdate .hwnd
    .Text = vbNullString
    .Text = S
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelStart = 0
    LockWindowUpdate 0
  End With
  Me.cmdUpdateMPV.Enabled = True
'  Me.lstGrkWords.SetFocus
  Me.rtbUser.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : hsGreek_Change
' Purpose           : The Horizontal Scroll value changed
'*******************************************************************************
Private Sub hsGreek_Change()
  If Not ChgSCroll Then
    Vrs = Me.hsGreek.Value + 1
    Call UpdateVerse
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : hsGreek_Scroll
' Purpose           : The user moved the horizontal scroll bar
'*******************************************************************************
Private Sub hsGreek_Scroll()
  If Not ChgSCroll Then
    Vrs = Me.hsGreek.Value + 1
    Call UpdateVerse
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : UpdateBible
' Purpose           : Common routine used to change bibles to view
'*******************************************************************************
Private Sub UpdateBible(ByVal Index As Long)
  Dim S As String
  Dim Idx As Long
  
  Me.mnuBibleMPV.Enabled = PersonalVersion
'
' save user personal bible if it had any updates done to it
'
  If PersonalVersion Then           'if personal version available...
    If Me.mnuBibleMPV.Checked Then  'if currently active...
      If PVDirty Then               'if it is also dirty...
        Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\MPV.txt", ForWriting, True)
        ts.Write Join(Bible, vbCrLf)
        ts.Close
        PVDirty = False             'is no longer dirty
      End If
    End If
  End If
  Me.mnuBibleYLT.Checked = False
  Me.mnuBibleRSV.Checked = False
  Me.mnuBibleMPV.Checked = False
  Me.mnuBibleKJV.Checked = False
  Me.mnuBibleMKJV.Checked = False
  Me.mnuBibleWeb.Checked = False
  Me.mnuBibleASV.Checked = False
  Me.mnuBibleDarby.Checked = False
  Me.mnuBibleWebster.Checked = False
  
  Me.cmdKJV.Enabled = KJVAvail
  Me.cmdMKJV.Enabled = MKJVAvail
  Me.cmdYLT.Enabled = YLTAvail
  Me.cmdRSV.Enabled = RSVAvail
  Me.cmdMPV.Enabled = MPVAvail
  Me.cmdWEB.Enabled = WEBAvail
  Me.cmdASV.Enabled = ASVAvail
  Me.cmdDBY.Enabled = DBYAvail
  Me.cmdWBS.Enabled = WBSAvail
  
  Me.lblEditPersonal.Visible = False
  
  BblVersion = Index
  SaveSetting App.Title, "Settings", "Bible", CStr(BblVersion)
  
  Select Case BblVersion
    Case 1
      S = "YLT"
      VersionText = "Young's Literal Translation"
      Me.mnuBibleYLT.Checked = True
      Me.cmdYLT.Enabled = False
    Case 2
      S = "RSV"
      VersionText = "Revised Standard Version"
      Me.mnuBibleRSV.Checked = True
      Me.cmdRSV.Enabled = False
    Case UserPVer
      S = "MPV"
      VersionText = "My Personal Version"
      Me.mnuBibleMPV.Checked = True
      Me.cmdMPV.Enabled = False
      Me.lblEditPersonal.Visible = True
    Case 4
      S = "MKJV"
      VersionText = "Modern King James Version"
      Me.mnuBibleMKJV.Checked = True
      Me.cmdMKJV.Enabled = False
    Case 5
      S = "WEB"
      VersionText = "World English Bible"
      Me.mnuBibleWeb.Checked = True
      Me.cmdWEB.Enabled = False
    Case 6
      S = "ASV"
      VersionText = "American Standard Version"
      Me.mnuBibleASV.Checked = True
      Me.cmdASV.Enabled = False
    Case 7
      S = "DBY"
      VersionText = "Darby's Translation"
      Me.mnuBibleDarby.Checked = True
      Me.cmdDBY.Enabled = False
    Case 8
      S = "WBS"
      VersionText = "Webster's Translation"
      Me.mnuBibleWebster.Checked = True
      Me.cmdWBS.Enabled = False
    Case Else
      S = "KJV"
      VersionText = "King James Version"
      Me.mnuBibleKJV.Checked = True
      Me.cmdKJV.Enabled = False
  End Select
  If S = "MPV" Then
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\" & S & ".txt", ForReading, False)
  Else
    Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\" & S & ".txt", ForReading, False)
  End If
  Bible = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' set or reset visibility on custom version display information
'
  Me.rtbUser.Visible = BblVersion = UserPVer
  If Me.rtbUser.Visible Then
    Me.rtbTranslate.Height = Me.cmdUpdateMPV.Top \ 2 - 40
    Me.cmdCopyDT.Top = Me.rtbTranslate.Top + Me.rtbTranslate.Height + 20
    Me.rtbUser.Top = Me.cmdCopyDT.Top + Me.cmdCopyDT.Height + 20
    Me.rtbUser.Height = Me.cmdUpdateMPV.Top - Me.rtbUser.Top - 20
  Else
    Me.rtbTranslate.Height = Me.cmdUpdateMPV.Top - 40
    Me.cmdCopyDT.Top = Me.cmdUpdateMPV.Top
  End If
  Me.rtbUser.Font.Size = FntSize
  
  With Me.cmdCopyDT
    Me.lblFeminine.Top = .Top + .Height - Me.lblFeminine.Height - 30
    Me.lblFeminine.Left = .Left + .Width + 60
  End With
  
  With Me.lblPlural
    .Top = Me.lblFeminine.Top
    .Left = Me.lblFeminine.Left + Me.lblFeminine.Width + 60
    Me.lblEditPersonal.Top = .Top + .Height - Me.lblEditPersonal.Height
    Me.lblEditPersonal.Left = .Left + .Width + 60
  End With
  
  With Me.cmdCopyDT
    Me.lblFeminine.Top = .Top + .Height - Me.lblFeminine.Height - 30
    Me.lblFeminine.Left = .Left + .Width + 60
  End With
  
  With Me.lblPlural
    .Top = Me.lblFeminine.Top
    .Left = Me.lblFeminine.Left + Me.lblFeminine.Width + 60
  End With
  
  Me.cmdUpdateMPV.Visible = Me.rtbUser.Visible
  Me.cmdCpyXlt.Visible = Me.rtbUser.Visible
  Me.cmdCopympv.Visible = Me.rtbUser.Visible
  Call UpdateVerse
End Sub

'*******************************************************************************
' Subroutine Name   : lblFeminine_Click
' Purpose           : SHow some info about how gender of words is checked
'*******************************************************************************
Private Sub lblFeminine_Click()
  frmAboutGender.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : lblPersonal_Click
' Purpose           : Ensure personal notes heading is displayed
'*******************************************************************************
Private Sub lblPersonal_Click()
  Dim T As String
  Dim SL As Long
  
  With Me.rtbVerseNotes
    T = .Text
    SL = InStr(1, T, MyPersonalNotes, vbTextCompare)    'find heading
    If SL <> 0 Then
      LockWindowUpdate Me.rtbVerseNotes.hwnd            'disable screen updates
      .SelStart = Len(T)                                'go to bottom
      .SelStart = SL - 1                                'back to top
      .SelLength = Len(MyPersonalNotes)                 'highlight heading
      LockWindowUpdate 0                                'allow refresh
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lblPlural_Click
' Purpose           : Show some help regarding the Plurality testing
'*******************************************************************************
Private Sub lblPlural_Click()
  frmAboutPlurality.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : lblVerseIndex_Click
' Purpose           : Prompt the user for a verse in the chapter to display
'*******************************************************************************
Private Sub lblVerseIndex_Click()
  Dim S As String, Ary() As String
  Dim Idx As Long
  
  S = CStr(Vrs)
  S = Trim$(InputMsgBox(Me, "Enter desired verse in chapter: (1 - " & CStr(VrsCnt) & ")", "Enter verse", S))
  If Len(S) = 0 Then Exit Sub
  If S = CStr(Vrs) Then Exit Sub
  Idx = CLng(Val(S))
  If Idx < 1 Or Idx > VrsCnt Then
    Ary = Split(Books(Bk), ",")
    MessageBox Me, "The selected verse is out of range for " & Ary(3) & ", Chapter " & CStr(Chp) & ".", vbOKOnly Or vbExclamation, "Verse Out of Chapter Range"
  Else
    Vrs = Idx
    Call UpdateVerse
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lstGrkWords_DblClick
' Purpose           : This feature lets you correct an invalid entry in the word
'                   : reference list.  This data tables I drew from were so complicated
'                   : and screwed up that I had to manually edit them, and so,
'                   : even though I double- and triple-checked the data, an
'                   : incorrect Strong's reference number may have slipped through.
'                   :
'                   : This allows you to correct it by setting it to the proper
'                   : Strong's Reference Number.  If you do find an error, please
'                   : notify me at davidgoben@yahoo,com, so I can post an update.
'*******************************************************************************
Private Sub lstGrkWords_DblClick()
  Dim S As String, T As String, Txt As String, Ary() As String
  Dim Idx As Long, I As Long
'
' ignore double-click if CTRL key not held down
'
  If (GetKeyState(VK_CONTROL) And &H80000) = 0 Then Exit Sub 'Cntrl key?
'
' the data to text for
'
  Txt = "[" & Me.lstGrkWords.List(Me.lstGrkWords.ListIndex) & "]"
  T = Trim$(InputMsgBox(Me, "Enter word's correct Strong's Number for " & Txt & ":", _
                        "Correct Strong's Number Index"))
  If T = vbNullString Then Exit Sub   'user cancelled
  If Val(T) = 0 Then Exit Sub
'
' scan the DefRef database for the correct Strong's reference entry
'
  For Idx = 1 To UBound(DefRef)
    S = DefRef(Idx)
    If Len(S) <> 0 Then
      Ary = Split(S, vbTab)     'break up the dabase line
      If Ary(5) = T Then        'match found?
        Exit For
      End If
    End If
  Next Idx
  
  If Idx > UBound(DefRef) Then Exit Sub 'invalid Strong's number if not found
'
' set testing and updating variables
'
  BBLLine(Me.lstGrkWords.ListIndex + 1) = CStr(Idx)
  GrkBBL(GrkIdx) = Join(BBLLine, " ")
  
  With Me.lstGrkWords
    Idx = .ListIndex          'save selection index
    Call UpdateVerse          'update the verse
    If Idx <> 0 Then .ListIndex = Idx 'reset the selection if not the first
  End With
  
  BBLDirty = True             'assume updates
  AutoDirty = True            'somewthing has changed
End Sub

'*******************************************************************************
' Subroutine Name   : lstWords_OLEDragDrop
' Purpose           : Dropping dragged selected text into the synonym list
'*******************************************************************************
Private Sub lstWords_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String
  
  S = Data.GetData(vbCFText)
  AddWord S
End Sub

'*******************************************************************************
' Subroutine Name   : mnuBBLFindPrevNoGreek_Click
' Purpose           : Find the previous verse which contains noa ctual Greek text
'*******************************************************************************
Private Sub mnuBBLFindPrevNoGreek_Click()
  Dim Idx As Long
  Dim S As String
  
  For Idx = VrsIdx - 1 To 0 Step -1
    S = Grk(Idx)
    If Len(S) < 8 Then Exit For
  Next Idx
  If Len(S) > 7 Then
    For Idx = UBound(Grk) - 1 To VrsIdx + 1 Step -1
      S = Grk(Idx)
      If Len(S) < 8 Then Exit For
    Next Idx
  End If
  If Len(S) < 8 Then
    NoGreekSupport S    'update display for the verse indicated by the text header
  End If
  Me.cmdAnalysis.Enabled = False
  Me.mnuFileAnalysis.Enabled = False
  Me.cmdBack.Visible = False
End Sub

'*******************************************************************************
' Subroutine Name   : NoGreekSupport
' Purpose           : This support routine will force the display of a verse,
'                   : display the chapter contents in the Definition panel,
'                   : and highlight and display the indicated verse
'*******************************************************************************
Private Sub NoGreekSupport(Text As String)
  Dim S As String, T As String, TT As String
  Dim Idx As Long, J As Long, K As Long
  
  ForceVerse Text            'update display for the verse indicated by the text header
  Call mnuFileViewChapter_Click
  T = "(" & CStr(Vrs) & ") "        'verse to find
  TT = "(" & CStr(Vrs + 1) & ") "   'next verse
  With Me.rtbNotes
    S = .Text
    Idx = InStr(1, S, T)
    If Idx <> 0 Then
      LockWindowUpdate .hwnd
      J = InStr(Idx + Len(T), S, TT)  'find next
      K = InStr(Idx + Len(T), S, MyPersonalNotes)  'find next
      If K > 0 And K < J Then J = K
      If J = 0 Then J = Len(S) + 1    'if not found, assume T was last
      .SelStart = Len(S)              'to end of chapter
      .SelStart = Idx - 1             'now select to force scrollup
      .SelLength = J - Idx            'highlight target verse
      S = RTrim$(.SelText)
      .SelLength = Len(S)
      LockWindowUpdate 0
    End If
  End With
  Me.lblVineRef.Visible = False
  Me.lblNoteInfo.Visible = False
  Me.cmdCopyDef.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : mnuBBLFindNextNoGreek_Click
' Purpose           : Find the next verse which contains noa ctual Greek text
'*******************************************************************************
Private Sub mnuBBLFindNextNoGreek_Click()
  Dim Idx As Long
  Dim S As String
  
  For Idx = VrsIdx + 1 To UBound(Grk) - 1 'search from current point, upward
    S = Grk(Idx)
    If Len(S) < 8 Then Exit For           'found a verse
  Next Idx
  If Len(S) > 7 Then                      'if not found, wrap around
    For Idx = 0 To VrsIdx - 1
      S = Grk(Idx)
      If Len(S) < 8 Then Exit For
    Next Idx
  End If
  If Len(S) < 8 Then
    NoGreekSupport S    'update display for the verse indicated by the text header
  End If
  Me.cmdAnalysis.Enabled = False
  Me.mnuFileAnalysis.Enabled = False
  Me.cmdBack.Visible = False
End Sub

'*******************************************************************************
' Subroutine Name   : mnuBBLOrgKJV_Click
' Purpose           : Explore the original KJV translation strategy
'*******************************************************************************
Private Sub mnuBBLOrgKJV_Click()
  frmKJVXlate.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuBBLViewAllPersonalNotes_Click
' Purpose           : View all personal notes
'*******************************************************************************
Private Sub mnuBBLViewAllPersonalNotes_Click()
  Dim ibk As Long, iChp As Long, iVrs As Long, J As Long, I As Long
  Dim lstBk As Long, lstChp As Long
  Dim S As String, T As String, Ary() As String, sBk As String, TT As String
  Dim UseBk As Boolean, UseChp As Boolean, UseVerse As Boolean, UseTheo As Boolean
  Dim NotFirst As Boolean, NewData As Boolean
  Dim CtrBk As Long, CtrChp As Long, Idx As Long, SS As Long
  
  frmViewPersonalNotes.Show vbModal, Me
  If bCancel Then Exit Sub
  
  UseBk = CBool(GetSetting(App.Title, "Settings", "PNIncBkH", "0"))
  UseChp = CBool(GetSetting(App.Title, "Settings", "PNIncChpH", "0"))
  If CBool(GetSetting(App.Title, "Settings", "PNCtrBkH", "0")) Then
    CtrBk = rtfCenter
  Else
    CtrBk = rtfLeft
  End If
  If CBool(GetSetting(App.Title, "Settings", "PNCtrChpH", "0")) Then
    CtrChp = rtfCenter
  Else
    CtrChp = rtfLeft
  End If
  UseVerse = CBool(GetSetting(App.Title, "Settings", "PNIncVerse", "0"))
  UseTheo = CBool(GetSetting(App.Title, "Settings", "PNIncTheo", "0"))
  lstBk = 0
  lstChp = 0
  NewData = True
  
  Call InitMerge
  For Idx = 0 To UBound(MyNotes) - 1
    S = MyNotes(Idx)
    If Len(S) > 7 Then
      ibk = CLng(Left$(S, 2))
      iChp = CLng(Mid$(S, 3, 2))
      iVrs = CLng(Mid$(S, 5, 2))
      Ary = Split(Books(ibk), ",")
      sBk = Ary(3)
      
      If lstBk <> ibk Then
        lstBk = ibk
        lstChp = 0
        If UseBk Then
          If NotFirst Then
            AddMergeCrlf vbCrLf & UCase$(sBk), "Arial", FntSize + 4, True, , CtrBk
          Else
            
            NotFirst = True
            AddMergeCrlf UCase$(sBk), "Arial", FntSize + 4, True, , CtrBk
          End If
          NewData = True
        End If
      End If
      
      If lstChp <> iChp Then
        lstChp = iChp
        If UseChp Then
          T = vbCrLf & "Chapter " & CStr(iChp)
          If Not UseBk Then T = vbCrLf & sBk & ", " & T
          If NotFirst Then
            AddMergeCrlf T, "Arial", FntSize + 2, , , CtrChp
          Else
            NotFirst = True
            AddMergeCrlf T, "Arial", FntSize + 2, , , CtrChp
          End If
          NewData = True
        End If
      End If
      
      If UseVerse Then
        If Not UseChp Then
          T = "Chapter " & CStr(iChp) & ":" & CStr(iVrs) & ";"
          If Not UseBk Then T = sBk & ", " & T
          If NewData Then
            AddMergeCrlf T, , FntSize, True, True
          Else
            AddMergeCrlf vbCrLf & T, , FntSize, True, True
            NewData = True
          End If
        End If
        
        T = Mid$(Bible(Idx), 8)
        If Len(T) = 0 Then
          T = "No Bible Verse Data for this verse"
        Else
          If UseChp Then T = "(" & CStr(iVrs) & ") " & T
        End If
        With Me.rtbMerge
          SS = Len(.Text)
          If NewData Then
            AddMergeCrLf2 T, , FntSize, , True
            NewData = False
          Else
            AddMergeCrLf2 vbCrLf & T, , FntSize, , True
          End If
          '
          ' make verses without Greek text obnoxious.
          '
          If Len(Grk(Idx)) < 8 Then
            .SelStart = SS
            .SelLength = Len(.Text) - SS
            .SelColor = vbMagenta
          End If
        End With
        NotFirst = True
      End If
          
      TT = vbNullString
      If UseTheo Then
        J = -1
        I = FindExactMatch(Me.lstVNotes, Left$(S, 6))         'find a match
        If I <> -1 Then                                       'found something...
          J = I                                               'save last-found index
          Do While I <> -1
            TT = TT & vbCrLf & Mid$(VNotes(I), 8)    'add a note
            I = FindExactMatch(Me.lstVNotes, Left$(S, 6), I)  'find another
            If J >= I Then Exit Do                            'ignore if index matches last
          Loop
        End If
        If Len(TT) <> 0 Then
          If Not UseVerse Then
             If Not UseChp Then
               T = "Chapter " & CStr(iChp) & ":" & CStr(iVrs)
             Else
               T = "Verse " & CStr(iVrs) & ":"
             End If
             If Not UseBk Then
               T = sBk & ", " & T
             End If
             T = "Theological Notes: " & T
          Else
            T = "Theological Notes:"
          End If
          AddMergeCrlf T, , FntSize - 2, True, True
          processText Mid$(TT, 3) & vbCrLf & vbCrLf, , FntSize - 2
        End If
      End If
      
      S = Mid$(S, 8)
      T = MyPersonalNotes
      If Not UseVerse Then
        If Len(TT) = 0 Then
          If UseChp = False And Len(TT) = 0 Then
            T = "Chapter " & CStr(iChp) & ":" & CStr(iVrs)
          Else
            T = "Verse " & CStr(iVrs) & ":"
          End If
          If Not UseBk Then
            T = sBk & ", " & T
          End If
        End If
        S = S & vbCrLf
      End If
      With Me.rtbMerge
        SS = Len(.Text)
        AddMergeCrlf T, , FntSize - 2, True, True
        .SelStart = SS
        .SelLength = Len(.Text) - SS
        .SelUnderline = True
        processText S & vbCrLf, , FntSize - 2
        .SelStart = SS
        .SelLength = Len(.Text) - SS
        .SelColor = PNotesColor
      End With
      NewData = False
      NotFirst = True
    End If
  Next Idx
  
  SetIndent                                         'set indenting
  LockWindowUpdate Me.rtbNotes.hwnd
  Me.rtbNotes.Text = vbNullString                   'force reset of scrolling
  Me.rtbNotes.BackColor = clBlue
  Me.rtbNotes.TextRTF = Me.rtbMerge.TextRTF         'stuff merged data to display
  LockWindowUpdate 0
  Me.rtbMerge.Text = vbNullString                   'flush pot
  Me.cmdAnalysis.Enabled = True                     'enable buttons
  Me.mnuFileAnalysis.Enabled = True
  Me.cmdCopyDef.Enabled = True
  Me.lblVineRef.Visible = False
  Me.lblNoteInfo.Visible = False
  Me.cmdBack.Visible = True
End Sub

'*******************************************************************************
' Subroutine Name   : processText
' Purpose           : Process text for addition to the formatted text output
'*******************************************************************************
Private Sub processText(Text As String, _
                     Optional FntName As String = "Times New Roman", _
                     Optional FSize As Long = 0, _
                     Optional Bld As Boolean = False, _
                     Optional Itl As Boolean = False, _
                     Optional Alignment As Long = rtfLeft)
  
  Dim SS As Long, Fnt As Long, I As Long, J As Long
  Dim S As String, GrkTxt As String
  
  Fnt = FSize                     'get font point size
  If Fnt = 0 Then Fnt = 10        'use 10 if not specified
  
  With Me.rtbMerge
    S = Text
    I = InStr(1, S, "{")
    Do While I <> 0
      J = InStr(I + 1, S, "}")
      If J = 0 Then Exit Do         'no match
      SS = Len(.Text)               'get the current length of the RTB text
      .SelStart = SS                'set the selection point to the end of the text
      .SelLength = 0                'usually not needed, but it makes intention clear
      .SelText = Left$(S, I - 1)    'stuff the new text
      .SelStart = SS                'ensure selstart is reset to the start of the new text
      .SelLength = Len(.Text) - SS  'select just the new text
      .SelFontName = FntName        'set the Font style
      .SelFontSize = Fnt            'set the point size
      .SelBold = Bld                'and if we want it enboldened
      .SelItalic = Itl              'and italicized
      .SelUnderline = False
      
      GrkTxt = Mid$(S, I + 1, J - I - 1)
      SS = Len(.Text)               'get the current length of the RTB text
      .SelStart = SS                'set the selection point to the end of the text
      .SelLength = 0                'usually not needed, but it makes intention clear
      .SelText = GrkTxt             'stuff the new text
      .SelStart = SS                'ensure selstart is reset to the start of the new text
      .SelLength = Len(.Text) - SS  'select just the new text
      .SelFontName = "Symbol"       'set the Font style
      .SelFontSize = Fnt            'set the point size
      .SelBold = True               'and if we want it enboldened
      .SelItalic = False            'and italicized
      .SelUnderline = False
      
      S = Mid$(S, J + 1)
      I = InStr(1, S, "{")
    Loop
    AddMerge S, FntName, FntSize, Bld, Itl, Alignment
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuBBLViewPNotesChapter_Click
' Purpose           : View Personal Notes for this chapter
'*******************************************************************************
Private Sub mnuBBLViewPNotesChapter_Click()
  Dim Idx As Long, SS As Long
  Dim S As String, sBook As String, Ary() As String
  Dim HaveNotes As Boolean
  
  Ary = Split(Books(Bk), ",")     'get the title for the Book
  
  sBook = Format$(Bk, "00") & Format$(Chp, "00")
  
  Idx = FindExactMatch(Me.lstGrk, sBook & "01")
  If Idx <> -1 Then
    Do While Left$(MyNotes(Idx), 4) = sBook
      S = MyNotes(Idx)
      If Len(S) > 7 Then
        If Not HaveNotes Then
          InitMerge
          AddMergeCrlf "Personal Notes for: " & Ary(3) & ", Chapter " & CStr(Chp), "Arial", FntSize + 2, True, , rtfCenter
          HaveNotes = True
        End If
        With Me.rtbMerge
          SS = Len(.Text)
          AddMergeCrLf2 vbCrLf & "(" & CStr(CLng(Mid$(S, 5, 2))) & ") " & Mid$(Bible(Idx), 8), , FntSize, , True
          '
          ' make verses without Greek text obnoxious.
          '
          If Len(Grk(Idx)) < 8 Then
            .SelStart = SS
            .SelLength = Len(.Text) - SS
            .SelColor = vbMagenta
          End If
          '
          ' handle personal notes
          '
          SS = Len(.Text)
          AddMergeCrlf "Personal Notes:", , FntSize - 2, True, True
          .SelStart = SS
          .SelLength = Len(.Text) - SS
          .SelUnderline = True
          processText Mid$(S, 8) & vbCrLf, , FntSize - 2
          .SelStart = SS
          .SelLength = Len(.Text) - SS
          .SelColor = PNotesColor
        End With
      End If
      Idx = Idx + 1
    Loop
  End If
'
' now display the accumulated data
'
  If HaveNotes Then
    LockWindowUpdate Me.rtbNotes.hwnd                   'lock display updates for window
    SetIndent
    Me.rtbNotes.Text = vbNullString                     'reset scrolling
    Me.rtbNotes.TextRTF = Me.rtbMerge.TextRTF           'stuff new text
    Me.rtbNotes.BackColor = clBlue                      'light blue
    LockWindowUpdate 0                                  'allow refresh
    InitMerge                                           'clear merge text
    Me.lblVineRef.Visible = False
    Me.lblNoteInfo.Visible = False
    Me.cmdCopyDef.Enabled = True
    Me.cmdAnalysis.Enabled = True
    Me.mnuFileAnalysis.Enabled = True
  Else
    MessageBox Me, "No Personal Notes for " & Ary(3) & ", Chapter " & CStr(Chp) & ".", _
                   vbOKOnly Or vbExclamation, "Not Chapter Notes"
  End If
End Sub

Private Sub mnuBibleDarby_Click()
  If BblVersion <> 7 Then UpdateBible 7
End Sub

'*******************************************************************************
' Subroutine Name   : mnuBibleTranslateKJV_Click
' Purpose           : Toggle an option to quickly translate ancient KJV word
'                   : usage to modern forms
'*******************************************************************************
Private Sub mnuBibleTranslateKJV_Click()
  TranslateKJV = Not Me.mnuBibleTranslateKJV.Checked
  mnuBibleTranslateKJV.Checked = TranslateKJV
  SaveSetting App.Title, "Settings", "TranslateKJV", CStr(TranslateKJV)
  If Not IsLoading Then
    Call UpdateVerse
  End If
End Sub

'*******************************************************************************
' Main menu selection of bible version to view
'*******************************************************************************
Private Sub mnuBibleKJV_Click()
  If BblVersion <> 0 Then UpdateBible 0
End Sub

Private Sub mnuBibleWebster_Click()
  If BblVersion <> 8 Then UpdateBible 8
End Sub

Private Sub mnuBibleYLT_Click()
  If BblVersion <> 1 Then UpdateBible 1
End Sub

Private Sub mnuBibleRSV_Click()
  If BblVersion <> 2 Then UpdateBible 2
End Sub

Private Sub mnuBibleMPV_Click()
  If BblVersion <> 3 Then UpdateBible 3
End Sub

Private Sub mnuBibleMKJV_Click()
  If BblVersion <> 4 Then UpdateBible 4
End Sub

Private Sub mnuBibleWEB_Click()
  If BblVersion <> 5 Then UpdateBible 5
End Sub

Private Sub mnuBibleASV_Click()
  If BblVersion <> 6 Then UpdateBible 6
End Sub

'*******************************************************************************
' main menu selection of background patterns
'*******************************************************************************
Private Sub mnuBKCloth_Click()
  Dim I As Long
  SetPat bkCloth
End Sub

Private Sub mnuBkCustom_Click()
  frmCustomBk.Show vbModal, Me
  If bCancel Then Exit Sub
  SetPat bkCustom
End Sub

Private Sub mnuBKIce_Click()
  SetPat bkIce
End Sub

Private Sub mnuBKParch1_Click()
  SetPat bkParch1
End Sub

Private Sub mnuBKParch2_Click()
  SetPat bkParch2
End Sub

Private Sub mnuBKParch3_Click()
  SetPat bkParch3
End Sub

Private Sub mnuBKRumpled_Click()
  SetPat bkRumpled
End Sub

Private Sub mnuBKStucco_Click()
  SetPat bkStucco
End Sub

Private Sub mnuBKMarble_Click()
  SetPat bkMarble
End Sub

Private Sub mnuBKMarbleTx_Click()
  SetPat bkMarbleTX
End Sub

Private Sub mnuBKWood_Click()
  SetPat bkWood
End Sub

'*******************************************************************************
' Subroutine Name   : SetPat
' Purpose           : Support interface for setting a new pattern
'*******************************************************************************
Private Sub SetPat(Patrn As Long)
  SetPattern Patrn
  Me.Hide   'this little trick will refresh to display with any ipdated formatting
  Me.Show   'with minimal flashing of the screen.
End Sub

'*******************************************************************************
' Subroutine Name   : SetPattern
' Purpose           : Update fields for the new pattern
'*******************************************************************************
Private Sub SetPattern(Patrn As Long)
  Dim Idx As Long
  
  Background = Patrn
  SaveSetting App.Title, "Settings", "Background", CStr(Background)
'
' init menu items
'
  Me.mnuBKCloth.Checked = False
  Me.mnuBKIce.Checked = False
  Me.mnuBKParch3.Checked = False
  Me.mnuBKParch2.Checked = False
  Me.mnuBKParch1.Checked = False
  Me.mnuBKRumpled.Checked = False
  Me.mnuBKStucco.Checked = False
  Me.mnuBKMarble.Checked = False
  Me.mnuBKMarbleTx.Checked = False
  Me.mnuBKWood.Checked = False
  Me.mnuBkCustom.Checked = False
'
' apply colors from actual colors extracted from the background images
'
  Select Case Background
    Case bkCloth
      Me.mnuBKCloth.Checked = True  'mark this option in the menu as selected
      cDark = RGB(206, 206, 206)    'Dark
      cMedium = RGB(226, 226, 226)  'Medium
      cLight = RGB(238, 238, 238)   'Light
      cVLight = RGB(243, 243, 243)  'Very Light
    Case bkIce
      Me.mnuBKIce.Checked = True
      cDark = RGB(222, 222, 239)
      cMedium = RGB(214, 231, 239)
      cLight = RGB(231, 239, 247)
      cVLight = RGB(247, 247, 247)
    Case bkParch2
      Me.mnuBKParch2.Checked = True
      cDark = RGB(202, 187, 168)
      cMedium = RGB(216, 204, 180)
      cLight = RGB(255, 242, 219)
      cVLight = RGB(255, 249, 232)
    Case bkParch3
      Me.mnuBKParch3.Checked = True
      cDark = RGB(248, 192, 83)
      cMedium = RGB(255, 210, 98)
      cLight = RGB(255, 231, 112)
      cVLight = RGB(255, 242, 124)
    Case bkRumpled
      Me.mnuBKRumpled.Checked = True
      cDark = RGB(196, 197, 189)
      cMedium = RGB(214, 216, 205)
      cLight = RGB(223, 225, 212)
      cVLight = RGB(241, 242, 236)
    Case bkStucco
      Me.mnuBKStucco.Checked = True
      cDark = RGB(220, 200, 173)
      cMedium = RGB(246, 224, 183)
      cLight = RGB(249, 225, 197)
      cVLight = RGB(255, 238, 210)
    Case bkMarble
      Me.mnuBKMarble.Checked = True
      cDark = RGB(180, 151, 121)
      cMedium = RGB(196, 167, 137)
      cLight = RGB(208, 179, 149)
      cVLight = RGB(219, 190, 160)
    Case bkMarbleTX
      Me.mnuBKMarbleTx.Checked = True
      cDark = RGB(213, 199, 186)
      cMedium = RGB(228, 215, 198)
      cLight = RGB(245, 233, 221)
      cVLight = RGB(254, 249, 244)
    Case bkWood
      Me.mnuBKWood.Checked = True
      cDark = RGB(229, 170, 94)
      cMedium = RGB(234, 175, 105)
      cLight = RGB(238, 182, 125)
      cVLight = RGB(241, 184, 131)
    Case bkCustom
      Me.mnuBkCustom.Checked = True
      cDark = custDark
      cMedium = custMedium
      cLight = custLight
      cVLight = CustVLight
    Case Else
      Me.mnuBKParch1.Checked = True
      cDark = RGB(230, 198, 134)
      cMedium = RGB(250, 225, 172)
      cLight = RGB(255, 235, 186)
      cVLight = RGB(255, 242, 186)
  End Select
'
' now set up the current window fields with the appropriate colors.
' when other forms load, they will update themselves at that time.
'
  Me.rtbGreek.BackColor = cMedium
  Me.rtbTranslate.BackColor = cMedium
  Me.lstGrkWords.BackColor = cDark
  Me.lstWords.BackColor = cDark
  Select Case Me.rtbNotes.BackColor
    Case clBlue, VeryLight
    Case Else
      Me.rtbNotes.BackColor = cLight
  End Select
  Me.rtbVerseNotes.BackColor = cLight
  Me.rtbUser.BackColor = vbWhite
  Me.lblNoteInfo.ToolTipText = Me.lblNoteInfo.Caption
  Me.lblTheoNote.ToolTipText = Me.lblTheoNote.Caption
  Me.cmdFind.BackColor = cLight
  Me.cmdHBack.BackColor = cLight
  Me.cmdHNext.BackColor = cLight
  Me.cmdVine.BackColor = cLight
  Me.cmdFindInText.BackColor = cLight
  Me.cmdFindNext.BackColor = cLight
  Me.cmdBack.BackColor = cLight
  Me.picBCV.BackColor = cVLight
'
' a special case.  When the verse display contains a verse that the user
' has upedated, the background is Navy blue, so we do not want to mess
' with it at this time.
'
  If Me.rtbVerse.BackColor <> Navy Then Me.rtbVerse.BackColor = cMedium
'
' set the background color for the treeview
'
  DoEvents
  Call SendMessage(Me.tvBooks.hwnd, TVM_SETBKCOLOR, 0, ByVal cDark)  'Change the background color
'
' set toobat background color
'
  SetToolbarBkgColor Me.Toolbar1, cLight
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFavAdd_Click
' Purpose           : Add the current verse to the favorites list
'*******************************************************************************
Private Sub mnuFavAdd_Click()
  Dim Ary() As String, S As String, A As String
  Dim Idx As Long, I As Long, J As Long, Cnt As Long
  
  Ary = Split(Books(Bk), ",")         'get the book title
  S = Ary(3) & " " & CStr(Chp) & ":" & CStr(Vrs)
  On Error Resume Next                'if error, then it already exists
  colFavs.Add S, S
  If Err.Number = 0 Then
    Cnt = CLng(GetSetting(App.Title, "Settings", "FavCnt", "0"))
    If Cnt > 0 Then                   'the count is the new index
      Load Me.mnuFavList(Cnt)         'if not 0 (the prototype) then load a new one
    End If
    On Error GoTo 0
    Me.mnuFavList(Cnt).Visible = True 'make it visible
    Me.mnuFavSep.Visible = True       'since at least 1 item in list, ensure separator is up
'
' bump count and save the count and entry to the registry
'
    Cnt = Cnt + 1
    SaveSetting App.Title, "Settings", "FavCnt", CStr(Cnt)
    Me.mnuFavDel.Enabled = True
'
' Now Sort the List
'
    With Me.lstSort
      .Clear
      For Idx = 1 To colFavs.Count
        .AddItem colFavs.Item(1)
        colFavs.Remove 1
      Next Idx
      If CBool(GetSetting(App.Title, "Settings", "SortBbl", "True")) Then
        For Idx = 27 To 1 Step -1             'scan through books in reverse order
          Ary = Split(Books(Idx), ",")
          A = Ary(3)                          'get current book
          For J = .ListCount - 1 To 0 Step -1 'scan list in reverse
            S = .List(J)                      'get an entry
            I = InStrRev(S, " ")              'find space before Chp:Vrs
            If Left$(S, I - 1) = A Then       'in current book?
              .RemoveItem J                   'yes, so remove the entry
              If colFavs.Count = 0 Then
                colFavs.Add S, S              'was first entry
              Else
                colFavs.Add S, S, 1           'else insert below first
              End If
            End If
          Next J                              'check next entry
        Next Idx                              'check for next book
      Else
        For Idx = 0 To .ListCount - 1         'else simply save in sorted order
          colFavs.Add .List(Idx), .List(Idx)
        Next Idx
      End If
    End With
'
' save data in appropriate order
'
    With colFavs
      For Idx = 1 To .Count
        S = .Item(Idx)
        SaveSetting App.Title, "Settings", "Fav" & CStr(Idx), S 'in registry
        Me.mnuFavList(Idx - 1).Caption = S  'stuff its caption in menu
      Next Idx
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFavDel_Click
' Purpose           : Edit the favorites list
'*******************************************************************************
Private Sub mnuFavDel_Click()
  frmFavs.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFavList_Click
' Purpose           : The user chose a favorite, so go to it
'*******************************************************************************
Private Sub mnuFavList_Click(Index As Integer)
  Dim Ary() As String, S As String
  Dim Idx As Long
  
  S = Me.mnuFavList(Index).Caption            'get the specification
  Idx = InStrRev(S, " ") - 1                  'find the space from the end, in case it is
                                              'something like "3 John 1:2"
  For Bk = 1 To 27
    Ary = Split(Books(Bk), ",")
    If Left$(S, Idx) = Ary(3) Then Exit For   'find out which book it it in
  Next Bk
  ChpCnt = CLng(Ary(4))                       'grab the chapter count
  S = Mid$(S, Idx + 1)                        'strip book name from the text
  Idx = InStr(1, S, ":")                      'find the chapter:verse separator
  Chp = CLng(Left$(S, Idx - 1))               'grab the chapter
  Vrs = CLng(Mid$(S, Idx + 1))                'grab the verse
  Call GetVerseCount                          'get the count of verses in the chapter
  Call UpdateVerse                            'now go and display the selected verse
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileAnalysis_Click
' Purpose           : Save a verse analys of a verse
'*******************************************************************************
Private Sub mnuFileAnalysis_Click()
  Me.cmdAnalysis.Value = True
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileBackup_Click
' Purpose           : Save modified files to a backup location
'*******************************************************************************
Private Sub mnuFileBackup_Click()
  Me.tmrAutoBackup.Enabled = False
  frmBackup.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileTheoNext_Click
' Purpose           : View Next verse containing Theological notes
'*******************************************************************************
Private Sub mnuFileTheoNext_Click()
  Dim Idx As Long, I As Long
  Dim S As String
  
  I = -1
  For Idx = VrsIdx + 1 To UBound(Grk) - 1 'search from current to top
    S = Left$(Grk(Idx), 6)                'get a header strong (BookChapVerse)
    I = FindExactMatch(Me.lstVNotes, S)   'find at least one match
    If I <> -1 Then Exit For              'found one
  Next Idx
  
  If I = -1 Then                          'none found yet?
    For Idx = 0 To VrsIdx - 1             'search from bottom if so
      S = Left$(Grk(Idx), 6)
      I = FindExactMatch(Me.lstVNotes, S)
      If I <> -1 Then Exit For
    Next Idx
  End If
'
' this will always succeed
'
  ForceVerse S                          'update display for the verse indicated by the text header
  Me.cmdCopyDef.Enabled = True
  Me.cmdAnalysis.Enabled = True
  Me.mnuFileAnalysis.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileTheoPrev_Click
' Purpose           : View Previous verse containing Theological notes
'*******************************************************************************
Private Sub mnuFileTheoPrev_Click()
  Dim Idx As Long, I As Long
  Dim S As String
  
  I = -1
  For Idx = VrsIdx - 1 To 0 Step -1       'search from current to bottom
    S = Left$(Grk(Idx), 6)                'get a header strong (BookChapVerse)
    I = FindExactMatch(Me.lstVNotes, S)   'find at least one match
    If I <> -1 Then Exit For              'found one
  Next Idx
  
  If I = -1 Then                          'none found yet?
    For Idx = UBound(Grk) - 1 To VrsIdx + 1 Step -1 'search from top if so
      S = Left$(Grk(Idx), 6)
      I = FindExactMatch(Me.lstVNotes, S)
      If I <> -1 Then Exit For
    Next Idx
  End If
'
' this will always succeed
'
  ForceVerse S                          'update display for the verse indicated by the text header
  Me.cmdCopyDef.Enabled = True
  Me.cmdAnalysis.Enabled = True
  Me.mnuFileAnalysis.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileViewChapter_Click
' Purpose           : View the current chapter
'*******************************************************************************
Private Sub mnuFileViewChapter_Click()
  Dim S As String, T As String, sBook As String, Ary() As String, Ntz As String
  Dim Idx As Long, V As Long, GrkI As Long, I As Long
  Dim HldVerseLines As Boolean
  
  IsRTF = True
  HldVerseLines = VerseLines
  VerseLines = False
  Ary = Split(Books(Bk), ",")
  sBook = Ary(3)
  InitMerge
  AddMergeCrlf UCase$(sBook), "Arial", FntSize + 4, True, , rtfCenter
  AddMergeCrlf sBook & ", Chapter " & CStr(Chp), "Arial", FntSize + 2, True, , rtfCenter
  T = Format$(Bk, "00") & Format$(Chp, "00")
  Idx = FindExactMatch(Me.lstGrk, T & "01")   'find verse 1
    
  Do
    S = Bible(Idx)
    V = CLng(Mid$(S, 5, 2))
    GrkI = FindExactMatch(Me.lstGrk, T & Format$(V, "00"))
    
    I = InStr(S, "\")
    If I <> 0 Then
      Ntz = " " & Mid$(S, I + 1)
      S = Left$(S, I - 1)
      I = InStr(2, Ntz, "\")
      Do While I <> 0
        Mid$(Ntz, I, 1) = " "
        I = InStr(I + 1, Ntz, "\")
      Loop
    Else
      Ntz = vbNullString
    End If
    
    If BblVersion = UserPVer Then
      If Len(Grk(GrkI)) > 7 Then
        AddVerse Mid$(S, 8), V, , FntSize, , , Idx
        If Len(Ntz) <> 0 Then
          With Me.rtbMerge
            I = Len(.Text)
            .SelStart = I
            .SelText = Ntz
            .SelStart = I
            .SelLength = Len(.Text) - I
            .SelColor = RGB(192, 0, 0)
          End With
        End If
      End If
    Else
      If Len(Grk(GrkI)) > 7 Then
        AddVerse Mid$(S, 8), V, , FntSize, , , Idx
      Else
        AddVerse Mid$(S, 8), -V, , FntSize, , , Idx
      End If
      If Len(Ntz) <> 0 Then
        With Me.rtbMerge
          I = Len(.Text)
          .SelStart = I
          .SelText = Ntz
          .SelStart = I
          .SelLength = Len(.Text) - I
          .SelColor = RGB(192, 0, 0)
        End With
      End If
      If Len(MyNotes(Idx)) > 7 Then
        AddVerse Mid$(MyNotes(Idx), 7), 0, , FntSize, , , Idx
      End If
    End If
    Idx = Idx + 1
  Loop While Left$(Bible(Idx), 4) = T
  
  LockWindowUpdate Me.rtbNotes.hwnd           'avoid flashing
  Me.rtbNotes.Text = vbNullString             'reset pointers
  Me.rtbNotes.TextRTF = Me.rtbMerge.TextRTF   'stuff new data
  Me.rtbNotes.BackColor = VeryLight           'set background
  LockWindowUpdate 0                          'allow refreshing
  InitMerge                                   'clear merge data
  
  Me.rtbNotes.ToolTipText = "Double-click verse number to select that verse, or word to view its definition"
  lblNoteInfo.Visible = False
  Me.lblVineRef.Visible = False
  Me.cmdCopyDef.Enabled = True                'allow copying
  Me.cmdAnalysis.Enabled = True
  Me.mnuFileAnalysis.Enabled = True
  Me.lblNoteInfo.Visible = False
  Me.cmdBack.Visible = True
  VerseLines = HldVerseLines
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileCreateBible_Click
' Purpose           : Write a personal copy of the selected Bible
'*******************************************************************************
Private Sub mnuFileCreateBible_Click()
  Dim S As String, T As String, sBook As String, Ary() As String, Ntz As String

  Dim Idx As Long, C As Long, V As Long, B As Long, I As Long
  Dim OB As Long, OC As Long, GrkI As Long
'
' prompt for bible
'
  frmSaveBible.Show vbModal, Me
  If bCancel Then Exit Sub
'
' begin writing bible
'
  On Error Resume Next
  Set ts = Fso.OpenTextFile(FileName, ForWriting, True)
  If Err.Number <> 0 Then
    MessageBox Me, "Cannot save Bible file: " & FileName, vbOKOnly Or vbExclamation, _
                       "Error Writing File"
    Exit Sub
  End If
  On Error GoTo 0
  OB = 0                          'old book
  OC = 0                          'old chapter
'
' bump point sizes of output fonts by a bump factor
'
  BumpFactor = CLng(GetSetting(App.Title, "Settings", "BumpFactor", "1")) * 2
'
' prepare to write data
'
  InitMerge
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  DoEvents
  
  If Me.cmdBack.Visible Then Me.cmdBack.Value = True
  
  With Me.pgrWrite                'set up progress bar
    .Left = 0
    .Top = Me.Toolbar1.Height
    .Width = Me.cmdFindInText.Left
    .Max = UBound(Bible)
    .Value = 0
    .Visible = True
  End With
  DoEvents
  
  DoBible = True                  'indicate writing directly to file
  WriteHeader PNotesColor         'stuff the header if RTF format
  For Idx = 0 To UBound(Bible)
    Me.pgrWrite.Value = Idx
    S = Bible(Idx)
    If Len(S) <> 0 Then
      B = CLng(Left$(S, 2))       'get book
      C = CLng(Mid$(S, 3, 2))     'get chapter
      V = CLng(Mid$(S, 5, 2))     'get verse
      
      If B <> OB Then             'if new book
        If OB <> 0 Then
          If IncludeTheoNotes Then
            DoTheoNotes OB, OC
          End If
          If BookNewPage Then
            AddMerge vbFormFeed
          Else
            AddMergeCrlf " ", , 14 + BumpFactor
          End If
          AddMergeCrlf " ", , 14 + BumpFactor, , , rtfLeft  'ensure we are leftward, here
        End If
        Ary = Split(Books(B), ",")   'get book data
        sBook = Ary(3)            'grab full book name
        AddMergeCrlf UCase$(sBook), "Arial", 18 + BumpFactor, True, , CenterBkHeading
        OB = B
        OC = 0
      End If
      
      If C <> OC Then             'if chapters do not match
        If OC <> 0 And IncludeTheoNotes = True Then
          DoTheoNotes OB, OC
        End If
        AddMergeCrlf " ", , 14
        If Not VerseLines And C <> 1 Then
          AddMergeCrlf vbCrLf & sBook & ", Chapter " & CStr(C), "Arial", 14 + BumpFactor, True, True, CenterChapHeading
        Else
          AddMergeCrlf sBook & ", Chapter " & CStr(C), "Arial", 14 + BumpFactor, True, True, CenterChapHeading
        End If
        If OB > 1 Then WriteChapter 'send out a chapter if not RTF
        OC = C
        DoEvents
      End If
'
' grab personal notes for verse
'
      T = Mid$(MyNotes(Idx), 8)
      If Len(T) <> 0 Then T = " " & T
'
' check which version of bible being printed
'
      If BblVersion = UserPVer Then
        GrkI = FindExactMatch(Me.lstGrk, Format$(B, "00") & Format$(C, "00") & Format$(V, "00"))
        If GrkI >= 0 Then
          '
          ' do not print if text for Greek is blank on personal bible
          '
          If Len(Grk(Idx)) > 7 Then
            If PNotesAbove = True And Len(T) <> 0 Then
              AddVerse T, 0, , FntSize - 2, , True, Idx
            End If
            
            I = InStr(S, "\")
            If I <> 0 Then
              Ntz = " " & Mid$(S, I + 1)
              S = Left$(S, I - 1)
              I = InStr(2, Ntz, "\")
              Do While I <> 0
                Mid$(Ntz, I, 1) = " "
                I = InStr(I + 1, Ntz, "\")
              Loop
            Else
              Ntz = vbNullString
            End If
            
            AddVerse Mid$(S, 8), V, , FntSize, , , Idx
            If Len(Ntz) <> 0 Then
              If IsRTF Then
                ts.Write "\cf4 " & Ntz & "\cf0 "
              Else
                ts.Write Ntz
              End If
            End If
            
            If AddPNotes = True And PNotesAbove = False And Len(T) <> 0 Then
              AddVerse T, 0, , FntSize - 2, , True  'verse 0 = personal note
            End If
            If AddNoteSpace Then AddMergeCrlf " ", , 12 + BumpFactor
          End If
        End If
      Else
        If PNotesAbove = True And Len(T) <> 0 Then
          AddVerse T, 0, , FntSize - 2, , True, Idx
        End If
        I = InStr(S, "\")
        If I <> 0 Then
          Ntz = " " & Mid$(S, I + 1)
          S = Left$(S, I - 1)
          I = InStr(2, Ntz, "\")
          Do While I <> 0
            Mid$(Ntz, I, 1) = " "
            I = InStr(I + 1, Ntz, "\")
          Loop
        Else
          Ntz = vbNullString
        End If
        
        If Len(Grk(Idx)) > 7 Then
          AddVerse Mid$(S, 8), V, , FntSize, , , Idx 'do normally
        Else
          AddVerse Mid$(S, 8), -V, , FntSize, , , Idx 'do grayed if verse does not actually exist
        End If
        
        If Len(Ntz) <> 0 Then
          If IsRTF Then
            ts.Write "\cf4 " & Ntz & "\cf0 "
          Else
            ts.Write Ntz
          End If
        End If
        
        If AddPNotes = True And PNotesAbove = False And Len(T) <> 0 Then
          AddVerse T, 0, , FntSize - 2 + BumpFactor, , True, Idx
        End If
        If AddNoteSpace Then AddMergeCrlf " ", , 12 + BumpFactor
      End If
    End If
  Next Idx
  If IncludeTheoNotes Then
    DoTheoNotes OB, OC
  End If
  WriteTrailer                    'terminate file data
  ts.Close                        'all done
  DoBible = False                 'disable flag
  Me.Enabled = True               'return state to normal
'
' save version of Bible based upon
'
  SaveSetting App.Title, "Settings", "BibleBase", VersionText
  Screen.MousePointer = vbDefault 'show not busy
'
' ask if they want to view it
'
  Me.pgrWrite.Visible = False     'remove the progress bar
  If MessageBox(Me, "Bible File Saved to: " & FileName & "." & vbCrLf & vbCrLf & "Would you like to view it?", _
                    vbYesNo Or vbQuestion, "Bible File Saved") = vbYes Then
    DoEvents
    SaveBible = FileName
    frmView.Show vbModeless, Me
  End If
  Me.mnuFileViewSavedBible.Enabled = True
  BumpFactor = 0
  VerseLines = False
End Sub

'*******************************************************************************
' Subroutine Name   : DoTheoNotes
' Purpose           : Process theological notes
'*******************************************************************************
Private Sub DoTheoNotes(ByVal lBk As Long, lChp As Long)
  Dim T As String, lHdr As String, Ary() As String, vData As String
  Dim Idx As Long, I As Long, J As Long, Vcnt As Long
  Dim DidNotes As Boolean
  
  T = Format(lBk, "00") & Format(lChp, "00")
  Vcnt = 0                                  'init to 0 verses
  I = FindExactMatch(Me.lstGrk, T & "01")   'find verse 1
  Do
    Vcnt = Vcnt + 1                         'find consecutive verse
  Loop While Left$(Grk(I + Vcnt), 4) = T
  
  lHdr = "000000"
  FntSize = FntSize - 2                     'back off fount size by 2
  For Idx = 1 To Vcnt
    vData = T & Format(Idx, "00")
    I = FindExactMatch(Me.lstVNotes, vData)
    J = I
    Do While I <> -1
      If Left$(VNotes(I), 4) = T Then
        If Left$(VNotes(I), 4) <> Left$(lHdr, 4) Then
          Ary = Split(Books(lBk), ",")
          AddMergeCrlf vbCrLf & vbCrLf & "Theological Notes for " & Ary(3) & _
                       ", Chapter " & CStr(lChp), , FntSize - 2 + BumpFactor, True, True
        End If
        If Mid$(VNotes(I), 5, 2) <> Right$(lHdr, 2) Then
          AddMergeCrlf "Verse " & CStr(CLng(Mid$(VNotes(I), 5, 2))), , _
                      FntSize - 4 + BumpFactor, False, True
        End If
        lHdr = Left$(VNotes(I), 6)
        AddMerge "Note: ", , FntSize + BumpFactor, True
        MergeCheck (Mid$(VNotes(I), 8)) & vbCrLf
        DidNotes = True
      End If
      I = FindExactMatch(Me.lstVNotes, vData, I)
      If J >= I Then Exit Do
    Loop
  Next Idx
  
  If DidNotes Then
    AddMerge " ", , 14
  End If
  FntSize = FntSize + 2                     'reset font size
End Sub

'*******************************************************************************
' Subroutine Name   : lstGrkWords_MouseMove
' Purpose           : When moving over word in the listbox, set the tooltip for the
'                   : listbox to the word the mouse is over if the text is longer than
'                   : the listbox
'*******************************************************************************
Private Sub lstGrkWords_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String, Tip As String
  Dim Plural As Boolean
  Static LastTip As String
  
  S = GetStringFromMouseMove(Me.lstGrkWords, X, Y)  'get the text under the cursor
  If S <> LastTip Then
    LastTip = S
    If Len(S) > 0 Then
      If S = Me.lstGrkWords.List(Me.lstGrkWords.ListIndex) Then
        Tip = vbNullString                          'redundant, but clarifying
      Else
        Tip = " and positioning"
      End If
      If Right$(S, 1) = "V" Then S = Left$(S, Len(S) - 1) & "s"
      S = """" & S & """. Click for definition" & Tip
    End If
    MyToolTips.ToolText(Me.lstGrkWords) = S         'set tooltip
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : CheckWord
' Purpose           : Check sex of word
'*******************************************************************************
Private Sub CheckWord(Text As String)
  Dim T As String
  Dim MayBeFeminine As Boolean
  
  T = vbNullString
  If Right$(Text, 1) = "h" Or Right$(Text, 3) = "aiV" Or Right$(Text, 3) = "ioV" Then
    MayBeFeminine = True
  ElseIf Right$(Text, 1) = "a" Then
    Select Case Right$(Text, 2)
      Case "ia", "ka", "ra", "qa"
        MayBeFeminine = True
    End Select
  Else
    Select Case Right$(Text, 2)
      Case "hV", "hn", "ai", "aV", "eV", "ew", "iV", "ar", "wV"
        MayBeFeminine = True
    End Select
  End If
  Me.lblFeminine.Visible = MayBeFeminine    'set masculine/feminine
  Me.lblPlural.Visible = CheckPlural(Text)  'check plurality
End Sub

'*******************************************************************************
' Subroutine Name   : CheckUsage
' Purpose           : Check for hints of word usage
'*******************************************************************************
Private Sub CheckUsage(Text As String)
  Select Case Text
    Case "o", "h", "oi", "ai", "touV"
      Favorz = "The"
    Case "tau", "thV", "twn", "toiV"
      Favorz = "This, These"
    Case "tw", "th", "toiV", "taiV", "twiV"
      Favorz = "The"
    Case "ton", "thn", "tauV", "taV"
      Favorz = "That, Those"
    Case "to", "ta"
      Favorz = "The, That, Those"
    
    Case "egw", "emou", "mou", "emoi", "moi", "eme", "me"
      Favorz = "I"
    Case "Vu", "Vou", "Voi", "Ve"
      Favorz = "You"
    Case "hmeiV", "hmwn", "hmin", "hmaV"
      Favorz = "We"
    Case "umeiV", "umwn", "umin", "umaV"
      Favorz = "You all"
    Case "autoV", "autou", "outoV", "outou"
      Favorz = "He, Him, This"
    Case "autw"
      Favorz = "He, Him, It"
    Case "auton"
      Favorz = "He, Him, That"
    Case "auth", "authV", "outh", "outhV"
      Favorz = "She, Her, This"
    Case "authn", "outhn"
      Favorz = "She, Her, That"
    Case "autu", "outu"
      Favorz = "It, This"
    Case "auto", "outo"
      Favorz = "It, This, That"
    Case "autoi", "autoiV", "autai", "autaiV", "autiV"
      Favorz = "They, Them"
    Case "outoi", "outoiV", "outai", "outaiV", "outiV"
      Favorz = "They, Them"
    Case "autwn", "outwn"
      Favorz = "They, These"
    Case "umaV"
      Favorz = "You all"
    Case "autaV", "outaV"
      Favorz = "They, Those"
    Case "auta", "outa"
      Favorz = "They, Them, Those"
    Case "auta", "outa"
      Favorz = "They (it)"
    Case "hmeiV"
      Favorz = "We"
    Case "umeuV"
      Favorz = "You all"
    Case "tou"
      Favorz = "This, These"
    Case Else
      Favorz = vbNullString
  End Select
  
  If Favorz = vbNullString Then
    If Len(Text) > 1 And Right$(Text, 1) = "w" Then
      Favorz = "implied (I)"
    ElseIf Len(Text) > 2 And Right$(Text, 2) = "ei" Then
      Favorz = "implied (He, She, It)"
    ElseIf Len(Text) > 3 And Right$(Text, 3) = "eiV" Then
      Favorz = "implied (You)"
    ElseIf Len(Text) > 3 And Right$(Text, 3) = "ete" Then
      Favorz = "implied (You all)"
    ElseIf Len(Text) > 4 And Right$(Text, 4) = "omen" Then
      Favorz = "implied (We)"
    ElseIf Len(Text) > 4 And Right$(Text, 4) = "amen" Then
      Favorz = "implied (We)"
    ElseIf Len(Text) > 5 And Right$(Text, 5) = "oiuiv" Then
      Favorz = "implied (They all)"
    ElseIf Len(Text) > 5 And Right$(Text, 3) = "uiv" Then
      Favorz = "implied (They all)"
    ElseIf Len(Text) > 3 And Right$(Text, 3) = "ouV" Then
      Favorz = "Them, To them, Their"
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lstGrkWords_Click
' Purpose           : User clicked a word in the greek list of words
'*******************************************************************************
Private Sub lstGrkWords_Click()
  Dim S As String, Txt As String, sAry() As String, tAry() As String
  Dim I As Long, M As Long, K As Long
  
  Me.lstWords.Clear
  With Me.lstGrkWords
    If .ListCount = 0 Then Exit Sub
    Call CheckWord(.List(.ListIndex))       'check sex and plurality of word
    Call CheckUsage(.List(.ListIndex))      'check for hints of usage
  End With
'
' get the list of text, strip off the leader
'
  S = Me.rtbGreek.Text
  K = InStr(1, S, vbCr) + 3                 'point to the start of the Greek data
  S = Mid$(S, K + 1)
  M = 0
'
' find the word position in the Greek text box
'
  For I = 0 To Me.lstGrkWords.ListIndex - 1 'find the postion from the list selection
    M = InStr(M + 1, S, " ")
  Next I
  I = InStr(M + 1, S, " ")                  'find the end of the word
  If I = 0 Then I = Len(S) + 1              'end of text+1 if it was the last word
  I = I - M - 1                             'length of data
  With Me.rtbGreek
    .SelStart = M + K                       'display selection
    .SelLength = I
  End With
'
' now do the same for the direct translation text
'
  With Me.rtbTranslate
    S = Mid$(.Text, DTHeaderOffset)    'skip past header
    M = 0
    For I = 0 To Me.lstGrkWords.ListIndex - 1 'skan for indexed word
      M = InStr(M + 1, S, "]")
    Next I
    I = InStr(M + 1, S, "]")
    If M = 0 Then M = -1
    I = I - M - 3
    .SelStart = M + DTHeaderOffset + 1   'highlight in text box
    .SelLength = I
  End With
'
' now display the selected word information
'
  S = BBLLine(Me.lstGrkWords.ListIndex + 1) 'get the Def list index for the word
  Call ShowDefRef(S)                        'display data
  Call ShowKJVCount(S)
  Call ShowKJVXlate(S)
  Call ShowGrkWrdCnt(S)
  tAry = Split(WordRef(Strong), vbTab)      'grab list of words from word list
  sAry = Split(tAry(2), ",")
  For I = 0 To UBound(sAry)
    Me.lstWords.AddItem sAry(I)             'add synonyms to the list
  Next I
  lstWordsClicked = True                    'set a flag that prevents double processes
  I = MiniMap(Me.lstGrkWords.ListIndex + 1) 'get current user word
  On Error Resume Next
  Me.lstWords.ListIndex = I
  If Err.Number <> 0 Then
    MiniMap(Me.lstGrkWords.ListIndex + 1) = 0
    Me.lstWords.Clear
  Else
    Me.cmdEdit.Enabled = True
    Me.cmdAdd.Enabled = True
    Me.cmdDel.Enabled = Me.lstWords.ListCount > 1
    With Me.lstWords
      If .ListIndex <> -1 Then
        S = LCase$(.List(.ListIndex))
        VineIndex = FindExactMatch(Me.lstVine, S)
        Me.cmdVine.Visible = VineIndex <> -1
        Me.cmdVine.Enabled = Me.cmdVine.Visible
        Me.cmdVine.ToolTipText = "View Vine reference for """ & .List(.ListIndex) & """"
      Else
        VineIndex = -1
        Me.cmdVine.Visible = False = False
      End If
    End With
  End If
  On Error Resume Next
  Me.lstGrkWords.SetFocus
  lstWordsClicked = False
  Me.lblVineRef.Visible = False
  Me.cmdAnalysis.Enabled = True
  Me.mnuFileAnalysis.Enabled = True
  Me.cmdCopyDef.Enabled = True
  Me.cmdBack.Visible = False
End Sub

'*******************************************************************************
' Subroutine Name   : ShowDefRef
' Purpose           : Show the definition of a Greek word, or a Strong Ref #
'*******************************************************************************
Public Sub ShowDefRef(Dref As String, Optional ShowGrkWord As Boolean = True)
  Dim S As String, S1 As String, S2 As String, S3 As String, sAry() As String
  Dim T As String, TT As String
  Dim I As Long, J As Long
  
  S = Dref
  
  If IsNumeric(S) Then                      'is a number
    DefRefIdx = CLng(S)                     'grab number
    sAry = Split(DefRef(DefRefIdx), vbTab)  'grab definition data
    S1 = Me.rtbGreek.SelText & _
                "    (" & sAry(1) & ")    " 'append base word
    S2 = sAry(2)                            'get English direct translation
    Strong = CLng(sAry(5))                  'grab Strong's index number
    S3 = "    (" & sAry(3) & ")    Strong's Reference # " & CStr(Strong) & vbCrLf
    If Len(sAry(4)) <> 0 Then
      S3 = S3 & sAry(4) & vbCrLf
      S = LCase$(sAry(4))
      If InStr(1, S, "feminine") Then
          Me.lblFeminine.Visible = True
      Else
        Me.lblFeminine.Visible = False
      End If
    End If
    S = sAry(6)                             'grab the display text
    If Len(Favorz) <> 0 Then
      S = S & "\\This particular usage may favor a sense of:  " & Favorz & "."
    End If
    I = InStr(1, S, "\")                    'covert all "\" to vbCrLf
    Do While I <> 0
      S = Left$(S, I - 1) & vbCrLf & Mid$(S, I + 1)
      I = InStr(I + 2, S, "\")
    Loop
    With Me.rtbNotes
      LockWindowUpdate .hwnd
      .BackColor = cLight
      If Not ShowGrkWord Then
        I = InStr(1, S1, " ")
        If I <> 0 Then
          S1 = LTrim$(Mid$(S1, I + 1))
        End If
      End If
      InitMerge
      
      T = S1 & S2 & S3 & vbCrLf & S
      J = InStr(1, T, "{")                    'handle embedded Greek words
      Do While J <> 0
        AddMerge Left$(T, J - 1), , FntSize
        I = InStr(J + 1, T, "}")
        If I = 0 Then Exit Do                 'no match
        TT = Mid$(T, J + 1, I - J - 1)
        AddMerge TT, "Symbol", FntSize, True
        T = Mid$(T, I + 1)
        J = InStr(1, T, "{")
      Loop
      If Len(T) <> 0 Then
        AddMerge T, , FntSize
      End If
      
      .TextRTF = Me.rtbMerge.TextRTF          'stuff merged text
      Me.rtbMerge.Text = vbNullString         'clear merge buffer
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelBold = False
      .SelItalic = False
      .SelLength = Len(S1)
      .SelFontName = "Symbol"               'display Greek in Greek
      .SelStart = Len(S1)
      .SelLength = Len(S2)
      .SelItalic = True
      .SelStart = 0
      .SelLength = Len(S1 & S2)
      .SelBold = True
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelHangingIndent = HIndent
      .SelFontSize = FntSize
      .SelLength = 0
      LockWindowUpdate 0
    End With
  Else
    With Me.rtbNotes
      LockWindowUpdate .hwnd
      .Text = S
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelFontName = "Symbol"
      .SelBold = True
      .SelItalic = False
      .SelStart = 0
      LockWindowUpdate 0
    End With
  End If
  Me.cmdBack.Visible = True
End Sub

'*******************************************************************************
' Subroutine Name   : ShowKJVCount
' Purpose           : Display KJV word usage counts
'*******************************************************************************
Public Sub ShowKJVCount(Dref As String)
  Dim S As String, sAry() As String, T As String, ST As String
  Dim I As Long, J As Long, K As Long, TL As Long
  
  sAry = Split(DefRef(CLng(Dref)), vbTab)
  S = sAry(5) & vbTab
  J = Len(S)
  K = UBound(KJVCount)
  For I = 1 To K - 1
    If Left$(KJVCount(I), J) = S Then Exit For
  Next I
  If I = K Then Exit Sub
  T = vbCrLf & vbCrLf & "KJV Word Usage            Count" & vbCrLf
  TL = Len(T)
  Do While Left$(KJVCount(I), J) = S
    sAry = Split(KJVCount(I), vbTab)
    ST = Trim$(sAry(1))
    If StrComp(ST, "not tr.") = 0 Then
      ST = "[not translated]"
    ElseIf StrComp(ST, "not translated") = 0 Then
      ST = "[not translated]"
    ElseIf StrComp(ST, "miscellaneous") = 0 Then
      ST = "[" & ST & "]"
    ElseIf Left$(ST, 3) = "vr " Then
      ST = "[various]" & Mid$(ST, 3)
    End If
    ST = ST & ":"
    If Len(ST) < 26 Then ST = ST & String$(26 - Len(ST), " ")
    T = T & ST & " " & sAry(2) & vbCrLf
    I = I + 1
  Loop
  If Len(T) <> 0 Then
    T = Left$(T, Len(T) - 2)
    With Me.rtbNotes
      LockWindowUpdate .hwnd
      I = Len(.Text)
      .SelStart = I
      .SelText = T
      .SelStart = I
      .SelLength = Len(.Text) - I
      .SelFontName = "Courier New"
      .SelFontSize = FntSize - 2
      .SelColor = Navy
      .SelBold = False
      .SelLength = TL
      .SelColor = vbBlue
      .SelUnderline = True
      .SelBold = True
      .SelStart = 0
      LockWindowUpdate 0
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ShowGrkWrdCnt
' Purpose           : Show statistics on Greek word usage
'*******************************************************************************
Private Sub ShowGrkWrdCnt(Dref As String)
  Dim S As String, sAry() As String, T As String, TT As String
  Dim I As Long, J As Long, K As Long, KK As Long, TL As Long
  Dim cGWC As Collection
  
  Set cGWC = New Collection
  T = Me.lstGrkWords.List(Me.lstGrkWords.ListIndex)
  TT = T & vbTab
  KK = Len(TT)
  S = vbTab & Dref
  K = Len(S)
  J = -1
  For I = 0 To UBound(GrkWrdCnt) - 1
    If Right$(GrkWrdCnt(I), K) = S Then
      cGWC.Add I
      If Left$(GrkWrdCnt(I), KK) = TT Then J = cGWC.Count
    End If
  Next I
'
' if no direct match yet, find the actual word in the list (it WILL exist)
'
  If J = -1 Then
    For I = 0 To UBound(GrkWrdCnt) - 1
      If Left$(GrkWrdCnt(I), KK) = TT Then
        cGWC.Add I
        J = cGWC.Count
        Exit For
      End If
    Next I
  End If
  
  S = vbCrLf & vbCrLf & "New Covenant Greek Word Usage" & vbCrLf
  If J <> -1 Then 'if a direct match was made
    sAry = Split(GrkWrdCnt(CLng(cGWC(J))), vbTab)
    T = """{" & sAry(0) & "}"" is used in the New Covenant " & sAry(1) & " time"
    If CLng(sAry(1)) <> 1 Then T = T & "s"
    T = T & "." & vbCrLf
  End If
  With cGWC
    If cGWC.Count > 1 Then
      If J <> -1 Then .Remove J 'remove any direct match
      T = T & vbCrLf & CStr(.Count) & " other Greek word"
      If .Count > 1 Then
        T = T & "s are"
      Else
        T = T & " is"
      End If
      T = T & " also assigned this definition:" & vbCrLf
      I = 1
      Do While .Count
        sAry = Split(GrkWrdCnt(CLng(.Item(1))), vbTab)
        TT = Format(I, "0")
        If Len(TT) = 1 Then TT = " " & TT
        T = T & TT & ") ""{" & sAry(0) & "}"" is used " & sAry(1) & " time"
        If CLng(sAry(1)) <> 1 Then T = T & "s"
        T = T & "." & vbCrLf
        .Remove 1
        I = I + 1
      Loop
    End If
  End With
  
  With Me.rtbNotes
    LockWindowUpdate .hwnd
    I = Len(.Text)
    .SelStart = I
    .SelText = S
    .SelStart = I
    .SelLength = Len(.Text) - I
    .SelFontName = "Courier New"
    .SelFontSize = FntSize - 2
    .SelBold = True
    .SelColor = vbBlue
    .SelUnderline = True
    .SelStart = 0
    
    InitMerge
    J = InStr(1, T, "{")
    Do While J <> 0
      AddMerge Left$(T, J - 1), , FntSize
      I = InStr(J + 1, T, "}")
      If I = 0 Then Exit Do                 'no match
      TT = Mid$(T, J + 1, I - J - 1)
      AddMerge TT, "Symbol", FntSize - 2, True
      T = Mid$(T, I + 1)
      J = InStr(1, T, "{")
    Loop
    If Len(T) <> 0 Then
      AddMerge T, , FntSize
    End If
    Me.rtbMerge.BackColor = Me.rtbNotes.BackColor
    I = Len(.Text)
    .SelStart = I
    .SelRTF = Me.rtbMerge.TextRTF
    .SelStart = I
    .SelLength = Len(.Text) - I
    .SelColor = Navy
    .SelStart = 0
    LockWindowUpdate 0
    Me.rtbMerge.Text = vbNullString
  End With
  Set cGWC = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : ShowKJVXlate
' Purpose           : Show KJV translation of this word in the verse
'*******************************************************************************
Private Sub ShowKJVXlate(Dref As String)
  Dim VsWords() As String, VsIndex() As String, S As String, T As String
  Dim Idx As Long, I As Long, J As Long
  Dim cWrds As Collection
  
  Set cWrds = New Collection
  VsIndex = Split(Mid$(KJVidxAry(VrsIdx), 8), ",")
  VsWords = Split(Mid$(KJVwrdAry(VrsIdx), 8), ",")
  With cWrds
    For Idx = 0 To UBound(VsIndex)
      If VsIndex(Idx) = Dref Then
        On Error Resume Next
        .Add VsWords(Idx), VsWords(Idx)
        On Error GoTo 0
      End If
    Next Idx
    If .Count = 1 Then
      T = vbCrLf & "  " & .Item(1)
      If .Item(1) = "*" Then T = T & " [not translated]"
      .Remove 1
    Else
      J = 0
      Do While .Count
        J = J + 1
        S = CStr(J)
        If Len(S) = 1 Then S = " " & S
        T = T & vbCrLf & S & ") " & .Item(1)
        If .Item(1) = "*" Then T = T & " [not translated]"
        .Remove 1
      Loop
    End If
    If Len(T) <> 0 Then
      S = "Here, the KJV translates this word to:"
      With Me.rtbNotes
        LockWindowUpdate .hwnd
        I = Len(.Text)
        .SelStart = I
        .SelText = vbCrLf & vbCrLf & S & T
        .SelStart = I
        .SelLength = Len(.Text) - 1
        .SelFontName = "Times New Roman"
        .SelFontSize = FntSize
        .SelColor = Navy
        .SelBold = False
        .SelUnderline = False
        .SelStart = I + 4
        .SelLength = Len(S)
        .SelFontName = "Courier New"
        .SelFontSize = FntSize - 2
        .SelColor = vbBlue
        .SelBold = True
        .SelUnderline = True
        .SelStart = 0
        LockWindowUpdate 0
      End With
    End If
  End With
  Set cWrds = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : lstWords_Click
' Purpose           : User selected a synonym
'*******************************************************************************
Private Sub lstWords_Click()
  Dim S As String, T As String, Ary() As String, Strng As String
  Dim sAry() As String, sIdx As Long
  Dim I As Long, Idx As Long, J As Long
  
  If lstWordsClicked Then Exit Sub          'ignore redundancy
  lstWordsClicked = True                    'set redudancy protections
  Idx = Me.lstWords.ListIndex               'get the word index
  BBLWIdx(Strong) = Idx                     'save as last-selected for this word group
  I = Me.lstGrkWords.ListIndex              'get the index to the selected synonym
  sIdx = MiniMap(I + 1)                     'save copy of old index
  MiniMap(I + 1) = Me.lstWords.ListIndex    'save the choice to the verse map
'
' if the user also held the CTRL key down, then change all matching instances
'
  If GetKeyState(VK_CONTROL) < 0 Then
    T = CStr(Me.lstWords.ListIndex)             'string version of new word index
    Ary = Split(Mid$(GrkBBL(VrsIdx), 8), " ")
    Strng = CStr(Strong)
    For J = 0 To UBound(Ary)
      sAry = Split(DefRef(CLng(Ary(J))), vbTab)
      If sAry(5) = Strng Then
        If MiniMap(J + 1) = sIdx Then
          MiniMap(J + 1) = T
        End If
      End If
    Next J
  End If
  WordMap(GrkIdx) = Join(MiniMap, " ")      'and update the overall bible map
  
  S = Trim$(Me.rtbUser.TextRTF)                   'save a copy of any user-edited data
  LockWindowUpdate Me.hwnd                        'keep screen from flashing
  With Me.lstWords
    Idx = .ListIndex                              'save selected word
    With Me.lstGrkWords
      Call UpdateVerse                            'update verse
      If I <> .ListIndex Then .ListIndex = I      'reset index
    End With
    Me.cmdUpdateMPV.Enabled = CBool(Len(S))       'enable update button if there is data
    .ListIndex = Idx
    T = LCase$(.List(.ListIndex))
    VineIndex = FindExactMatch(Me.lstVine, T)
    Me.cmdVine.Visible = VineIndex <> -1
    Me.cmdVine.Enabled = Me.cmdVine.Visible
    Me.cmdVine.ToolTipText = "View Vine reference for """ & .List(.ListIndex) & """"
    lstWordsClicked = False                       'turn off redundancy protection
  End With
  Me.rtbUser.TextRTF = S                          'reset the user text
  LockWindowUpdate 0                              'refresh updates allowed
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : lstWords_MouseMove
' Purpose           : When moving over word in the listbox, set the tooltip for the
'                   : listbox to the word the mouse is over if the text is longer than
'                   : the listbox
'*******************************************************************************
Private Sub lstWords_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String, G As String, Tip As String
  Static LastTip As String
  
  S = GetStringFromMouseMove(Me.lstWords, X, Y)     'get the text under the cursor
  If S <> LastTip Then
    LastTip = S                                     'different item
    If Len(S) > 0 Then                              'if it contains data...
      G = Me.lstGrkWords.List(Me.lstGrkWords.ListIndex)
      If Right$(G, 1) = "V" Then G = Left$(G, Len(G) - 1) & "s"
      Set Me.lblwidth.Font = Me.lstGrkWords.Font    'match the label font to the list
      Me.lblwidth.Caption = S                       'set the label text
      If Me.lblwidth.Width > Me.lstWords.Width Then 'is the label wider than the list?
        S = "'" & S & "'. "                         'expose word and tip
      Else
        S = vbNullString                            'else just tip
      End If
      If LastTip = Me.lstWords.List(Me.lstWords.ListIndex) Then
        Tip = "S"
      Else
        Tip = "Click to select this s"
      End If
      S = S & Tip & "ynonym for the Greek word """ & G & """. Also hold CTRL to change all other instances of this word in the verse" 'Add tip
    End If
    MyToolTips.ToolText(Me.lstWords) = S            'stuff tootip
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Changefont
' Purpose           : Change the font's point size
'*******************************************************************************
Private Sub Changefont(ByVal Index As Long)
  Dim Idx As Long, Sstart As Long, sLen As Long, Lidx As Long
  
  SaveSetting App.Title, "settings", "FontSize", CStr(Index)  'save new setting
  Lidx = Me.lstGrkWords.ListIndex
  Me.mnuFont8.Checked = False                                 'reset all checks
  Me.mnuFont10.Checked = False
  Me.mnuFont12.Checked = False
  Me.mnuFont14.Checked = False
  Me.mnuFont16.Checked = False
  Select Case Index
    Case 0
      FntSize = 10                   '10 point
      Me.mnuFont10.Checked = True
    Case 2
      FntSize = 14                   '14 point
      Me.mnuFont14.Checked = True
    Case 3
      FntSize = 16                   '16 point
      Me.mnuFont16.Checked = True
    Case 4
      FntSize = 8                    '8 point
      Me.mnuFont8.Checked = True
    Case Else
      FntSize = 12                   '(default) 12 point
      Me.mnuFont12.Checked = True
  End Select
  
  Me.lstGrkWords.FontSize = FntSize  'set to listboxes
  Me.lstWords.FontSize = FntSize
  
  With Me.lblwidth
    .FontSize = FntSize
    .Caption = "    (b) "
    HIndent = .Width
  End With
  
  With Me.rtbGreek                    'update rich text boxes
    LockWindowUpdate .hwnd            'disable screen updates for this text box
    Sstart = .SelStart                'save the selection info
    sLen = .SelLength
    .SelStart = 0
    .SelLength = Len(.Text)           'select everything
    .SelFontSize = FntSize            'set the font size
    .SelStart = Sstart                'reset selection points
    .SelLength = sLen
    LockWindowUpdate 0                're-enable screen updates
  End With
  With Me.rtbVerse
    LockWindowUpdate .hwnd
    Sstart = .SelStart
    sLen = .SelLength
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelStart = Sstart
    .SelLength = sLen
    LockWindowUpdate 0
  End With
  With Me.rtbNotes
    LockWindowUpdate .hwnd
    Sstart = .SelStart
    sLen = .SelLength
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelStart = Sstart
    .SelLength = sLen
    LockWindowUpdate 0
  End With
  With Me.rtbTranslate
    LockWindowUpdate .hwnd
    Sstart = .SelStart
    sLen = .SelLength
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelStart = Sstart
    .SelLength = sLen
    LockWindowUpdate 0
  End With
  With Me.rtbUser
    LockWindowUpdate 0
    Sstart = .SelStart
    sLen = .SelLength
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelStart = Sstart
    .SelLength = sLen
    LockWindowUpdate 0
  End With
  With Me.rtbVerseNotes
    LockWindowUpdate .hwnd
    Sstart = .SelStart
    sLen = .SelLength
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelStart = Sstart
    .SelLength = sLen
    LockWindowUpdate 0
  End With
  With Me.rtbMerge
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelStart = 0
  End With
  
  If Not IsLoading Then
    Me.lstWords.Height = Me.cmdAdd.Top - Me.lblSynonyms.Height
    Me.lstGrkWords.Height = Me.picGreekControl.Top - Me.lblGrkWords.Height
    Call UpdateVerse
    Me.lstGrkWords.ListIndex = Lidx
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileCreate_Click
' Purpose           : Create a new Personal version of the Bible
'*******************************************************************************
Private Sub mnuFileCreate_Click()
  Dim S As String, Bbl() As String, SS As String
  Dim Idx As Long
'
' if a personal version already exists, first ask them if they want to over-write it
'
  If PersonalVersion Then
    Select Case MessageBox(Me, "A Personal Version File Already Exists? OVerwrite it?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Confirm Overwrite")
      Case vbNo
        Exit Sub
    End Select
  End If
'
' prompt for version to base this new bible upon
'
  frmPersonalVersion.Show vbModal, Me
  If PersonalVersionBase < 0 Then Exit Sub    'user had cancelled
  Screen.MousePointer = vbHourglass
  DoEvents
'
' read in base bible
'
  Select Case PersonalVersionBase
    Case 1
      S = "YLT"
      SS = "Young's Literal Translation"
    Case 2
      S = "RSV"
      SS = "Revised Standard Version"
    Case 4
      S = "MKJV"
      SS = "Modern King James Version"
    Case 5
      S = "WEB"
      SS = "World English Bible"
    Case 6
      S = "ASV"
      SS = "American Standard Version"
    Case 7
      S = "DBY"
      SS = "Darby's Translation"
    Case 8
      S = "WBS"
      SS = "Webster's Translation"
    Case Else
      S = "KJV"
      SS = "King James Version"
  End Select
'
' read master template
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\" & S & ".txt", ForReading, False)
  Bbl = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' Write MPV.txt
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\MPV.txt", ForWriting, True)
  ts.Write Join(Bbl, vbCrLf)
  ts.Close
  
  PersonalVersion = True                          'tag personal version available
  Me.mnuBibleMPV.Enabled = True                   'enable MPV in menu
  Me.cmdMPV.Enabled = True                        'enable MPV button
  Me.lblEditPersonal.Visible = True
  MPVAvail = True                                 'MPV now available
  Me.cmdMPV.Value = True                          'select it
  
  If BblVersion = UserPVer Then
    Me.cmdMPV.Enabled = False
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\" & S & ".txt", ForReading, False)
    Bible = Split(ts.ReadAll, vbCrLf)
    ts.Close
    PVDirty = False
    Call UpdateVerse
  End If
  Screen.MousePointer = vbDefault
  
  MessageBox Me, "Your personalizable version of the New Covenant is active and ready to use." & vbCrLf & _
                 "It is based upon the '" & SS & "' and presently contains its text" & vbCrLf & _
                 "(except for verses that contain no original text to validate them)." & vbCrLf & vbCrLf & _
                 "Use the edit frame and options in the lower left panel to edit and update" & vbCrLf & _
                 "your personal version (MPV = My Personal Version).", vbOKOnly Or vbInformation, "Personal Version Ready"
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileGoto_Click
' Purpose           : Go to a user-select book, chapter, and verse
'*******************************************************************************
Private Sub mnuFileGoto_Click()
  Dim S As String, Ary() As String, T As String
  Dim Idx As Long, B As Long, C As Long, V As Long
  Dim hB As Long, hC As Long, hV As Long, hCC As Long, hVC As Long
  
  S = vbNullString
  T = vbNullString
  For B = 1 To 27
    Ary = Split(Books(B), ",")
    S = S & ", " & Ary(1)
  Next B
  
  Ary = Split(Books(Bk), ",")
  T = Ary(1) & " " & CStr(Chp) & ":" & CStr(Vrs)
  S = "Select Book C:V (ie, rev 13:18):" & vbCrLf & "(" & Mid$(S, 3) & ")"
  S = UCase$(Trim$(InputMsgBox(Me, S, "Go to Book Chapter:verse", T)))
  If Len(S) = 0 Then Exit Sub             'cancel
  
  Idx = InStr(1, S, " ")        'find separator between book and chapter
  If Idx > 0 Then
    T = Left$(S, Idx - 1)       'get book
    B = InStr(1, T, " ")        'check for book titles, like 1 PETER, and combine them (1PETER)
    If B <> 0 Then T = Left$(T, B - 1) & Mid$(T, B + 1)
    T = Left$(T, 3)             'maintain only leftmost 3 letters
    
    For B = 1 To 27               'check for something like 1PE and change to 1PT
      Ary = Split(Books(B), ",")
      If Left$(Ary(2), 3) = T Then
        T = Left$(Ary(1), 3)
        Exit For
      End If
    Next B
  End If
  
  For B = 1 To 27               'now scan for specified book
    Ary = Split(Books(B), ",")
    If Ary(1) = T Then Exit For
  Next B
  If B = 28 Then
   MessageBox Me, "Could not find specified book in: " & S, vbOKOnly Or vbExclamation, "Book Not Found"
    Exit Sub
  End If
  
  S = Trim$(Mid$(S, Idx + 1)) 'strip off book
  V = InStr(1, S, ":")    'get chapter and verse separator
  If V = 0 Then           'cannot find it
   MessageBox Me, "Bad chapter:verse specification: " & S, vbOKOnly Or vbExclamation, "Bad Specification"
    Exit Sub
  End If
  On Error Resume Next
  C = CLng(Left$(S, V - 1))
  V = CLng(Mid$(S, V + 1))
  Idx = CLng(Ary(4))
  If C < 1 Or C > Idx Then    'check for chapter being out of range
   MessageBox Me, "Chapter number " & CStr(C) & " is out of range (1 to " & CStr(Idx) & ")", vbOKOnly Or vbExclamation, "Chapter Invalid"
    Exit Sub
  End If
  If V > 0 Then
    hB = Bk
    hC = Chp
    hV = Vrs
    hCC = ChpCnt
    hVC = VrsCnt
    Bk = B
    Chp = C
    Vrs = V
    ChpCnt = Idx
    Call GetVerseCount
  End If
  If Vrs < 1 Or Vrs > VrsCnt Then
   MessageBox Me, "Verse number " & CStr(Vrs) & " is out of range (1 to " & CStr(VrsCnt) & ")", vbOKOnly Or vbExclamation, "Verse Invalid"
    Bk = hB
    Chp = hC
    Vrs = hV
    ChpCnt = hCC
    VrsCnt = hVC
  Else
    Call UpdateVerse    'all ok, so display the user selection
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileSave_Click
' Purpose           : Save the current data to the database files
'*******************************************************************************
Private Sub mnuFileSave_Click()
  Dim Ary() As String, S As String
  Dim Idx As Long, I As Long
  Dim TmrFlag As Boolean
  
  TmrFlag = Me.tmrAutoBackup.Enabled    'save timer flag
  Me.tmrAutoBackup.Enabled = False      'disable timer for now
  
  Screen.MousePointer = vbHourglass     'show that we are busy
  Me.Enabled = False
  DoEvents
'
' save book, chapter and verse last read
'
  Call SaveSetting(App.Title, "Settings", "Book", CStr(Bk))
  Call SaveSetting(App.Title, "Settings", "Chapter", CStr(Chp))
  Call SaveSetting(App.Title, "Settings", "Verse", CStr(Vrs))
'
' update word references
'
  If Not MakeVirgin Then
    For Idx = 1 To UBound(BBLWIdx)
      If Len(WordRef(Idx)) <> 0 Then
        Ary = Split(WordRef(Idx), vbTab)
        Ary(1) = CStr(BBLWIdx(Idx))
        WordRef(Idx) = Join(Ary, vbTab)
      End If
    Next Idx
  End If
'
' ALWAYS save word reference table (constantly updated)
'
  Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\GreekWordRef.txt", ForWriting, True)
  ts.Write Join(WordRef, vbCrLf)
  ts.Close
'
' and word map (user-selections for syninyms of words)
'
  If MakeVirgin Then
    If Fso.FileExists(AddSlash(AppPath) & "DB\WordMap.txt") Then
      On Error Resume Next
      Fso.DeleteFile AddSlash(AppPath) & "DB\WordMap.txt", True
      On Error GoTo 0
    End If
  Else
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\WordMap.txt", ForWriting, True)
    ts.Write Join(WordMap, vbCrLf)
    ts.Close
  End If
'
' if the Bible index file has had corrections applied, update the file
'
  If BBLDirty Then      'prevent CD-ROM update
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\GreekBBL.txt", ForWriting, False)
    ts.Write Join(GrkBBL, vbCrLf)
    ts.Close
    BBLDirty = False
  End If
'
' save user personal bible if it had any updates done to it
'
  If PVDirty Then
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\MPV.txt", ForWriting, True)
    ts.Write Join(Bible, vbCrLf)
    ts.Close
    PVDirty = False
  End If
'
' save user personal notes if it had any updates done to it
'
  If MyNotesDirty Then
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\MyNotes.txt", ForWriting, True)
    ts.Write Join(MyNotes, vbCrLf)
    ts.Close
    MyNotesDirty = False
  End If
'
' save viewing history
'
  S = AddSlash(AppPath) & "DB\History.txt"
  If CInt(GetSetting(App.Title, "Settings", "SaveHistory", "1")) = vbChecked Then
    Set ts = Fso.OpenTextFile(S, ForWriting, True)
    With colHistory
      I = .Count - 1000
      If I < 1 Then I = 1
      For Idx = I To .Count
        ts.WriteLine .Item(Idx)
      Next Idx
      ts.Close
    End With
  Else
    If Fso.FileExists(S) Then Call Fso.DeleteFile(S, True)
  End If
  Screen.MousePointer = vbDefault 'no longer busy
  Me.Enabled = True
  AutoDirty = False               'turn off the dirty flag
  Me.tmrAutoBackup.Enabled = TmrFlag
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileSearch_Click
' Purpose           : Launch Search (F5) from the Menu
'*******************************************************************************
Private Sub mnuFileSearch_Click()
  If SearchOpen Then
    With frmSearch
      If .WindowState = vbMinimized Then .WindowState = vbNormal
      DoEvents
      .ZOrder 0
    End With
  Else
    frmSearch.Show vbModeless, Me
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileSort_Click
' Purpose           : Sort all synonyms and adjust word mapping
'*******************************************************************************
Private Sub mnuFileSort_Click()
  Dim S As String, T As String, TT As String
  Dim wrAry() As String, wAry() As String, BBLMap() As String
  Dim wIdx As Long, I As Long, J As Long, K As Long, UB As Long
  Dim ChangeCount As Long, TmrFlag As Boolean
  
  If Me.cmdBack.Visible Then Me.cmdBack.Value = True
  TmrFlag = Me.tmrAutoBackup.Enabled
  Me.tmrAutoBackup.Enabled = False
  Screen.MousePointer = vbHourglass       'show that we are busy
  Me.Enabled = False
  Me.pgrWrite.Left = 0                    'set up the progress bar
  Me.pgrWrite.Top = Me.Toolbar1.Height
  Me.pgrWrite.Width = Me.cmdFindInText.Left
  Me.pgrWrite.Max = UBound(WordRef)
  Me.pgrWrite.Value = 0
  Me.pgrWrite.Visible = True
  DoEvents                                'let screen catch up
'
' scan all words in word list
'
  For wIdx = 1 To UBound(WordRef)
    Me.pgrWrite.Value = wIdx              'show progress
    S = WordRef(wIdx)                     'grab a word list entry
    If Len(S) <> 0 And Right$(S, 1) <> vbTab Then
      wrAry = Split(WordRef(wIdx), vbTab) 'break into any array
      wAry = Split(wrAry(2), ",")         'grab the synonyms
      UB = UBound(wAry)                   'get the upper bounds count
      If UB > 0 Then                      'anything to sort?
        With Me.lstSort
          .Clear                          'init the sorted list
          For I = 0 To UB
            .AddItem wAry(I)              'add word to sorted list
          Next I
          '
          ' now find the new index of each word in the list
          '
          T = vbNullString                'init re-accumulator
          For I = 0 To UB
            S = .List(I)                  'grab a sorted word
            T = T & "," & S               'add to accumulator
            '
            ' find the match in the old list, and replace the word with the new index
            '
            For J = 0 To UB
              If S = wAry(J) Then         'found a match?
                wAry(J) = CStr(I)         'yes, so stuff new index
                Exit For                  'done with this portion of the scan
              End If
            Next J
          Next I
          If wrAry(2) <> Mid$(T, 2) Then
            ChangeCount = ChangeCount + 1   'count a word to do
            wrAry(2) = Mid$(T, 2)           'stuff new list (less initial comma)
            wrAry(1) = wAry(CLng(wrAry(1))) 'stuff new "default" index
            BBLWIdx(wIdx) = CLng(wrAry(1))  'update local map
            WordRef(wIdx) = Join(wrAry, vbTab)  'update WordRef() entry
            
            T = vbTab & wrAry(0) & vbTab    'init search text
            For I = 1 To UBound(DefRef)     'found Strong # in DefRef() list
              If InStr(1, DefRef(I), T) <> 0 Then Exit For
            Next I
            J = InStr(1, DefRef(I), vbTab)  'find first tab
            T = Left$(DefRef(I), J - 1)     'grab search index
            
            TT = " " & T & " "                          'search mask
            For I = 0 To UBound(GrkBBL)
              S = GrkBBL(I)                             'grab a line
              If Len(S) <> 0 Then                       'if something there
                If InStr(1, S & " ", TT) <> 0 Then      'found DefRef index?
                  wrAry = Split(GrkBBL(I), " ")         'yes, itemize the table
                  BBLMap = Split(WordMap(I), " ")       'tab offset index table
                  For J = 1 To UBound(wrAry)
                    If T = wrAry(J) Then                'entry found?
                      If BBLMap(J) <> "-1" Then         'was it mapped? (-1 = no)
                        BBLMap(J) = wAry(CLng(BBLMap(J))) 'yes, so set to new offset index
                      End If
                    End If
                  Next J
                  WordMap(I) = Join(BBLMap, " ")        'put line back
                End If
              End If
            Next I                                      'process all entries
          End If
        End With
      End If
    End If
  Next wIdx                                           'do all words
  
  Me.pgrWrite.Visible = False       'hide progress bar
  Screen.MousePointer = vbDefault   'show no longer busy
  If ChangeCount <> 0 Then
    MessageBox Me, "All synonym lists have been sorted, and the word map" & vbCrLf & _
                           "database has been adjusted to retain integirty." & vbCrLf & _
                           CStr(ChangeCount) & " synonym lists updated.", _
                           vbOKOnly Or vbInformation, "Operation Completed"
    AutoDirty = True                'something has changed
  Else
    MessageBox Me, "All synonyms were already properly sorted." & vbCrLf & _
                           "Nothing needed to be done.", _
                           vbOKOnly Or vbInformation, "Operation Completed"
  End If
  Me.Enabled = True                 're-enable ourselves
  Call UpdateVerse                  'refresh the current verse
  Me.tmrAutoBackup.Enabled = TmrFlag
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileViewKJVDict_Click
' Purpose           : View the KJV dictionary
'*******************************************************************************
Private Sub mnuFileViewKJVDict_Click()
  If Not ShowKJVDict Then
    frmKJVDict.Show vbModeless, Me
  Else
    frmKJVDict.WindowState = vbNormal
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileViewSavedBible_Click
' Purpose           : View the currently saved bible using our in-program viewer
'*******************************************************************************
Private Sub mnuFileViewSavedBible_Click()
  If ViewBible Then
    With frmView
      If .WindowState = vbMinimized Then
        .WindowState = CLng(GetSetting(App.Title, "Settings", "BBLViewer", "0"))
      End If
      DoEvents
      .ZOrder 0
    End With
  Else
    frmView.Show vbModeless, Me
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileViewTheo_Click
' Purpose           : View ALl Theological Notes for the current chapter
'
' Every chapter in the New Covenant has notes, so we will not have to check first
' to see if notes exist for the chapter; it will be simply assumed.
'*******************************************************************************
Private Sub mnuFileViewTheo_Click()
  Dim Idx As Long, V As Long, I As Long, J As Long, Vcnt As Long, SS As Long
  Dim S As String, BkChp As String, sBook As String, Ary() As String, T As String
  Dim HaveShownOne As Boolean
  
  Ary = Split(Books(Bk), ",")     'get the title for the Book
  sBook = Ary(3)
  
  InitMerge
  AddMergeCrlf "Theological Notes for: " & sBook & ", Chapter " & CStr(Chp), "Arial", FntSize + 2, True, , rtfCenter
  
  sBook = Format$(Bk, "00") & Format$(Chp, "00")
  For V = 1 To VrsCnt
    sBook = Left$(sBook, 4) & Format$(V, "00")
    S = vbNullString
    I = FindExactMatch(Me.lstVNotes, sBook)           'find a match
    If I <> -1 Then                                   'found something...
      If HaveShownOne Then AddMergeCrlf Underline
      HaveShownOne = True
      Idx = FindExactMatch(Me.lstGrk, sBook)          'find verse index
      T = Mid$(Bible(Idx), 8)
      With Me.rtbMerge
        SS = Len(.Text)
        If Len(T) <> 0 Then
          AddMergeCrlf vbCrLf & "(" & CStr(V) & ") " & T, , FntSize, , True
        Else
          AddMerge vbCrLf & "Verse " & CStr(V) & ":", , FntSize - 2, True, True
        End If
        '
        ' make verses without Greek text obnoxious.
        '
        If Len(Grk(Idx)) < 8 Then
          .SelStart = SS
          .SelLength = Len(.Text) - SS
          .SelColor = vbMagenta
        End If
      End With
      J = I                                           'save last-found index
      Do While I <> -1
        S = S & vbCrLf & vbCrLf & "NOTE: " & Mid$(VNotes(I), 8)  'add a note"
        I = FindExactMatch(Me.lstVNotes, sBook, I)    'find another
        If J >= I Then Exit Do                        'ignore if index matches last
      Loop
    End If
    If Len(S) <> 0 Then
      With Me.rtbMerge
        SS = Len(.Text)
        processText Mid$(S, 4) & vbCrLf, , FntSize
        .SelStart = SS
        .SelLength = Len(.Text) - SS
        .SelColor = PNotesColor
        S = .Text
        Idx = InStr(1, S, "NOTE: ")
        Do While Idx <> 0
          .SelStart = Idx - 1
          .SelLength = 5
          .SelBold = True
          .SelItalic = True
          .SelUnderline = True
          Idx = InStr(Idx + 6, S, "NOTE: ")
        Loop
      End With
    End If
  Next V
'
' now display the accumulated data
'
  LockWindowUpdate Me.rtbNotes.hwnd                   'lock display updates for window
  SetIndent
  Me.rtbNotes.Text = vbNullString                     'reset scrolling
  Me.rtbNotes.TextRTF = Me.rtbMerge.TextRTF           'stuff new text
  Me.rtbNotes.BackColor = vbInfoBackground            'tooltip background color
  LockWindowUpdate 0                                  'allow refresh
  InitMerge                                           'clear merge text
  Me.lblVineRef.Visible = False
  Me.lblNoteInfo.Visible = False
  Me.cmdAnalysis.Enabled = True
  Me.mnuFileAnalysis.Enabled = True
  Me.cmdCopyDef.Enabled = True
  Me.cmdBack.Visible = True
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileWordpad_Click
' Purpose           : Launch the Wordpad text editor application
'*******************************************************************************
Private Sub mnuFileWordpad_Click()
  Shell WordPadPath, vbNormalFocus
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileNotepad_Click
' Purpose           : Launch the Notepad text editor application
'*******************************************************************************
Private Sub mnuFileNotepad_Click()
  Shell NotePadPath, vbNormalFocus
End Sub

'*******************************************************************************
' set font point sizes
'*******************************************************************************
Private Sub mnuFont10_Click()
  Changefont 0
End Sub

Private Sub mnuFont12_Click()
  Changefont 1
End Sub

Private Sub mnuFont14_Click()
  Changefont 2
End Sub

Private Sub mnuFont16_Click()
  Changefont 3
End Sub

Private Sub mnuFont8_Click()
  Changefont 4
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHlpAbout_Click
' Purpose           : Display the ABOUT box
'*******************************************************************************
Private Sub mnuHlpAbout_Click()
  On Error Resume Next
  frmSplash.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHlpCC_Click
' Purpose           : View a crash course in Greek
'*******************************************************************************
Private Sub mnuHlpCC_Click()
  OpenFilePath Me.hwnd, AddSlash(App.Path) & "DB\CrashCourse.htm"
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHLPKeyboard_Click
' Purpose           : Keyboard navigation help
'*******************************************************************************
Private Sub mnuHLPKeyboard_Click()
  frmKeyboardHelp.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHlpMascFem_Click
' Purpose           : SHow some info about how gender of words is checked
'*******************************************************************************
Private Sub mnuHlpMascFem_Click()
  frmAboutGender.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHlpPlurality_Click
' Purpose           : Show some help regarding the Plurality testing
'*******************************************************************************
Private Sub mnuHlpPlurality_Click()
  frmAboutPlurality.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHLPQR_Click
' Purpose           : View Greek text quick reference tables
'*******************************************************************************
Private Sub mnuHLPQR_Click()
  OpenFilePath Me.hwnd, AddSlash(App.Path) & "DB\QRef.htm"
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHLPTipoftheDay_Click
' Purpose           : Users want to view tip of the daya
'*******************************************************************************
Private Sub mnuHLPTipoftheDay_Click()
  SaveSetting App.Title, "Settings", "Show Tips at Startup", "1"
  frmTip.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHLPUsing_Click
' Purpose           : View the main help file
'*******************************************************************************
Public Sub mnuHLPUsing_Click()
  OpenFilePath Me.hwnd, AddSlash(App.Path) & "DB\UsingNCX.htm"
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHLPViewDemo_Click
' Purpose           : View a demo of the application
'*******************************************************************************
Private Sub mnuHLPViewDemo_Click()
  frmViewDemo.Show vbModeless, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHLPViewReadme_Click
' Purpose           : View the application ReadMe.txt file
'*******************************************************************************
Private Sub mnuHLPViewReadme_Click()
  OpenFilePath Me.hwnd, AddSlash(App.Path) & "Readme.txt"
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHLPVisitSponsor_Click
' Purpose           : Browse to a sponsor website
'*******************************************************************************
Public Sub mnuHLPVisitSponsor_Click()
'
' check for an internet connection
'
  If Not CheckInternetConnect Then  'if none...
    MessageBox Me, "Internet Connection Not Detected. View the book details at:" & vbCrLf & _
    "www.authorhouse.com/BookStore/ItemDetail~bookid~33204.aspx", vbOKOnly Or vbInformation, "No Internet Connnection"
    Exit Sub
  End If
'
' insert your sponsoring site here
'
  OpenFilePath Me.hwnd, "http://www.authorhouse.com/BookStore/ItemDetail~bookid~33204.aspx"
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpBible_Click
' Purpose           : Search Bible for a reference to the selected text
'*******************************************************************************
Private Sub mnuPopUpBible_Click()
  ForceSearch = Me.rtbNotes.SelText
  frmSearch.Show vbModeless, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpCopy_Click
' Purpose           : Save selection to the clipboard
'*******************************************************************************
Private Sub mnuPopUpCopy_Click()
  Clipboard.Clear
  Clipboard.SetText Me.rtbNotes.SelText, vbCFText
  Clipboard.SetText Me.rtbNotes.SelRTF, vbCFRTF
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpCopy2_Click
' Purpose           : Save selection to the clipboard
'*******************************************************************************
Private Sub mnuPopUpCopy2_Click()
  Clipboard.Clear
  Clipboard.SetText Me.rtbVerse.SelText, vbCFText
  Clipboard.SetText Me.rtbVerse.SelRTF, vbCFRTF
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpCopy3_Click
' Purpose           : Save selection to the clipboard
'*******************************************************************************
Private Sub mnuPopUpCopy3_Click()
  Clipboard.Clear
  Clipboard.SetText Me.rtbVerseNotes.SelText, vbCFText
  Clipboard.SetText Me.rtbVerseNotes.SelRTF, vbCFRTF
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopupVine_Click
' Purpose           : Search for a vine reference for the selected text
'*******************************************************************************
Private Sub mnuPopupVine_Click()
  Call rtbNotes_DblClick
End Sub

'*******************************************************************************
' Subroutine Name   : mnuViewReset_Click
' Purpose           : Reset frames to default positions
'*******************************************************************************
Private Sub mnuViewReset_Click()
  Me.lstGrkWords.Width = orgLstGrkWordsWidth
  Me.lstWords.Width = orgLstWordsWidth
  Me.picTree.Width = orgTvWidth
  Call ForceResize
  Me.Hide
  Me.Show
End Sub

'*******************************************************************************
' Subroutine Name   : mnuViewSyn_Click
' Purpose           : Show Synonym word list
'*******************************************************************************
Private Sub mnuViewSyn_Click()
  If Not ShowWordList Then
    Screen.MousePointer = vbHourglass
    DoEvents
    frmShowWords.Show vbModeless, Me
  Else
    frmShowWords.WindowState = vbNormal
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuViewVine_Click
' Purpose           : Show Vine word list
'*******************************************************************************
Private Sub mnuViewVine_Click()
  If Not ShowVineList Then
    Screen.MousePointer = vbHourglass
    DoEvents
    frmShowVineList.Show vbModeless, Me
  Else
    With frmShowVineList
      If .WindowState = vbMinimized Then .WindowState = vbNormal
      DoEvents
      .ZOrder 0
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuWinBible_Click
' Purpose           : Activate Bible reader if it is available
'*******************************************************************************
Private Sub mnuWinBible_Click()
  With frmView
    If .WindowState = vbMinimized Then
      .WindowState = CLng(GetSetting(App.Title, "Settings", "BBLViewer", "0"))
    End If
    DoEvents
    .ZOrder 0
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuWinKJV_Click
' Purpose           : Activate KJV dictionary if it is available
'*******************************************************************************
Private Sub mnuWinKJV_Click()
  With frmKJVDict
    If .WindowState = vbMinimized Then .WindowState = vbNormal
    DoEvents
    .ZOrder 0
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuWinSearch_Click
' Purpose           : Activate search dialog if it is available
'*******************************************************************************
Private Sub mnuWinSearch_Click()
  With frmSearch
    If .WindowState = vbMinimized Then .WindowState = vbNormal
    DoEvents
    .ZOrder 0
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuWinStrong_Click
' Purpose           : Activate Strong # viewer if it is available
'*******************************************************************************
Private Sub mnuWinStrong_Click()
  With frmShowWords
    If .WindowState = vbMinimized Then .WindowState = vbNormal
    DoEvents
    .ZOrder 0
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuWinVine_Click
' Purpose           : Activate Vine reference viewer if it is available
'*******************************************************************************
Private Sub mnuWinVine_Click()
  With frmShowVineList
    If .WindowState = vbMinimized Then .WindowState = vbNormal
    DoEvents
    .ZOrder 0
  End With
End Sub

'*******************************************************************************
' picture background updates for parchement data
'*******************************************************************************
Private Sub picTop_Paint()
  PaintTilePicBackground Me.picTop, Me.picTile(Background)
End Sub

Private Sub picVbar3_Paint()
  If vBarDown3 Then Exit Sub
  PaintTilePicBackground Me.picVbar3, Me.picTile(Background)  'background
  Vbarz Me.picVbar3                                           '3D border
End Sub

Private Sub picVbar4_Paint()
  If vBarDown4 Then Exit Sub
  PaintTilePicBackground Me.picVbar4, Me.picTile(Background)  'background
  Vbarz Me.picVbar4                                           '3D border
End Sub

Private Sub picVerse_Paint()
  PaintTilePicBackground Me.picVerse, Me.picTile(Background)
End Sub

Private Sub picNotes_Paint()
  PaintTilePicBackground Me.picNotes, Me.picTile(Background)
End Sub

Private Sub picEditor_Paint()
  PaintTilePicBackground Me.picEditor, Me.picTile(Background)
End Sub

Private Sub picGreek_Paint()
  PaintTilePicBackground Me.picGreek, Me.picTile(Background)
End Sub

Private Sub picGreekControl_Paint()
  PaintTilePicBackground Me.picGreekControl, Me.picTile(Background)
End Sub

Private Sub picHbar1_Paint()
  If hBarDown1 Then Exit Sub
  PaintTilePicBackground Me.picHbar1, Me.picTile(Background)
End Sub

Private Sub picVbar1_Paint()
  PaintTilePicBackground Me.picVbar1, Me.picTile(Background)
End Sub

Private Sub picVbar2_Paint()
  If vBarDown2 Then Exit Sub
  PaintTilePicBackground Me.picVbar2, Me.picTile(Background)
End Sub

'*******************************************************************************
' norizontal center bar
'*******************************************************************************
Private Sub picHbar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 0 Then
    hBarDown1 = True
    Me.picHbar1.BackColor = cdGray
  End If
End Sub

Private Sub picHbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Tip As Long, Tp As Long
  
  If Not hBarDown1 Then Exit Sub
  With Me.picHbar1
    Tip = .Top - (.Height \ 2 - CLng(Y))
    Tp = Me.ScaleHeight - 4320
    If (Tip + .Height) > Tp Then Tip = Tp
    Tp = 4320
    If Tip < Tp Then Tip = Tp
    If .Top <> Tip Then .Top = Tip
  End With
End Sub

Private Sub picHbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.picHbar1.BackColor = cLight
  hBarDown1 = False
  GreekHeight = Me.picHbar1.Top - Me.picGreek.Top
  Call ForceResize
  Me.picHbar1.Refresh
End Sub

Private Sub picHbar1_DblClick()
  GreekHeight = orgPicGreekHeight
  Call ForceResize
  Me.picHbar1.Refresh
End Sub

'*******************************************************************************
' Vertical left bar
'*******************************************************************************
Private Sub picVbar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 0 Then
    vBarDown1 = True
    Me.picVbar1.BackColor = cdGray
  End If
End Sub

Private Sub picVbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Lft As Long, Lf As Long

  If Not vBarDown1 Then Exit Sub
  With Me.picVbar1
    Lft = .Left - (.Width \ 2 - CLng(X))
    Lf = 2880
    If (Lft + .Width) > Lf Then Lft = Lf
    Lf = 1440
    If Lft < Lf Then Lft = Lf
    If .Left <> Lft Then .Left = Lft
  End With
End Sub

Private Sub picVbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.picVbar1.BackColor = cLight
  vBarDown1 = False
  Call ForceResize
  Me.picVbar1.Refresh
End Sub

Private Sub picVbar1_DblClick()
  Me.picVbar1.Left = orgTvWidth - Me.picTree.Left
  Call ForceResize
  Me.picVbar2.Refresh
End Sub

'*******************************************************************************
' Vertical center bar
'*******************************************************************************
Private Sub picVbar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 0 Then
    vBarDown2 = True
    Me.picVbar2.BackColor = cdGray
  End If
End Sub

Private Sub picVbar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Lft As Long, Lf As Long
  
  If Not vBarDown2 Then Exit Sub
  With Me.picVbar2
    Lft = .Left - (.Width \ 2 - CLng(X))
    Lf = Me.picVerse.Left + Me.cmdMKJV.Left
    If (Lft + .Width) > Lf Then Lft = Lf
    Lf = Me.picEditor.Left + Me.cmdCopympv.Left + Me.cmdCopympv.Width
    If Lft < Lf Then Lft = Lf
    If .Left <> Lft Then .Left = Lft
  End With
End Sub

Private Sub picVbar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.picVbar2.BackColor = cLight
  vBarDown2 = False
  GreekWidth = Me.picVbar2.Left - Me.picGreek.Left
  Call ForceResize
  Me.picVbar2.Refresh
End Sub

Private Sub picVbar2_DblClick()
  GreekWidth = orgPicGreekWidth
  Call ForceResize
  Me.picVbar2.Refresh
End Sub

'*******************************************************************************
' Internal Greek data resizer control bar
'*******************************************************************************
Private Sub picVbar3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 0 Then
    vBarDown3 = True
    Me.picVbar3.BackColor = cdGray
  End If
End Sub

Private Sub picVbar3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Lft As Long, Lf As Long
  
  If Not vBarDown3 Then Exit Sub
  With Me.picVbar3
    Lft = .Left - (.Width \ 2 - CLng(X))
    Lf = Me.cmdEdit.Left - Me.cmdEdit.Width
    If (Lft + .Width) > Lf Then Lft = Lf
    Lf = Me.cmdCpyXlt.Left + Me.cmdCpyXlt.Width
    If Lf > Lft Then Lft = Lf
    If .Left <> Lft Then .Left = Lft
  End With
End Sub

Private Sub picVbar3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.picVbar3.BackColor = cLight
  vBarDown3 = False
  Me.lstGrkWords.Width = Me.picGreek.ScaleWidth - (Me.picVbar3.Left + Me.picVbar3.Width)
  Call ForceResize
  Me.picVbar3.Refresh
End Sub

Private Sub picVbar3_DblClick()
  If Me.picGreek.ScaleWidth - orgLstGrkWordsWidth - 60 > 0 Then
    Me.lstGrkWords.Width = orgLstGrkWordsWidth
    Call ForceResize
    Me.picVbar3.Refresh
  End If
End Sub

'*******************************************************************************
' Internal editbox resizer control bar
'*******************************************************************************
Private Sub picVbar4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 0 Then
    vBarDown4 = True
    Me.picVbar4.BackColor = cdGray
  End If
End Sub

Private Sub picVbar4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Lft As Long, Lf As Long
  
  If Not vBarDown4 Then Exit Sub
  With Me.picVbar4
    Lft = .Left - (.Width \ 2 - CLng(X))
    Lf = Me.cmdEdit.Left - Me.cmdEdit.Width
    If (Lft + .Width) > Lf Then Lft = Lf
    Lf = Me.cmdCpyXlt.Left + Me.cmdCpyXlt.Width
    If Lf > Lft Then Lft = Lf
    If .Left <> Lft Then .Left = Lft
  End With
End Sub

Private Sub picVbar4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.picVbar4.BackColor = cLight
  vBarDown4 = False
  Me.lstWords.Width = Me.picEditor.ScaleWidth - (Me.picVbar4.Left + Me.picVbar4.Width)
  Call ForceResize
  Me.picVbar4.Refresh
End Sub

Private Sub picVbar4_DblClick()
  If Me.picEditor.ScaleWidth - orgLstWordsWidth - 60 > 0 Then
    Me.lstWords.Width = orgLstWordsWidth
    Call ForceResize
    Me.picVbar4.Refresh
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : rtbGreek_MouseMove
' Purpose           : de/activate a tooltip, based upon mouse location
'*******************************************************************************
Private Sub rtbGreek_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String
  Static LastTip As String
  
  If Y < 360! Then
    S = "Double-click the Top line here to view the verse in context"
  Else
    S = vbNullString
  End If
  If LastTip = S Then Exit Sub
  LastTip = S
  MyToolTips.ToolText(Me.rtbGreek) = S
End Sub

'*******************************************************************************
' Subroutine Name   : rtbNotes_DblClick
' Purpose           : See if we can Extract a Vine reference from the data
'*******************************************************************************
Private Sub rtbNotes_DblClick()
  Dim S As String, T As String, Ary() As String
  Dim I As Long, J As Long, K As Long
  
  If CheckUndl(Me.rtbNotes) Then Exit Sub       'user selected underlined text
  Screen.MousePointer = vbHourglass             'show that we are busy
  Me.Enabled = False
  DoEvents
  With Me.rtbNotes
    T = .Text                                   'grab the data
    I = .SelStart + 1                           'set the cursor point
    J = InStrRev(T, " ", I)                     'find a leading space
    K = InStrRev(T, vbLf, I)
    If J = 0 Then J = 1                         'none
    If K > J Then J = K
    K = InStr(I, T, " ")                        'find a terminating space
    If K = 0 Then K = Len(T) + 1
    On Error Resume Next
    S = Trim$(Mid$(T, J + 1, K - J - 1))        'grab data
    If Err.Number <> 0 Then                     'may have clicked on a space
      S = vbNullString
    End If
    On Error GoTo 0
    If Len(S) = 0 Then Exit Sub                 'nothing to do
'
' check for verse number selection
'
  If Left$(S, 1) = "(" And Right$(S, 1) = ")" Then
    T = Mid$(S, 2, Len(S) - 2)
    If IsNumeric(T) Then
      Vrs = CLng(T)
      Call UpdateVerse
      NoGreekSupport Format$(Bk, "00") & Format$(Chp, "00") & Format$(Vrs, "00")
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      Me.lstGrkWords.SetFocus
      Exit Sub
    End If
  End If
'
' construct a usable string
'
    T = vbNullString
    For I = 1 To Len(S)
      Select Case Mid$(S, I, 1)
        Case "'", "_", "#", "0" To "9", "A" To "Z", "a" To "z", " "
          T = T & Mid$(S, I, 1)
        Case Else
          T = T & " "
      End Select
    Next I
    
    T = Trim$(T)
    I = InStr(1, T, " ")
    If I <> 0 Then T = Left$(T, I - 1)          'remove intervening spaces
    If Len(T) = 0 Then Exit Sub                 'should never happen, but...
    If IsNumeric(Left$(T, 1)) Then              'a strong #? Assume so if numeric
      For I = 1 To UBound(DefRef) - 1
        S = DefRef(I)
        If Len(S) <> 0 Then
          Ary = Split(DefRef(I), vbTab)
          If T = Ary(5) Then                    'found Strong #?
            ShowDefRef CStr(I), False           'yes, so show the data
            ShowKJVCount CStr(I)
            ShowGrkWrdCnt CStr(I)
            Me.cmdVine.Enabled = True
            Me.cmdAnalysis.Enabled = True
            Me.mnuFileAnalysis.Enabled = True
            Me.cmdCopyDef.Enabled = True
            Me.cmdBack.Visible = True
            Me.rtbNotes.BackColor = clBlue
            Exit For                            'all done
          End If
        End If
      Next I
    ElseIf Left$(T, 1) <> "#" Then              'not a Vine reference #
      If .SelFontName = "Symbol" Then
        
        S = T & vbTab
        J = Len(S)
        For K = 0 To UBound(GrkWrdCnt) - 1
          If Left$(GrkWrdCnt(K), J) = S Then
            J = InStrRev(GrkWrdCnt(K), vbTab)
            I = CLng(Mid$(GrkWrdCnt(K), J + 1))
            ShowDefRef CStr(I), False
            ShowKJVCount CStr(I)
            ShowGrkWrdCnt CStr(I)
            Me.cmdVine.Enabled = True
            Me.cmdAnalysis.Enabled = True
            Me.mnuFileAnalysis.Enabled = True
            Me.cmdCopyDef.Enabled = True
            Me.cmdBack.Visible = True
            Me.rtbNotes.BackColor = clBlue
            Exit For
          End If
        Next K
      Else
        I = InStr(1, T, "_")
        Do While I <> 0
          Mid$(T, I, 1) = " "
          I = InStr(1, T, "_")
        Loop
        I = FindExactMatch(Me.lstVine, T)
        If I = -1 Then
          Screen.MousePointer = vbDefault
          MessageBox Me, "Cannot find a match in the Vine Database for the word '" & T & "'.", vbOKOnly Or vbInformation, "Word Not Found"
        Else
          DisplayVine I, T
          Me.cmdVine.Enabled = True
          Me.cmdAnalysis.Enabled = True
          Me.mnuFileAnalysis.Enabled = True
          Me.cmdCopyDef.Enabled = True
          Me.cmdBack.Visible = True
        End If
      End If
    End If
  End With
  
  DoEvents
  Me.rtbNotes.SelLength = 0
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : rtbNotes_MouseDown
' Purpose           : Display options if user does a right-click on the textbox when
'                   : data is selected within it
'*******************************************************************************
Private Sub rtbNotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton And Shift = 0 And Me.rtbNotes.SelLength <> 0 Then
    PopupMenu Me.mnuPopUp, vbPopupMenuRightButton, , , Me.mnuPopupVine
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : rtbTranslate_Click
' Purpose           : User clicked a word in the direct translation window
'*******************************************************************************
Private Sub rtbTranslate_Click()
  Dim S As String
  Dim I As Long, SL As Long, M As Long, Idx As Long
  
  If Me.lstGrkWords.ListCount = 0 Then Exit Sub
  With Me.rtbTranslate
    If CBool(.SelLength) Then Exit Sub
    S = Mid$(.Text, DTHeaderOffset)
    SL = .SelStart - DTHeaderOffset
'    If CBool(.SelLength) Then SL = SL + 1
    M = 1
    Idx = 0
    I = InStr(1, S, "]")
    Do While I < SL
      M = I + 1
      I = InStr(M + 2, S, "]")
      If I = 0 Then I = Len(S) + 1
      Idx = Idx + 1
    Loop
    If Not CBool(.SelLength) Then Me.lstGrkWords.ListIndex = Idx
    If M = 1 Then
      .SelStart = M + DTHeaderOffset - 1
      .SelLength = I - M - 1
    Else
      .SelStart = M + DTHeaderOffset
      .SelLength = I - M - 2
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : rtbTranslate_DblClick
' Purpose           : Echo a double-click in the text to the list
'*******************************************************************************
Private Sub rtbTranslate_DblClick()
  Call lstGrkWords_DblClick
End Sub

'*******************************************************************************
' Subroutine Name   : rtbTranslate_GotFocus
' Purpose           : Force focus back to the list
'*******************************************************************************
Private Sub rtbTranslate_GotFocus()
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : rtbUser_Change
' Purpose           : Monitor changes to the user-edit field
'*******************************************************************************
Private Sub rtbUser_Change()
  Me.cmdUpdateMPV.Enabled = CBool(Len(Trim$(Me.rtbUser.Text)))
End Sub

'*******************************************************************************
' Subroutine Name   : rtbVerse_MouseDown
' Purpose           : Display options if user does a right-click on the textbox when
'                   : data is selected within it
'*******************************************************************************
Private Sub rtbVerse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton And Shift = 0 And Me.rtbVerse.SelLength <> 0 Then
    PopupMenu Me.mnuPopUp2, vbPopupMenuRightButton
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : rtbVerseNotes_MouseDown
' Purpose           : Display options if user does a right-click on the textbox when
'                   : data is selected within it
'*******************************************************************************
Private Sub rtbVerseNotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton And Shift = 0 And Me.rtbVerseNotes.SelLength <> 0 Then
    PopupMenu Me.mnuPopUp3, vbPopupMenuRightButton
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : tmrAutoBackup_Timer
' Purpose           : Timer autobackups
'*******************************************************************************
Private Sub tmrAutoBackup_Timer()
  AutoTimeUpd = AutoTimeUpd - 1
  If AutoTimeUpd > 0 Then Exit Sub
'
' see if we should auto-save backup files
'
  If AutoDirty Then Call mnuFileSave_Click
  AutoTimeUpd = AutoTime          'reset the timer
End Sub

'*******************************************************************************
' Subroutine Name   : Toolbar1_ButtonClick
' Purpose           : Handle toolbar selection
'*******************************************************************************
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "save"
      mnuFileSave_Click
    Case "goto"
      mnuFileGoto_Click
    Case "backup"
      mnuFileBackup_Click
    Case "search"
      mnuFileSearch_Click
    Case "vine"
      mnuViewVine_Click
    Case "words"
      mnuViewSyn_Click
    Case "rebuild"
      mnuFileSort_Click
    Case "fav"
      mnuFavAdd_Click
    Case "keyboard"
      mnuHLPKeyboard_Click
    Case "help"
      mnuHLPUsing_Click
    Case "prevnote"
      mnuBBLFindPrev_Click
    Case "nextnote"
      mnuBBLFindNext_Click
    Case "prevtheo"
      mnuFileTheoPrev_Click
    Case "nexttheo"
      mnuFileTheoNext_Click
    Case "prevgreek"
      mnuBBLFindPrevNoGreek_Click
    Case "nextgreek"
      mnuBBLFindNextNoGreek_Click
    Case "kjvdict"
      mnuFileViewKJVDict_Click
    Case "wordpad"
      mnuFileWordpad_Click
    Case "notepad"
      mnuFileNotepad_Click
  End Select
End Sub

Private Sub tvBooks_Collapse(ByVal Node As ComctlLib.Node)
  Node.Image = 2
  Node.SelectedImage = 2
End Sub

Private Sub tvBooks_Expand(ByVal Node As ComctlLib.Node)
  Node.Image = 1
  Node.SelectedImage = 1
End Sub

Private Sub tvBooks_GotFocus()
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileExit_Click
' Purpose           : Exit program
'*******************************************************************************
Private Sub mnuFileExit_Click()
  Call mnuFileSave_Click
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : rtbGreek_Click
' Purpose           : User clicked a word in the Greek text box
'*******************************************************************************
Private Sub rtbGreek_Click()
  Dim I As Long, K As Long, SL As Long, Idx As Long, J As Long
  Dim S As String
  
  If Me.lstGrkWords.ListCount = 0 Then Exit Sub 'nothing to do
  S = Me.rtbGreek.Text                          'grab text
  J = InStr(1, S, vbCr) + 4                     'jump past header
  S = Mid$(S, J)
  If Len(S) = 0 Then Exit Sub
  With Me.rtbGreek
    If CBool(.SelLength) Then Exit Sub
    SL = .SelStart - J + 2              'get cursor position
    If SL < 0 Then Exit Sub             'ignore if clicking below actual text
    I = InStr(1, S, " ")
    K = 0
    Idx = 0
'
' figure the index of the selected word
'
    Do While I < SL
      K = I
      I = InStr(I + 1, S, " ")
      If I = 0 Then I = Len(S) + 1
      Idx = Idx + 1
    Loop
'
' select in the Greek word list
'
    If Not CBool(.SelLength) Then Me.lstGrkWords.ListIndex = Idx
    .SelStart = K + J - 1  'ensure selection set
    .SelLength = I - K - 1
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : rtbGreek_DblClick
' Purpose           : Double-click on Greek panel. If no greek text, bring up verse
'                   : in Chapter-view
'*******************************************************************************
Private Sub rtbGreek_DblClick()
  Dim S As String
  Dim J As Long
  
  If Bk = 0 Then Exit Sub                 'if no book, then do nothing
  With Me.rtbGreek
    S = .Text                             'grab text
    J = InStr(1, S, vbCr) + 4             'jump past header
    J = .SelStart - J + 2                 'get cursor position
    If J < 0 Then
      NoGreekSupport Format$(Bk, "00") & Format$(Chp, "00") & Format$(Vrs, "00")
      Exit Sub
    End If
  End With
  
  If Me.lstGrkWords.ListCount > 0 Then    'if a list of words exist
    Call lstGrkWords_DblClick             'handle the list as a double-click
    Exit Sub
  End If
End Sub

Private Sub rtbGreek_GotFocus()
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : tvBooks_MouseMove
' Purpose           : Do live updates on books and chapters on mouse moves
'*******************************************************************************
Private Sub tvBooks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Nd As Node
  Dim S As String
  Static LastTip As String
  
  Set Nd = Me.tvBooks.HitTest(X, Y)
  If Nd Is Nothing Then
    S = "Select New Covenant Books and Chapters"
  Else
    S = Nd.Tag
    If IsNumeric(Right$(S, 1)) Then
      S = S & ". Click to select it"
    Else
      If Nd.Expanded Then
        S = S & ". Click to hide this book's chapter list"
      Else
        S = S & ". Click to expose this book's chapter list"
      End If
    End If
  End If
  
  If S <> LastTip Then
    LastTip = S
    MyToolTips.ToolText(Me.tvBooks) = S
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : tvBooks_NodeClick
' Purpose           : User selected a chapter from the treeview
'*******************************************************************************
Private Sub tvBooks_NodeClick(ByVal Node As ComctlLib.Node)
  Dim S As String, Ary() As String
  Dim Idx As Long
'
' if it is a book (image id=1 or 2) then auto-expand/collapse on single-click
'
  If Node.Image <> 3 Then
    Node.Expanded = Not Node.Expanded
    Exit Sub                        'handle rest with Expanded and Collapse events
  End If
'
' handle chapter selections
'
  Chp = CLng(Node.Text)             'get the chapter number
  S = Node.Key                      'get the title for the chapter (BOOKxx)
  S = Left$(S, Len(S) - 2)          'strip the chapter number (01, 02, etc.)
  For Bk = 1 To 27
    Ary = Split(Books(Bk), ",")     'find the book we are playing with
    If S = Ary(1) Then Exit For
  Next Bk
  ChpCnt = CLng(Ary(4))             'get the chapter count
  Vrs = 1                           'always start at the first verse
  S = ShowVerse()                   'display verse information
  ChgSCroll = True                  'prevent redundancy
  If Len(S) <> 0 Then               'if valid data
    Call GetVerseCount              'get the number of verses
    S = "Verse 1 of " & CStr(VrsCnt)
  End If
  Me.lblVerseIndex.Caption = S
  Me.hsGreek.Value = 0              'init to scrollbar to verse-1
  ChgSCroll = False                 'turn off redundancy protection
  DoEvents                          'let updates catch up
  Me.lstGrkWords.SetFocus
End Sub

'*******************************************************************************
' Function Name     : ShowVerse
' Purpose           : Display Verse data
'*******************************************************************************
Private Function ShowVerse() As String
  Dim S As String, sAry() As String, T As String, tAry() As String, TT As String, TTT As String
  Dim ST As String, Edt As String, GrkAry() As String, DefRef4 As String, SS As String
  Dim Idx As Long, I As Long, J As Long, K As Long, TTTL As Long, TL As Long
  Dim Plurality As Boolean, nBmp As Boolean, bolVal As Boolean, bolClr As Boolean, Bol As Boolean
  Dim colLcl As Collection, colLcl2 As Collection
'
' ensure proper book and chapter displayed in treeview
'
  Me.rtbNotes.ToolTipText = vbNullString
  Call EnsureTVSet
'
' obtain a title for this entry in the format "Book, chapter:verse"
'
  sAry = Split(Books(Bk), ",")  'grab the book information
  Ttl = sAry(3) & " " & CStr(Chp) & ":" & CStr(Vrs)
  Me.lstGrkWords.Clear
  Me.lstWords.Clear
  
  Me.rtbUser.Text = vbNullString
  Me.cmdUpdateMPV.Enabled = False
  S = Format$(Bk, "00") & Format$(Chp, "00") & Format$(Vrs, "00")
'
' see if we should update the history
'
  If Not HistUpdt Then
    With colHist
      If .Count > 0 Then                'something stored?
        Do While .Count > HistIdx       'back off list to history index, if needed
          .Remove .Count
        Loop
        HistIdx = .Count                'save count
        If .Item(colHist.Count) = S Then nBmp = True  'do not update if matches current
      End If
      If Not nBmp Then
        .Add S                          'add new entry
        If .Count = 1000 Then .Remove 1  'keep list down to size by removing ancient verse
        HistIdx = .Count                'set new index to top of list
        bolVal = .Count > 1
        Me.cmdHBack.Enabled = bolVal    'enable/disable buttons as needed
        If bolVal Then Me.cmdHBack.ToolTipText = "Previous verse in history: " & GetVerseData(HistIdx - 1)
        Me.cmdHNext.Enabled = False
      End If
    End With
'
' add verse to session history, if different from current
'
    With colHistory
      If .Count <> 0 Then
        If .Item(.Count) <> S Then
          .Add S
        End If
      Else
        .Add S
      End If
    End With
  End If
  
  VrsIdx = FindExactMatch(Me.lstGrk, S)     'point to Greek text for verse
  If VrsIdx <> -1 Then
    ST = Mid$(Grk(VrsIdx), 8)               'strip off leader
    If Len(ST) <> 0 Then
      GrkAry = Split(ST, " ")               'get word list
      For Idx = 0 To UBound(GrkAry)         'fill listbox
        Me.lstGrkWords.AddItem GrkAry(Idx)
      Next Idx
    End If
  End If
  
  InitMerge
'
' add theological verse notes if present
'
  T = "No Theological Verse Notes for " & Ttl           'init no theological verse notes
  TL = Len(T)
  I = FindExactMatch(Me.lstVNotes, S)       'find a match
  If I <> -1 Then                           'found something...
    J = I                                   'save last-found index
    T = "Theological Verse Update Notes for " & Ttl 'init header
    TL = Len(T)
    Do While I <> -1
      T = T & vbCrLf & vbCrLf & Mid$(VNotes(I), 8)  'add a note
      I = FindExactMatch(Me.lstVNotes, S, I)  'find another
      If J >= I Then Exit Do                 'ignore if index matches last
    Loop
  End If
'
' add personal notes, if present
'
  TT = Mid$(MyNotes(VrsIdx), 8)
  If Len(TT) <> 0 Then                      'found notes
    Me.lblPersonal.Visible = True           'so show field indicating personal notes here
    J = InStr(1, TT, "\")                   'convert '\' to vbCrLf
    Do While J <> 0
      TT = Left$(TT, J - 1) & vbCrLf & Mid$(TT, J + 1)
      J = InStr(J + 2, TT, "\")
    Loop
'''    AddMergeCrLf2 T, , FntSize
'
' handle embedded Greek text
'
    T = T & vbCrLf & vbCrLf
    J = InStr(1, T, "{")
    Do While J <> 0
      AddMerge Left$(T, J - 1), , FntSize
      I = InStr(J + 1, T, "}")
      If I = 0 Then Exit Do                 'no match
      AddMerge Mid$(T, J + 1, I - J - 1), "Symbol", FntSize
      T = Mid$(T, I + 1)
      J = InStr(1, T, "{")
    Loop
    If Len(T) <> 0 Then
      AddMerge T, , FntSize
    End If
    AddMergeCrlf MyPersonalNotes, , FntSize, True, True ', , PNotesColor
    T = TT
  Else
    Me.lblPersonal.Visible = False          'else ensure personal note flag off
  End If
  
  With Me.rtbVerseNotes                     'set verse notes
    J = InStr(1, T, "{")
    Do While J <> 0
      AddMerge Left$(T, J - 1), , FntSize
      I = InStr(J + 1, T, "}")
      If I = 0 Then Exit Do                 'no match
      AddMerge Mid$(T, J + 1, I - J - 1), "Symbol", FntSize
      T = Mid$(T, I + 1)
      J = InStr(1, T, "{")
    Loop
    If Len(T) <> 0 Then
      AddMerge T, , FntSize
    End If
    
    SetIndent                             'set indenting
    LockWindowUpdate .hwnd                'lock updates to avoid flashing
    
    .Text = vbNullString                  'clear receiving buffer
    .TextRTF = Me.rtbMerge.TextRTF        'stuff new data
    .SelStart = 0                         'init all data black
    .SelLength = Len(.Text)
    .SelColor = vbBlack
    
    I = InStr(1, .Text, MyPersonalNotes)  'find personal notes
    If I <> 0 Then
      .SelStart = I - 1
      .SelLength = Len(.Text) - I
      .SelColor = PNotesColor             'set color if found
    End If
    
    .SelStart = 0                         'set title for window
    .SelLength = TL
    .SelBold = True
    .SelColor = vbBlue
    .SelLength = 0
    LockWindowUpdate 0                    'now refresh display
  End With
'
' update the English verse data
'
  TTT = vbNullString
  UserIndex = FindExactMatch(Me.lstGrk, S)  'find tranditional/personal entry
  If UserIndex <> -1 Then
    With Me.rtbVerse
      LockWindowUpdate .hwnd
      If Len(Bible(UserIndex)) < 8 Then    'if no verse data exists
        .BackColor = cMedium
        .Text = VersionText & ": " & Ttl & vbCrLf & vbCrLf & NoVerseTextAvail
      Else
        T = Mid$(Bible(UserIndex), 8)
'
' change [] to {} for versions that embrace certain words, such YLT
'
        I = InStr(1, T, "[")
        Do While I <> 0
          Mid$(T, I, 1) = "{"
          I = InStr(1, T, "[")
        Loop
        
        I = InStr(1, T, "]")
        Do While I <> 0
          Mid$(T, I, 1) = "}"
          I = InStr(1, T, "]")
        Loop
'
' see if we should extract JKV words and provide modern definitions for them
'
        If TranslateKJV Then                    'option in Bible Menu checked?
          SS = UCase$(T)                        'yes, so grab a copy of the verse text
          For I = 1 To Len(SS)
            Select Case Mid$(SS, I, 1)          'strip non-character letters
              Case "A" To "Z", "-", "'"
              Case Else
                Mid$(SS, I, 1) = " "
            End Select
          Next I
          TT = " " & LCase$(Trim$(SS)) & " "    'init test string
'
' read the KJV dictionary
'
          Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\KJVDict.txt", ForReading, False)
          tAry = Split(ts.ReadAll, vbCrLf)
          ts.Close
'
' init work collections
'
          Set colLcl = New Collection
          Set colLcl2 = New Collection
'
' build a list of KJV words in colLcl collection
'
          With colLcl
            For I = 1 To UBound(tAry)
              SS = tAry(I)
              J = InStr(1, SS, vbTab)
              If J <> 0 Then
                .Add Left$(SS, J - 1)
              End If
            Next I
'
' now find each word/phrase in the current verse that matches KJV English
'
            Do While .Count
              SS = " " & .Item(1) & " "     'get test word/phrase
              J = InStr(1, TT, LCase$(SS))  'check for a match
              If J <> 0 Then                'match found?
                Do While J <> 0             'yes
                  TT = Left$(TT, J) & "[" & Trim$(SS) & "]" & Mid$(TT, J + Len(SS) - 1)
                  T = Left$(T, J - 1) & "[" & Mid$(T, J, Len(SS) - 2) & "]" & Mid$(T, J + Len(SS) - 2)
                  J = InStr(1, TT, LCase$(SS))  'add braces, and check for all that match
                Loop
                colLcl2.Add Trim$(SS)       'save copy of found match
              End If
              .Remove 1
            Loop
          End With
'
' now see if any matches for KJV words were located
'
          With colLcl2
            If .Count <> 0 Then             'any matches?
              T = T & vbCrLf & vbCrLf       'yes, add blank space to main text
              TTT = "Common KJV Translations to Modern English:"  'init additional text
              TTTL = Len(TTT)
              
              Do While .Count
                For I = 1 To UBound(tAry)   'find each word in master KJV, get translation
                  TT = tAry(I)
                  J = InStr(1, TT, vbTab)
                  If Left$(TT, J - 1) = .Item(1) Then Exit For
                Next I
                SS = "[" & .Item(1) & "]"   'embrace KJV word
                If Len(SS) < 15 Then SS = SS & String$(15 - Len(SS), " ")
                TTT = TTT & vbCrLf & SS & Mid$(TT, J + 1) & "." 'append with description
                .Remove 1
              Loop
            End If
          End With
          Set colLcl = Nothing              'release collection data
          Set colLcl2 = Nothing
        End If
'
' check main text for additional notes ("\" = newline)
'
        I = InStr(1, T, "\")
        If I <> 0 Then
          ST = vbCrLf & Mid$(T, I + 1)
          T = Left$(T, I - 1)
          I = InStr(3, ST, "\")
          Do While I <> 0
            ST = Left$(ST, I - 1) & vbCrLf & Mid$(ST, I + 1)
            I = InStr(I + 2, ST, "\")
          Loop
        Else
          ST = vbNullString
        End If
        .Text = VersionText & ": " & Ttl & vbCrLf & vbCrLf & T
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = FntSize
        .SelBold = False
        If Mid$(Bible(UserIndex), 7, 1) = "*" Then  'if user-updated...
          .SelColor = vbWhite
          .BackColor = Navy
        Else
          .SelColor = vbBlack               'use default display
          .BackColor = cMedium
        End If
        If Len(ST) <> 0 Then
          I = Len(.Text)
          .SelStart = I
          .SelText = ST
          .SelStart = I
          .SelLength = Len(.Text) - I
          If .BackColor = cMedium Then      'if verse if not personal
            .SelColor = RGB(192, 0, 0)      'Red note
          Else
            .SelColor = vbYellow            'else personal, so add as Yellow on Navy
          End If
        End If
'
' append KJV text
'
        If Len(TTT) <> 0 Then
          bolClr = Mid$(Bible(UserIndex), 7, 1) = "*" 'if user-updated...
          I = Len(.Text)
          .SelStart = I
          .SelText = TTT                'append new text
          .SelStart = I
          .SelLength = Len(.Text) - I
          If bolClr Then
            .SelColor = vbCyan
          Else
            .SelColor = Navy            'use default display
          End If
          .SelFontName = "Courier New"  'fixed pitch
          .SelFontSize = FntSize - 2    'smaller point size
          With Me.lblWidth2
            .FontName = "Courier New"
            .FontSize = FntSize - 2     'compute hanging indent twip size
            .Caption = String$(15, " ")
            I = .Width
          End With
          .SelHangingIndent = I         'set hanging indent for new text
          If bolClr Then
            .SelColor = vbCyan
          Else
            .SelColor = Navy            'use default display
          End If
          .SelBold = True
          .SelLength = TTTL
          .SelUnderline = True
          If bolClr Then
            .SelColor = vbWhite          'use default display
          Else
            .SelColor = vbBlue
          End If
        End If
        .SelStart = 0
        .SelLength = 0
      End If
      LockWindowUpdate 0              'allow updates to window again
    End With
'
' get the references for the greek words in the verse
'
    GrkIdx = FindExactMatch(Me.lstGrk, S)
    ST = vbNullString
    Edt = vbNullString
    If Len(GrkBBL(GrkIdx)) > 7 Then
      BBLLine = Split(GrkBBL(GrkIdx), " ")      'get the word index for the greek
      MiniMap = Split(WordMap(GrkIdx), " ")     'grab the map of user selected words
      For Idx = 1 To UBound(BBLLine)            'process each entry in the list
        Plurality = False                       'init singular
        T = BBLLine(Idx)
        If IsNumeric(T) Then                    'valid index for greek word?
          I = CLng(T)                           'yes, so grab it
          tAry = Split(DefRef(I), vbTab)        'get the word reference data
          K = CLng(tAry(5))                     'grab the Strongs word index
          tAry = Split(WordRef(K), vbTab)       'grab the word data
          tAry = Split(tAry(2), ",")            'obtain the word list
          J = CLng(MiniMap(Idx))                'get the minimap data for the word
          If J < 0 Then                         'not yet user-defined?
            J = BBLWIdx(K)                      'yes, so grab the current default
            MiniMap(Idx) = J                    'add also to minimap
            AutoDirty = True                    'something has changed
          End If
          T = GrkAry(Idx - 1)                   'get a word
'
' do a minimal plurality check...
'
          Plurality = CheckPlural(T)
'
' get the word mapped to. The error checking was for early tests that have since been
' fixed, but this is just in case I forgot one.  If an error, ensure that the index is
' simply set to the last valid entry in the list
'
          On Error Resume Next
          T = tAry(J)                           'grab the indexed English word
          If Err.Number <> 0 Then               'if the index is bad
            J = UBound(tAry)                    'get upper bound
            MiniMap(Idx) = J                    'update map
            BBLWIdx(K) = J                      'fix the general default as well
            T = tAry(J)                         'get word there
          End If
          On Error GoTo 0
'
' if plurality assumed, apply it in an obvious manner
'
          If Plurality Then
            T = T & "(s)"
          End If
          If Idx = 1 Then                       'if first word in verse...
            T = UCase$(Left$(T, 1)) & Mid$(T, 2)  'begin sentense with capitalized letter
          End If
          T = "[" & T & "]"                     'enbrace word(s); ie, "[I will not]"
        End If
        ST = ST & " " & T                       'accumulate direct translation text
'
' update storage for extracting data from the direct translation
'
        If T <> "[*]" And T <> "[*(s)]" Then    'ignorable enties. Personally, I do not use
          Edt = Edt & " " & Mid$(T, 2, Len(T) - 2)  'this. Words were put there for a reason.
        End If
      Next Idx
      WordMap(GrkIdx) = Join(MiniMap, " ")   'update word selection mapping
'
' now update the Direct Translation textbox, and the "extract" text
'
      With Me.rtbTranslate
        .Text = "Direct Translation:" & vbCrLf & vbCrLf & Mid$(ST, 2)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = FntSize
        .SelBold = False
        .SelColor = vbBlack
        .SelLength = 19
        .SelBold = True
        .SelColor = vbBlue
      End With
      UserText = Mid$(Edt, 2)                   'save for user hitting "Copy translation"
    Else
      UserText = vbNullString
      Me.rtbTranslate.Text = vbNullString
      Me.rtbNotes.Text = vbNullString
    End If
  Else
    UserText = vbNullString
    Me.rtbTranslate.Text = vbNullString
    Me.rtbVerse.Text = VersionText & ": " & Ttl & "[NO VERSE DATA]"
    Me.rtbNotes.Text = vbNullString
  End If
  
  With Me.rtbVerse
    .SelStart = 0
    .SelLength = Len(VersionText & Ttl) + 2
    .SelBold = True
    If .BackColor = Navy Then
      .SelColor = vbCyan
    Else
      .SelColor = vbBlue
    End If
    .SelLength = 0
  End With
'
' update command buttons for extracting data
'
  Bol = PersonalVersion = True And BblVersion = UserPVer And Len(UserText) > 0
  Me.cmdCpyXlt.Enabled = Bol
  Me.cmdCopympv.Enabled = Bol  'if no actual Greek verse, then avoid copies
'
' update Greek text
'
  T = "Greek: " & Ttl & vbCrLf & vbCrLf
  TL = Len(T) - 4
  TT = vbNullString
  VrsIdx = FindExactMatch(Me.lstGrk, S) 'variable S still contains "010203" verse header
  If VrsIdx <> -1 Then
    S = Mid$(Grk(VrsIdx), 8)            'grab the Greek text, strip header
    If Len(S) = 0 Then                  'no actual Greek text exists (many bible verses were
      'ecclesiastically added, either to the Greek, or they
      'have NEVER existed, except in phantom "translations".
      T = NoGreekText & Ttl & vbCrLf & vbCrLf
      TL = Len(T) - 4
      If Right$(Grk(VrsIdx), 1) = "*" Then
        T = T & "The Greek text for this verse has NEVER existed. " & _
                "The reason that this verse exists at all is because " & _
                "it was invalidly added later to the Latin Vugate Bible, or some " & _
                "other invalid later-added tradition, which many subsequent Bible " & _
                "versions naively adopted." & vbCrLf & vbCrLf & _
                "See Theological Verse Update Notes for details."
      Else
        T = T & "The Greek text for this verse was added long after this book " & _
                "was originally written. As such, this invalid Greek text has " & _
                "therefore been properly excised, per Theological instructions." & vbCrLf & vbCrLf & _
                "See Theological Verse Update Notes for details."
      End If
      Me.lblFeminine.Visible = False
      Me.lblPlural.Visible = False
      lblNoteInfo.Visible = False
      Me.rtbUser.Enabled = False
      Me.cmdVine.Visible = False
'
' if we are working with a personal version, and the Greek text does not exist, do not show
' any "traditional" translations they have not yet worked on. Ensure the text is also
' removed from their saved version of the book.
'
      If BblVersion = UserPVer Then     'if personal version, remove any verse notes
        Me.rtbVerse.Text = VersionText & ": " & Ttl & vbCrLf & vbCrLf & NoVerseTextAvail
        ST = Bible(UserIndex)           'get the user verse data
        If Len(ST) > 7 Then              'if data exists, remove this invalid content
          Bible(UserIndex) = Left$(ST, 6)
          PVDirty = True                'tag for update
        End If
      End If
      Call mnuFileViewChapter_Click
      SS = "(" & CStr(Vrs) & ") "        'verse to find
      TT = "(" & CStr(Vrs + 1) & ") "   'next verse
      With Me.rtbNotes
        S = .Text
        Idx = InStr(1, S, SS)
        If Idx <> 0 Then
          LockWindowUpdate .hwnd
          J = InStr(Idx + Len(SS), S, TT)  'find next
          K = InStr(Idx + Len(SS), S, MyPersonalNotes)  'find next
          If K > 0 And K < J Then J = K
          If J = 0 Then J = Len(S) + 1    'if not found, assume SS was last
          .SelStart = Len(S)              'to end of chapter
          .SelStart = Idx - 1             'now select to force scrollup
          .SelLength = J - Idx            'highlight target verse
          S = RTrim$(.SelText)
          .SelLength = Len(S)
          LockWindowUpdate 0
        End If
        S = vbNullString
      End With
      lblNoteInfo.Visible = False
      Me.cmdAnalysis.Enabled = False
      Me.mnuFileAnalysis.Enabled = False
    Else
      Me.rtbUser.Enabled = True
      lblNoteInfo.Visible = True
      Me.cmdAnalysis.Enabled = True
      Me.mnuFileAnalysis.Enabled = True
    End If
'
' now update the Greek text display
'
    With Me.rtbGreek
      LockWindowUpdate .hwnd
      .Text = T & S
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelFontName = "Symbol"
      .SelColor = vbBlack
      .SelFontSize = FntSize + 4
      .SelLength = Len(T)
      .SelFontName = "Times New Roman"
      If Len(TT) <> 0 Then
        .SelColor = cBurgandy
      Else
        .SelColor = vbBlack
      End If
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelBold = False
      .SelLength = TL
      .SelBold = True
      .SelColor = vbBlue
      .SelFontSize = FntSize
      LockWindowUpdate 0
    End With
'
' set the first word of the greek text in the list (or nothing if no words)
'
    If Len(S) > 0 Then
      Me.lstGrkWords.ListIndex = 0
    Else
      Me.lstGrkWords.ListIndex = -1
    End If
    S = "X"
  Else
    S = vbNullString
    With Me.rtbGreek
      LockWindowUpdate .hwnd
      .Text = T
      .SelStart = 0
      .SelLength = Len(T)
      .SelFontName = "Times New Roman"
      .SelColor = vbBlack
      .SelLength = 0
      LockWindowUpdate 0
    End With
  End If
'
' update the current book/chapter/verse in the form caption
'
  Me.Caption = "New Covenant Bible Greek Translator - " & Ttl
'
' enable buttons/menu entries as needed
'
  On Error Resume Next                          'if during Form_Load...
  Me.cmdCopy.Enabled = True
  Me.cmdCopyAll.Enabled = True
  Me.mnuFileCreateBible.Enabled = True
  Me.lstGrkWords.SetFocus
  Me.mnuFav.Enabled = True
  Me.cmdCopyDef.Enabled = True
  Me.cmdCopy.Enabled = True
  Me.cmdCopyVerse.Enabled = True
  Me.cmdCopyNotes.Enabled = True
  Me.cmdCopyDT.Enabled = True
  Me.cmdFindInText.Enabled = Len(Me.rtbNotes.Text) > 0
  Me.cmdFindNext.Enabled = False
  Me.mnuFileViewChapter.Enabled = True
  Me.mnuFileTheoNext.Enabled = True
  Me.mnuFileTheoPrev.Enabled = True
  Me.mnuBBLCompare.Enabled = True
  Me.cmdAddNote.Enabled = True
  Me.mnuBBLFindNext.Enabled = HavePersonalNotes
  Me.mnuBBLFindPrev.Enabled = HavePersonalNotes
  Me.mnuBBLFindNextNoGreek.Enabled = True
  Me.mnuBBLFindPrevNoGreek.Enabled = True
  Me.mnuBBLViewAllPersonalNotes.Enabled = HavePersonalNotes
  Me.mnuBBLViewPNotesChapter.Enabled = HavePersonalNotes
  Me.mnuFileViewTheo.Enabled = True
  Me.mnuBBLOrgKJV.Enabled = True
  Me.mnuBibleTranslateKJV.Enabled = True
  With Me.Toolbar1
    .Buttons("prevnote").Enabled = HavePersonalNotes
    .Buttons("nextnote").Enabled = HavePersonalNotes
    .Buttons("prevtheo").Enabled = True
    .Buttons("nexttheo").Enabled = True
    .Buttons("prevgreek").Enabled = True
    .Buttons("nextgreek").Enabled = True
  End With
'
' update menu bar ebtry
'
  If SetGoto = False And Bk <> 0 Then
    Me.cboBk.ListIndex = Bk - 1
    Me.cboChp.ListIndex = Chp - 1
    Me.cboVrs.ListIndex = Vrs - 1
    Me.cmdGo.Enabled = False
  End If
  ShowVerse = S
End Function

'*******************************************************************************
' Function Name     : CheckPlural
' Purpose           : do a minimal plurality check...
'*******************************************************************************
Private Function CheckPlural(Text As String) As Boolean
  Dim Plurality As Boolean, DidIt As Boolean

  DidIt = False
  Select Case Text
    Case "oi", "ai", "ta", "twn", "toiV", "taiV", "toiV", "touV", "taV", "autoV", "touV", "tiV"
      DidIt = True      'inicate we have handled the situation
      Plurality = True  'known plural words
    Case "o", "h", "to", "ton", "th", "thn", "tou", "thV" ' "o" is acutall rather general
      DidIt = True      'known singular words
    Case "enaV", "ena", "enan", "mia", "mian", "mias", "enoV"
      DidIt = True      'known signular words
      
  End Select
  '
  ' if no matches, check for the word being longer than 3 characters,
  ' and check endings on them
  '
  If Not DidIt And Len(Text) > 3 Then               'if it contains more than 3 characters
    Select Case Right$(Text, 3)
      Case "eiV", "oiV", "ouV", "aiV", "auV", "ewn", "eiV", "enh", "ete", "uin", "uma"
        Plurality = True                            'assume plural
        DidIt = True
      Case "eon", "qia", "quV", "mia", "ria", "iaV", "ion"
        DidIt = True                                'assume singular
    End Select
    If Not Plurality And Not DidIt Then             'if matches have not yet been hit
      Select Case Right$(Text, 2)                   'check 2-character endings
        Case "wn", "un", "uV", "in", "aV", "iV", "on", "oi", "ai", "ta", "la", "ia"
          Plurality = True                          'assume plural
          DidIt = True
      End Select
    End If
  End If                                            'assume sigular otherwise
  CheckPlural = Plurality
End Function

'*******************************************************************************
' Function Name     : GetVerseData
' Purpose           : Support routine for History processing
'*******************************************************************************
Public Function GetVerseData(ByVal Index As Long) As String
  Dim S As String, Ary() As String
  Dim B As Long
  
  S = colHist(Index)
  B = CLng(Left$(S, 2))
  Ary = Split(Books(B), ",")
  GetVerseData = Ary(3) & " " & CStr(CLng(Mid$(S, 3, 2))) & ":" & CStr(CLng(Right$(S, 2)))
End Function

'*******************************************************************************
' Subroutine Name   : UpdateVerse
' Purpose           : Apply updates to verse
'*******************************************************************************
Public Sub UpdateVerse()
  Dim S As String
    
  ChgSCroll = True              'redunancy protection
  S = ShowVerse()               'display the verse
  If Len(S) <> 0 Then           '"X" or ""
    S = "Verse " & CStr(Vrs) & " of " & CStr(VrsCnt)  'update for report label
    Me.hsGreek.Value = 0        'reset selection scroll bar
    Me.hsGreek.Max = VrsCnt - 1
  End If
  Me.hsGreek.Value = Vrs - 1    'set position of scroll tab
  ChgSCroll = False             'turn off protection
  Me.lblVerseIndex.Caption = S  'update the report label
End Sub

'*******************************************************************************
' Subroutine Name   : InitMerge
' Purpose           : Clear the merge box, repare for adding
'*******************************************************************************
Public Sub InitMerge()
  Me.rtbMerge.Text = vbNullString
  Me.txtMerge.Text = vbNullString
End Sub

'*******************************************************************************
' Subroutine Name   : AddMergeCrlf
' Purpose           : Add text, followed by a CRLF
'*******************************************************************************
Public Sub AddMergeCrlf(Text As String, _
                         Optional FntName As String = "Times New Roman", _
                         Optional FSize As Long = 0, _
                         Optional Bld As Boolean = False, _
                         Optional Itl As Boolean = False, _
                         Optional Alignment As AlignmentConstants = vbLeftJustify, _
                         Optional Clr As Long = vbBlack)
  Call AddMerge(Text & vbCrLf, FntName, FSize, Bld, Itl, Alignment, Clr)
End Sub

'*******************************************************************************
' Subroutine Name   : AddMergeCrLf2
' Purpose           : Add text, followed by 2 CRLFs
'*******************************************************************************
Public Sub AddMergeCrLf2(Text As String, _
                          Optional FntName As String = "Times New Roman", _
                          Optional FSize As Long = 0, _
                          Optional Bld As Boolean = False, _
                          Optional Itl As Boolean = False, _
                          Optional Alignment As AlignmentConstants = vbLeftJustify, _
                          Optional Clr As Long = vbBlack)
  Call AddMerge(Text & vbCrLf & vbCrLf, FntName, FSize, Bld, Itl, Alignment, Clr)
End Sub

'*******************************************************************************
' Subroutine Name   : AddMergeSep
' Purpose           : Append an underscore line, followed by 2 CRLF
'*******************************************************************************
Public Sub AddMergeSep()
  AddMergeCrLf2 Underline
End Sub

'*******************************************************************************
' Subroutine Name   : AddMerge
' Purpose           : Accumulative add of text to the merge box without a Newline
'*******************************************************************************
Public Sub AddMerge(Text As String, _
                     Optional FntName As String = "Times New Roman", _
                     Optional FSize As Long = 0, _
                     Optional Bld As Boolean = False, _
                     Optional Itl As Boolean = False, _
                     Optional Alignment As AlignmentConstants = vbLeftJustify, _
                     Optional Clr As Long = vbBlack)
  
  Dim SS As Long, Fnt As Long, I As Long
  Dim lclText As String
  
  Fnt = FSize                     'get font point size
  If Fnt = 0 Then Fnt = 10        'use 10 if not specified
  
  If Not DoBible Then
    With Me.rtbMerge
      SS = Len(.Text)               'get the current length of the RTB text
      .SelStart = SS                'set the selection point to the end of the text
      .SelLength = 0                'usually not needed, but it makes intention clear
      .SelText = Text               'stuff the new text
      .SelStart = SS                'ensure selstart is reset to the start of the new text
      .SelLength = Len(.Text) - SS  'select just the new text
      .SelFontName = FntName        'set the Font style
      .SelFontSize = Fnt            'set the point size
      .SelBold = Bld                'and if we want it enboldened
      .SelItalic = Itl              'and italicized
      .SelColor = Clr
      .SelAlignment = Alignment     'set alignment
      .SelStart = 0                 'remove any selection of text
    End With
    Exit Sub
  End If
  
  With Me.txtMerge
    SS = Len(.Text)               'get the current length of the RTB text
    .SelStart = SS                'set the selection point to the end of the text
    .SelLength = 0                'usually not needed, but it makes intention clear
    If IsRTF Then
      Select Case FntName
        Case "Arial"
          lclText = "\f0"
        Case "Symbol"
          lclText = "\f2"
        Case Else
          lclText = "\f1"
      End Select
      Select Case Alignment
        Case rtfLeft
          lclText = lclText & "\ql"
        Case rtfCenter
          lclText = lclText & "\qc"
        Case Else
          lclText = lclText & "\qr"
      End Select
      If Bld Then lclText = lclText & "\b"
      If Itl Then lclText = lclText & "\i"
      lclText = lclText & "\fs" & CStr(Fnt * 2) & " " & Text
      I = InStr(1, lclText, vbFormFeed)
      Do While I
        lclText = Left$(lclText, I - 1) & "\page " & Mid$(lclText, I + 1)
        I = InStr(I + 6, lclText, vbFormFeed)
      Loop
      I = InStr(1, lclText, vbCrLf)
      Do While I
        lclText = Left$(lclText, I - 1) & "\par " & Mid$(lclText, I)
        I = InStr(I + 7, lclText, vbCrLf)
      Loop
      If Itl Then lclText = lclText & "\i0"
      If Bld Then lclText = lclText & "\b0"
      lclText = lclText & "\ql"
      ts.Write lclText
    Else
      ts.Write Text
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : SetIndent
' Purpose           : Set hanging indet for the RTB Merge box
'*******************************************************************************
Public Sub SetIndent(Optional Start As Long = 0, Optional Length As Long = 0, Optional Indent As Long = 0)
  Dim Idnt As Long
  
  Idnt = Indent
  If Idnt = 0 Then Idnt = HIndent
  
  With Me.rtbMerge
    .SelStart = Start           'set the start of the selected text (usually at the beginning)
    If Length > 0 Then
      .SelLength = Length       'apply a user-defined length
    Else
      .SelLength = Len(.Text)   'else go to the end of the text (usual venue)
    End If
    .SelHangingIndent = Idnt    'set the hanging indent
    .SelStart = 0               'deselect the text
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : AddVerse
' Purpose           : Add a Verse to the chapter
'*******************************************************************************
Private Sub AddVerse(Text As String, _
                     Verse As Long, _
                     Optional FntName As String = "Times New Roman", _
                     Optional FSize As Long = 0, _
                     Optional Bld As Boolean = False, _
                     Optional Itl As Boolean = False, _
                     Optional Index As Long = -1)
  
  Dim Term As String, lclText As String, GrkTxt As String, S As String
  Dim SS As Long, Fnt As Long, I As Long, J As Long
 
  Fnt = FSize                     'get font point size
  If Fnt = 0 Then Fnt = 10        'use 10 if not specified
  Fnt = Fnt + BumpFactor
  
  If VerseLines Then            '1 verse per line?
    Term = vbCrLf               'yes, so dump full line
  Else
    If Index = -1 Then
      Select Case Right$(Text, 1)
        Case ".", "?", "!"        'check for sentense terminators
          Term = vbCrLf & "    "  'if terminator, then add CRLF
        Case Else
          Term = " "               'else append a space
      End Select
    Else
      Term = vbCrLf & "    "
      If Mid$(ParMap, Index + 2, 1) = " " Then
        Term = " "
      End If
    End If
  End If
  
  If Not DoBible Then
    With Me.rtbMerge
      If Verse <> 0 Then
        lclText = "(" & CStr(Abs(Verse)) & ") "
        If Abs(Verse) = 1 And VerseLines = False Then
          SS = Len(.Text)               'get the current length of the RTB text
          .SelStart = SS                'set the selection point to the end of the text
          .SelLength = 0                'usually not needed, but it makes intention clear
          .SelText = "    "             'stuff indent
          .SelStart = SS                'ensure selstart is reset to the start of the new text
          .SelLength = Len(.Text) - SS  'select just the new text
          .SelFontName = FntName        'set the Font style
          .SelFontSize = FntSize + BumpFactor 'set the point size
          .SelBold = False              'and if we want it enboldened
          .SelItalic = False            'and italicized
          .SelUnderline = False
        End If
        SS = Len(.Text)               'get the current length of the RTB text
        .SelStart = SS                'set the selection point to the end of the text
        .SelLength = 0                'usually not needed, but it makes intention clear
        .SelText = lclText            'stuff the new verse data
        .SelStart = SS                'ensure selstart is reset to the start of the new text
        .SelLength = Len(.Text) - SS  'select just the new text
        .SelFontName = "Arial"        'set the Font style
        .SelFontSize = 6 + BumpFactor 'set the point size
        .SelBold = True               'and if we want it enboldened
        .SelItalic = False            'and italicized
        .SelUnderline = False
        .SelColor = cBurgandy
      Else
        SS = Len(.Text)               'get the current length of the RTB text
        .SelStart = SS                'set the selection point to the end of the text
        .SelLength = 0                'usually not needed, but it makes intention clear
        .SelText = MyPersonalNotes    'stuff the new verse data
        .SelStart = SS                'ensure selstart is reset to the start of the new text
        .SelLength = Len(.Text) - SS  'select just the new text
        .SelFontName = "Arial"        'set the Font style
        .SelFontSize = Fnt - 2        'set the point size
        .SelBold = True               'and if we want it enboldened
        .SelItalic = True             'and italicized
        .SelUnderline = True
        .SelColor = PNotesColor
      End If
      
      S = Text
      I = InStr(1, S, "{")
      Do While I <> 0
        J = InStr(I + 1, S, "}")
        If J = 0 Then Exit Do         'no match
        SS = Len(.Text)               'get the current length of the RTB text
        .SelStart = SS                'set the selection point to the end of the text
        .SelLength = 0                'usually not needed, but it makes intention clear
        .SelText = Left$(S, I - 1)    'stuff the new text
        .SelStart = SS                'ensure selstart is reset to the start of the new text
        .SelLength = Len(.Text) - SS  'select just the new text
        .SelFontName = FntName        'set the Font style
        .SelFontSize = Fnt            'set the point size
        .SelBold = Bld                'and if we want it enboldened
        .SelItalic = Itl              'and italicized
        .SelUnderline = False
        If Verse < 0 Then
          .SelColor = cMissing
        ElseIf Verse = 0 Then
          .SelColor = PNotesColor
        Else
          .SelColor = vbBlack
        End If
        
        GrkTxt = Mid$(S, I + 1, J - I - 1)
        SS = Len(.Text)               'get the current length of the RTB text
        .SelStart = SS                'set the selection point to the end of the text
        .SelLength = 0                'usually not needed, but it makes intention clear
        .SelText = GrkTxt             'stuff the new text
        .SelStart = SS                'ensure selstart is reset to the start of the new text
        .SelLength = Len(.Text) - SS  'select just the new text
        .SelFontName = "Symbol"       'set the Font style
        .SelFontSize = Fnt            'set the point size
        .SelBold = True               'and if we want it enboldened
        .SelItalic = False            'and italicized
        .SelUnderline = False
        If Verse < 0 Then
          .SelColor = cMissing
        ElseIf Verse = 0 Then
          .SelColor = PNotesColor
        Else
          .SelColor = vbBlack
        End If
        
        S = Mid$(S, J + 1)
        I = InStr(1, S, "{")
      Loop
      S = S & Term                    'add possible termator
      If Len(S) <> 0 Then             'anything to process
        SS = Len(.Text)               'get the current length of the RTB text
        .SelStart = SS                'set the selection point to the end of the text
        .SelLength = 0                'usually not needed, but it makes intention clear
        .SelText = S                  'stuff the new text
        .SelStart = SS                'ensure selstart is reset to the start of the new text
        .SelLength = Len(.Text) - SS  'select just the new text
        .SelFontName = FntName        'set the Font style
        .SelFontSize = Fnt            'set the point size
        .SelBold = Bld                'and if we want it enboldened
        .SelUnderline = False
        If Verse < 0 Then
          .SelColor = cMissing
        ElseIf Verse = 0 Then
          .SelColor = PNotesColor
        Else
          .SelColor = vbBlack
        End If
      End If
    End With
    Exit Sub
  End If
    
  S = Text & Term                   'add possible termator to text
  
  With Me.txtMerge
    If Not IsRTF Then
      SS = Len(.Text)               'get the current length of the RTB text
      .SelStart = SS                'set the selection point to the end of the text
      .SelLength = 0                'usually not needed, but it makes intention clear
    End If
    
    If Verse <> 0 Then
      If IsRTF Then
        If Abs(Verse) = 1 And VerseLines = False Then
          lclText = "\fs" & CStr(Fnt * 2) & "     " & _
                    "\f0\fs" & CStr((6 + BumpFactor) * 2) & "\cf3 (" & CStr(Abs(Verse)) & ")\cf0  "
        Else
          lclText = "\f0\fs" & CStr((6 + BumpFactor) * 2) & "\cf3 (" & CStr(Abs(Verse)) & ")\cf0  "
        End If
      Else
        If Abs(Verse) = 1 And VerseLines = False Then
          S = "    (" & CStr(Abs(Verse)) & ") " & S
        Else
          S = "(" & CStr(Abs(Verse)) & ") " & S
        End If
      End If
    Else
      If IsRTF Then
        lclText = lclText & "\fs" & CStr(Fnt * 2) & " " & "\ul\i \b\cf1 " & MyPersonalNotes & "\cf0 \b0\i0\ulnone"
      Else
        S = MyPersonalNotes & S
      End If
    End If
    
    If IsRTF Then
      If Len(lclText) <> 0 Then ts.Write lclText
'
' convert vbCrLf to "\par "
'
      I = InStr(1, S, vbCrLf)
      Do While I
        S = Left$(S, I - 1) & "\par " & Mid$(S, I)
        I = InStr(I + 7, S, vbCrLf)
      Loop
'
' scan for Greek text
'
      I = InStr(1, S, "{")
      Do While I <> 0
        J = InStr(I + 1, S, "}")
        If J = 0 Then Exit Do     'no match
        
        Select Case FntName
          Case "Arial"
            lclText = "\f0"
          Case Else
            lclText = "\f1"
        End Select
        If Verse = 0 Then
          lclText = lclText & "\fs" & CStr(Fnt * 2) & " " & "\cf1 " & Left$(S, I - 1) & "\cf0 "
        ElseIf Verse < 0 Then
          lclText = lclText & "\fs" & CStr(Fnt * 2) & " " & "\cf2 " & Left$(S, I - 1) & "\cf0 "
        Else
          lclText = lclText & "\fs" & CStr(Fnt * 2) & " " & Left$(S, I - 1)
        End If
        ts.Write lclText
'
' process greek word
'
        GrkTxt = Mid$(S, I + 1, J - I - 1)
        ts.Write "\f2" & "\fs" & CStr(Fnt * 2) & " " & "\b\cf1 " & GrkTxt & "\cf0 \b0"
        S = Mid$(S, J + 1)
        I = InStr(1, S, "{")
      Loop
'
' process (remaining) non-greek data
'
      If Len(S) <> 0 Then             'anything to process
        Select Case FntName
          Case "Arial"
            lclText = "\f0"
          Case Else
            lclText = "\f1"
        End Select
        If Verse = 0 Then
          lclText = lclText & "\fs" & CStr(Fnt * 2) & " " & "\cf1 " & S & "\cf0 "
        ElseIf Verse < 0 Then
          lclText = lclText & "\fs" & CStr(Fnt * 2) & " " & "\cf2 " & S & "\cf0 "
        Else
          lclText = lclText & "\fs" & CStr(Fnt * 2) & " " & S
        End If
      End If
      ts.Write lclText
    Else
      ts.Write S
    End If
  End With
End Sub

'*******************************************************************************
' The following routines will significantly speed writing the bible to an RTF
' formatted file by writing it in segments.  The more formatting you apply to
' an RTF file, the more time it takes to process.  Hence, to write the file in
' blocks (by chapter), will allow this formatted write to run several times
' faster than formatting it as one big file.

' An RTF file's first line (a line is terminated by vbCrLf) contains the RTF
' opening block.  By removing the "}" at the end of the entire data, the block
' is left "open".  Chapters are written by removing this first line and the
' final, terminating "}".  Once all is written, the terminating "}" is added
' thus creating a complete RTF file.
'*******************************************************************************

'*******************************************************************************
' Subroutine Name   : WriteHeader
' Purpose           : This routine writes the RTF data, less the terminating "}"
'*******************************************************************************
Private Sub WriteHeader(Optional Clr As Long = 0)
  Dim Rd As Long, Gr As Long, Bl As Long
  Dim S As String
  
  If IsRTF Then
'
' compute color elements for the color table
'
  S = Hex$(Clr)
  If Len(S) <> 6 Then S = Left$("000000", 6 - Len(S)) & S
  Rd = CLng("&H" & Right$(S, 2))  'red
  Gr = CLng("&H" & Mid$(S, 3, 2)) 'green
  Bl = CLng("&H" & Left$(S, 2))   'blue
  '
  ' build header with color table
  '
    ts.Write "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Arial;}" & _
             "{\f1\fnil\fcharset0 Times New Roman;}" & vbCrLf & _
             "{\f2\fnil\fcharset2 Symbol;}}" & vbCrLf & _
             "{\colortbl ;\red" & CStr(Rd) & "\green" & CStr(Gr) & "\blue" & CStr(Bl) & ";" & _
             "\red255\green0\blue255;" & _
             "\red128\green0\blue0;" & _
             "\red192\green0\blue0;" & _
             "}" & vbCrLf & _
             "{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1"
  End If
  InitMerge
End Sub

'*******************************************************************************
' Subroutine Name   : WriteChapter
' Purpose           : Append the current RTF box contents to the file in progress
'*******************************************************************************
Private Sub WriteChapter()
  If Not DoBible Then
    ts.Write Me.txtMerge.Text     'if normal text, then simply write everything
    InitMerge
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : WriteTrailer
' Purpose           : Close up a segmented RTF write
'*******************************************************************************
Private Sub WriteTrailer()
  If IsRTF Then
    ts.WriteLine "}"                  'if RTF, send terminating "}"
  ElseIf Len(Me.txtMerge.Text) <> 0 Then
    WriteChapter                      'write any remaining data
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Vbarz
' Purpose           : Draw a light vertical bar on the left and a dark vertical bar
'                   : on the right of a picture slider that would not otherwise have
'                   : a border, giving it a 3D effect that it would otherwise not have.
'*******************************************************************************
Public Sub Vbarz(Pic As PictureBox)
  With Pic
    Pic.Line (0, 0)-(0, .Height), RGB(241, 239, 226)                     'left side
    Pic.Line (15, 0)-(15, .Height), vbWhite                              'left side
    Pic.Line (.Width - 30, 0)-(.Width - 30, .Height), RGB(172, 168, 153) 'right side
    Pic.Line (.Width - 15, 0)-(.Width - 15, .Height), RGB(113, 111, 100) 'right side
  
    Pic.Line (0, 0)-(.Width - 15, 0), RGB(241, 239, 226)                  'top side
    Pic.Line (15, 15)-(.Width - 15, 15), vbWhite                          'top side
    Pic.Line (30, .Height - 30)-(.Width - 30, .Height - 30), RGB(172, 168, 153) 'right side
    Pic.Line (30, .Height - 15)-(.Width - 15, .Height - 15), RGB(113, 111, 100) 'right side
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : CheckWin
' Purpose           : Set Windows menu item visiblility if at least 1 item is active
'*******************************************************************************
Public Sub CheckWin()
  Me.mnuWin.Visible = Me.mnuWinBible.Enabled Or _
                      Me.mnuWinKJV.Enabled Or _
                      Me.mnuWinVine.Enabled Or _
                      Me.mnuWinSearch.Enabled Or _
                      Me.mnuWinStrong.Enabled
End Sub

'*******************************************************************************
' Subroutine Name   : SetVerseCount
' Purpose           : Rebuild the verse list for the current chapter
'*******************************************************************************
Private Sub SetVerseCount()
  Dim I As Long
  Dim S As String
  
  S = Format$(Me.cboBk.ListIndex + 1, "00") & Format$(Me.cboChp.ListIndex + 1, "00")
  I = FindExactMatch(Me.lstGrk, S & "01")   'find verse 1
  With Me.cboVrs
    .Clear                                  'remove any old data
    Do
      .AddItem CStr(.ListCount + 1)         'add an item
    Loop While Left$(frmGrkXlate.lstGrk.List(I + .ListCount), 4) = S
    .ToolTipText = "Select Verse 1 - " & CStr(.ListCount)
    .ListIndex = 0                          'force verse 1 (-1)
  End With
End Sub

