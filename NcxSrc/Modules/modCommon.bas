Attribute VB_Name = "modCommon"
Option Explicit
'
' set the following constant to TRUE to recreate a "Virgin" app image
'
Public Const MakeVirgin As Boolean = False

'*********************************************
' Global Constants
'*********************************************
Public Const TVM_SETBKCOLOR = 4381&
Public Const VK_CONTROL = &H11
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Const Underline As String = "____________________________________"
Public Const MyPersonalNotes As String = "My Personal Notes:"
Public Const NoVerseTextAvail As String = "(No Verse Data Available)"
Public Const DTHeaderOffset As Long = 24

'*********************************************
' Global Variables
'*********************************************
Public BblVersion As Long           'current bible version
Public PersonalVersion As Boolean   'if a personal version is available
Public PersonalVersionBase As Long  'the version to base a persional version on
Public PVDirty As Boolean           'if the personal version has been updated
Public BBLDirty As Boolean          'if the bible index has been updated
Public FntSize As Long              'keep track of selected font size
Public Const UserPVer As Long = 3
Public SplashIsOn As Boolean        'True when the splash screen is actually a splash screen
Public HavePersonalNotes As Boolean 'True if Personal Note entries exist
Public TipData As String            'contains flags for the tip of the day.
Public TranslateKJV As Boolean      'trnslate KJV to modern English
Public IsLoading As Boolean         'true when program is loading
Public WordPadPath As String        'Path to WordPad executable
Public NotePadPath As String        'Path to NotePad executable
'
' color definition storage
'
Public cDark As Long, cMedium As Long, cLight As Long, cVLight As Long
Public custDark As Long, custMedium As Long, custLight As Long, CustVLight As Long

Public clBlue As Long               'Vine database background
Public cdGray As Long               'user personal version edited verse
Public cMissing As Long             'set to verses that do not exist
Public cBurgandy As Long            'verse Numbers and missing Greek text
'
' file I/O support
'
Public Fso As FileSystemObject
Public ts As TextStream

Public Ttl As String                'Book, chapter:verse title for current verse
Public VersionText As String        'the full text of the current bible version
Public CurWord As String            'the current word
Public UserText As String           'version of direct translations without brackets
Public Books() As String            'Books.txt database
Public Grk() As String              'Greek.txt bible database
Public DefRef() As String           'GreekDefRef.txt database
Public WordRef() As String          'WordRef.txt database
Public GrkBBL() As String           'GreekBBL.txt database
Public Bible() As String            'KJV.TXT or RSV.TXT or YLT.TXT or MPV.TXT database
Public VNotes() As String           'Vnotes.txt database
Public BBLLine() As String          'local copy of the current verse from WordRef database
Public WordMap() As String          'GreekWordRef.txt database
Public MiniMap() As String          'local copy of the current verse from WordMap database
Public Vine() As String             'Vine word reference
Public VineText As String
Public MyNotes() As String          'My Personal notes
Public MyNotesDirty As Boolean      'if the personal notes contents has changed
Public ParMap As String             'paramgraph map
Public SearchBooks(27) As Boolean   'books to search
Public KJVCount() As String         'KJV word counts
Public GrkWrdCnt() As String        'Greek Word Counts
Public KJVidxAry() As String
Public KJVwrdAry() As String

Public Bk As Long                   'current book
Public Chp As Long                  'current chapter
Public Vrs As Long                  'current verse
Public ChpCnt As Long               'number of chapters in the current book
Public VrsCnt As Long               'number of verses in the current chapter
Public VrsIdx As Long               'a common indexing variable for verse offsets
Public GrkIdx As Long               'index into Greek words within a verse
Public BBLWIdx() As Long            'track last-used word-choice
Public Strong As Long               'current word's index (Strong's Reference Number)
Public DefRefIdx As Long            'keep track of DefRef Index
Public UserIndex As Long            'index into user verse list
Public VineIndex As Long            'index into Vine list
Public BumpFactor As Long           'font bumping factor when writing bible
Public ShowingBible As Boolean      'when viewing Bible in viewer

Public lstWordsClicked As Boolean   'redundency guard
Public ChgSCroll As Boolean         'scrolling flag to prevent redundency
Public Plural As Boolean            'plurality flag
'
' treeview support
'
Public RootNode As Node   'root node of treeview
Public BkNode As Node     'current book node
Public ChpNode As Node    'current chapter node
'
' search options. Keep track of "Any word", "All words", or "Phrase" selection.
'
Public SearchOption As Long
'
' collection for favorites
'
Public colFavs As Collection
'
' History
'
Public colHist As Collection  'storage bin for the history list
Public HistIdx As Long        'a navigation index into the history list
Public HistUpdt As Boolean    'flag used when we are navigating, to prevent history adds
Public colHistory As Collection 'storage of all visited verses in current session
'
' Search Support
'
Public colSrch As Collection
Public LastSearch As String   'last search text
Public ForceSearch As String  'force a search for this text
'
' background mapping
'
Public Background As Long                 'the currently active tiling map index
Public Const bkParch1 As Long = 0         'avoid "magic" numbers
Public Const bkParch2 As Long = 1
Public Const bkParch3 As Long = 2
Public Const bkIce As Long = 3
Public Const bkCloth As Long = 4
Public Const bkRumpled As Long = 5
Public Const bkStucco As Long = 6
Public Const bkMarble As Long = 7
Public Const bkMarbleTX As Long = 8
Public Const bkWood As Long = 9
Public Const bkCustom As Long = 10

Public Const Navy As Long = &H800000      'background for displaying user-edited verses
'
' Bible verse availability flags
'
Public ASVAvail As Boolean                'American Standard Version available
Public KJVAvail As Boolean                'King James Version available
Public MKJVAvail As Boolean               'Modern King James Version available
Public MPVAvail As Boolean                'My Personal Version available
Public RSVAvail As Boolean                'Revised Standard Version available
Public WEBAvail As Boolean                'World English Bible available
Public YLTAvail As Boolean                'Young's Literal Translation available
Public DBYAvail As Boolean                'Darby's Translation Available
Public WBSAvail As Boolean                'Webster's Translation available
Public HIndent As Long                    'hanging indent for note text
'
' Store the App.Path property here.  This is used so that if the application
' is actually running from a CD-ROM, it will still work, using this path as
' a writable-medium path for data that must be stored and updated from softer media.
'
Public AppPath As String
'
' flags for writing chapter/book contents
'
Public IsRTF As Boolean                   'true if writing as an REF file
Public bCancel As Boolean                 'cancel flag storage
Public FileName As String                 'file to write Bible to
Public SaveBible As String                'path to local bible out-file
Public VerseLines As Boolean              'non-zero if each verse should be on its own line
Public BookNewPage As Boolean             'if each book should start on a fresh page
Public AddNoteSpace As Boolean            'add a blank line between each verse for notes
Public CenterBkHeading As Long            'Center book headings
Public CenterChapHeading As Long          'center chapter headings
Public DoBible As Boolean                 'true if building chapters or bible
Public AddPNotes As Boolean               'True if personal Notes to be added
Public PNotesAbove As Boolean             'True if personal notes are to precede verse
Public PNotesColor As Long                'color to print personal notes as
Public IncludeTheoNotes As Boolean        'include theological notes at the end of chapters

Public AutoTimer As Boolean               'if the autotimer is to be used
Public AutoTime As Long                   'the numer of seconds
Public AutoTimeUpd As Long                'modifiable version
Public AutoDirty As Boolean               'set when data changes
Public BackupPath As String               'path for saving the backup

Public ShowVineList As Boolean            'true when Vine List displayed
Public ShowWordList As Boolean            'true when word list displayed
Public ShowKJVDict  As Boolean            'true when KJV dictionary displayed
Public SearchOpen As Boolean              'True when Search Window Open
Public ViewBible As Boolean               'True when the Bible Viewer is open

Public VineTop As Long
Public VineLeft As Long
Public VineWidth As Long
Public VineHeight As Long

Public WordTop As Long
Public WordLeft As Long
Public WordWidth As Long
Public WordHeight As Long

'*******************************************************************************
' Function Name     : InputMsgBox
' Purpose           : Input Box Support
'*******************************************************************************
Public LastInputBox As Long
Public Function InputMsgBox(iForm As Form, _
                            Prompt As String, _
                            Optional PromptCaption As String = vbNullString, _
                            Optional Default As String = vbNullString, _
                            Optional ShowOption As String = vbNullString, _
                            Optional IsVine As Boolean = False) As String
  InputMsgBox = frmInputBox.dwInputBox(iForm, Prompt, PromptCaption, Default, ShowOption, IsVine)
End Function
