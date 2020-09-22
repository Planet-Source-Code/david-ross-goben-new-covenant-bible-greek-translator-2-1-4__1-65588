VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmKJVXlate 
   Caption         =   "Explore the Original KJV Translation Strategy"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7785
   Icon            =   "frmKJVXlate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   7785
   Begin VB.PictureBox picBCV 
      Height          =   315
      Left            =   1680
      ScaleHeight     =   255
      ScaleWidth      =   4035
      TabIndex        =   20
      Top             =   6960
      Width           =   4095
      Begin VB.ComboBox cboBk 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Select Book"
         Top             =   0
         Width           =   1275
      End
      Begin VB.ComboBox cboChp 
         Height          =   315
         Left            =   2355
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Select Chapter"
         Top             =   0
         Width           =   615
      End
      Begin VB.ComboBox cboVrs 
         Height          =   315
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Select Verse"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Height          =   255
         Left            =   3585
         TabIndex        =   21
         ToolTipText     =   "Go to the selected Book, Chapter, and Verse"
         Top             =   0
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custom Select:"
         Height          =   195
         Left            =   0
         TabIndex        =   25
         ToolTipText     =   "Custom select a book, chapter, and verse to view"
         Top             =   60
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< &Prev. Verse"
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      ToolTipText     =   "View previous, sequential verse (Ctrl-Left Arrow)"
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next Verse >"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      ToolTipText     =   "View next, sequential verse (Ctrl-Right Arrow)"
      Top             =   120
      Width           =   1155
   End
   Begin RichTextLib.RichTextBox rtbClip 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6660
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"frmKJVXlate.frx":030A
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "Copy data to the clipboard"
      Top             =   6960
      Width           =   1035
   End
   Begin RichTextLib.RichTextBox txtKeys 
      Height          =   1215
      Left            =   180
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2220
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   2143
      _Version        =   393217
      HideSelection   =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmKJVXlate.frx":0395
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
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   210
      Left            =   0
      TabIndex        =   10
      Top             =   7410
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   370
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFocus 
      Height          =   315
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "txtFocus"
      Top             =   5040
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtbVerse 
      Height          =   1215
      Left            =   180
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   2143
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmKJVXlate.frx":041D
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
   Begin RichTextLib.RichTextBox rtbDef 
      Height          =   1155
      Left            =   180
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5460
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   2037
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmKJVXlate.frx":04A6
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
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6540
      TabIndex        =   0
      Top             =   6960
      Width           =   1035
   End
   Begin RichTextLib.RichTextBox rtbOrder 
      Height          =   1215
      Left            =   180
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3840
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   2143
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmKJVXlate.frx":052D
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Asterisks mark untranslated words, or corresponding Greek words do  not match)"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   6
      Left            =   1860
      TabIndex        =   19
      ToolTipText     =   "Asterisks mark untranslated words, or corresponding Greek words do  not match."
      Top             =   5040
      Width           =   5730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Above KJV Translation shown in actual Greek order:"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   3660
      Width           =   3705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Use the Down or Up arrow keys to navigate through the text lines, or Click them)"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   1860
      TabIndex        =   16
      ToolTipText     =   "Use the Down or Up arrow keys to navigate through the text lines, or Click them."
      Top             =   6660
      Width           =   5715
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Use the Right or Left arrow keys to navigate through words, or Click them)"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   2340
      TabIndex        =   15
      ToolTipText     =   "Use the Right or Left arrow keys to navigate through words, or Click them."
      Top             =   3420
      Width           =   5265
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblWidth"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4440
      TabIndex        =   9
      Top             =   7140
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key KJV Translation:   (click for definition.  * = untranslated words)"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Key KJV Translation: (click words to find definition of Greek word they are translated from.  Asterisks mark untranslated words)"
      Top             =   2040
      Width           =   4635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Word Definitions:"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Original KJV Verse:   (non-bold text are added English words)"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   540
      Width           =   4275
   End
   Begin VB.Label lblVerseID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VerseID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   60
      Width           =   1125
   End
End
Attribute VB_Name = "frmKJVXlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' API stuff
'
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByRECT Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As RECT) As Long
Private Const EM_LINESCROLL = &HB6
Private Const EM_GETRECT = &HB2
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
'
' screen work area variables
'
Private ScrL As Long, ScrW As Long, ScrT As Long, ScrH As Long

Private VsWords() As String, VsIndex() As String, KJVAry() As String
Private WordIndex As Long
Private IsLoading As Boolean
Private SetGoto As Boolean

Private BookDropHandler As clscboFullDrop   'handle fulldrop on book list
Private ChapDropHandler As clscboFullDrop   'handle fulldrop on chapter list
Private VerseDropHandler As clscboFullDrop  'handle fulldrop on verse list

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
    .ToolTipText = "Select Chapter 1 - " & CStr(I)
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


Private Sub cmdGo_Click()
  Bk = Me.cboBk.ListIndex + 1
  Chp = Me.cboChp.ListIndex + 1
  Vrs = Me.cboVrs.ListIndex + 1
  ChpCnt = Me.cboChp.ListCount
  VrsCnt = Me.cboVrs.ListCount
  SetGoto = True
  Call frmGrkXlate.UpdateVerse    'all ok, so display the user selection
  WordIndex = 1
  Call ShowData
  Call PointWord
  Me.cmdGo.Enabled = False
  SetGoto = False
  Me.cmdNext.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Initialize the form
'*******************************************************************************
Private Sub Form_Load()
  Dim Idx As Long
  Dim Ary() As String
  
  IsLoading = True                    'show that the form is loading
  Me.rtbDef.BackColor = cMedium
  Me.rtbVerse.BackColor = cDark
  Me.txtKeys.BackColor = cDark
  Me.rtbOrder.BackColor = cDark
  
  Me.txtFocus.Left = -2440
  GetScreenWorkArea ScrL, ScrW, ScrT, ScrH
  Me.Width = CLng(GetSetting(App.Title, "Settings", "OrgKJVWidth", CStr(Me.Width)))
  Me.Height = CLng(GetSetting(App.Title, "Settings", "OrgKJVHeight", CStr(ScrH)))
  Me.Left = CLng(GetSetting(App.Title, "Settings", "OrgKJVLeft", CStr(ScrW - Me.Width)))
  Me.Top = CLng(GetSetting(App.Title, "Settings", "OrgKJVTop", CStr(ScrT)))
  Me.WindowState = CLng(GetSetting(App.Title, "Settings", "OrgKJVState", "0"))
'
' read the KJV dictionary
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\KJVDict.txt", ForReading, False)
  KJVAry = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' build the book combo list
'
  With Me.picBCV
    .BorderStyle = 0
    .BackColor = cMedium
    Me.cmdGo.Height = .Height
    .Width = Me.cmdGo.Left + Me.cmdGo.Width
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
  
  Set BookDropHandler = New clscboFullDrop
  Set ChapDropHandler = New clscboFullDrop
  Set VerseDropHandler = New clscboFullDrop
'''*** Comment out following 3 lines if you are debugging this form code,
'''*** as otherwise the WndProc handler will hange the VB IDE if you try to STOP
'''*** (not step through, this is OK) this code with this code active
  BookDropHandler.hwnd = Me.cboBk.hwnd
  ChapDropHandler.hwnd = Me.cboChp.hwnd
  VerseDropHandler.hwnd = Me.cboVrs.hwnd
  
  ShowData          'build display information
  WordIndex = 1     'start with first word
  Call PointWord    'point to it

  IsLoading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting App.Title, "Settings", "OrgKJVState", CStr(Me.WindowState)
  If Me.WindowState = vbNormal Then
    SaveSetting App.Title, "Settings", "OrgKJVWidth", CStr(Me.Width)
    SaveSetting App.Title, "Settings", "OrgKJVHeight", CStr(Me.Height)
    SaveSetting App.Title, "Settings", "OrgKJVLeft", CStr(Me.Left)
    SaveSetting App.Title, "Settings", "OrgKJVTop", CStr(Me.Top)
  End If
End Sub

Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
  On Error Resume Next
  Me.txtFocus.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClose_Click
' Purpose           : Close out the form
'*******************************************************************************
Private Sub cmdClose_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopy_Click
' Purpose           : Copy the data to the clipboard
'*******************************************************************************
Private Sub cmdCopy_Click()
  Dim I As Long
  
  With Me.rtbClip
    .Text = Me.lblVerseID.Caption
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelBold = True
    I = Len(.Text)
    
    .SelStart = I
    .SelText = vbCrLf & vbCrLf
    
    .SelStart = Len(.Text)
    .SelText = "Original KJV Verse:" & vbCrLf
    
    .SelStart = Len(.Text)
    .SelRTF = Me.rtbVerse.TextRTF
    .SelStart = Len(.Text)
    .SelText = Underline & vbCrLf & vbCrLf
    
    .SelStart = Len(.Text)
    .SelText = "Key KJV Translation:" & vbCrLf
    .SelRTF = Me.txtKeys.TextRTF
    .SelStart = Len(.Text)
    .SelText = Underline & vbCrLf & vbCrLf
    
    .SelStart = Len(.Text)
    .SelText = "Above KJV Translation shown in actual Greek order:" & vbCrLf
    .SelRTF = Me.rtbOrder.TextRTF
    .SelStart = Len(.Text)
    .SelText = Underline & vbCrLf & vbCrLf
    
    .SelStart = Len(.Text)
    .SelText = "Word Definitions:" & vbCrLf & vbCrLf
    .SelRTF = Me.rtbDef.TextRTF
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    
    Clipboard.Clear
    Clipboard.SetText .Text, vbCFText
    Clipboard.SetText .TextRTF, vbCFRTF
    .Text = vbNullString
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdNext_Click
' Purpose           : Display the next sequential verse
'*******************************************************************************
Private Sub cmdNext_Click()
  frmGrkXlate.Form_KeyDown 39, vbCtrlMask
  WordIndex = 1
  Call ShowData
  Call PointWord
  Me.cmdNext.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdPrevious_Click
' Purpose           : Display the previous sequential verse
'*******************************************************************************
Private Sub cmdPrevious_Click()
  frmGrkXlate.Form_KeyDown 37, vbCtrlMask
  WordIndex = 1
  Call ShowData
  Call PointWord
  Me.cmdPrevious.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyDown
' Purpose           : Check for keyboard navigation keys
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim Ary() As String, S As String
  Dim I As Long, J As Long, K As Long, Idx As Long
  Dim rc As RECT
  
  I = WordIndex
  If Shift = vbCtrlMask Then
    Select Case KeyCode
      Case 39           'right arrow
        Me.cmdNext.Value = True
      Case 37           'left arrow
        Me.cmdPrevious.Value = True
    End Select
    Exit Sub
  ElseIf Shift <> 0 Then
    Exit Sub
  End If
  Select Case KeyCode
    Case 39           'right arrow
      I = I + 1
    Case 37           'left arrow
      I = I - 1
    Case 38           'up arrow
      Call SendMessageByNum(Me.rtbDef.hwnd, EM_LINESCROLL, 0&, -1&)
      Exit Sub
    Case 40           'down arrow
      Call SendMessageByNum(Me.rtbDef.hwnd, EM_LINESCROLL, 0&, 1&)
      Exit Sub
    Case 33           'page down
      Call SendMessageByRECT(Me.rtbDef.hwnd, EM_GETRECT, 0&, rc)
      I = (rc.Bottom - rc.Top) * 15 \ (FntSize * -20)   'get lines per window
      Call SendMessageByNum(Me.rtbDef.hwnd, EM_LINESCROLL, 0&, I)
      Exit Sub
    Case 34           'page up
      Call SendMessageByRECT(Me.rtbDef.hwnd, EM_GETRECT, 0&, rc)
      I = (rc.Bottom - rc.Top) * 15 \ (FntSize * 20)   'get lines per window
      Call SendMessageByNum(Me.rtbDef.hwnd, EM_LINESCROLL, 0&, I)
      Exit Sub
    Case Else
      Exit Sub
  End Select
'
' scroll horizontally in the key list to the desired word
'
  Ary = Split(Me.txtKeys.Text, " ")
  If I < 1 Then I = UBound(Ary) + 1
  If I > UBound(Ary) + 1 Then I = 1
  WordIndex = I
  Call PointWord  'now point to the corresponding word definition list
End Sub

'*******************************************************************************
' Subroutine Name   : ShowData
' Purpose           : Fill the form with data for the current verse
'*******************************************************************************
Public Sub ShowData()
  Dim S As String, VsText As String, sAry() As String, tAry() As String, Ary() As String
  Dim T As String, OT As String
  Dim I As Long, J As Long, K As Long
  
'
' obtain a title for this entry in the format "Book, chapter:verse"
'
  Me.rtbDef.Text = vbNullString
  Screen.MousePointer = vbHourglass
  DoEvents
  Me.Enabled = False
  Ary = Split(Books(Bk), ",")  'grab the book information
  Me.lblVerseID.Caption = Ary(3) & " " & CStr(Chp) & ":" & CStr(Vrs)
'
' read the KJV bible
'
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\KJV.txt", ForReading, False)
  Ary = Split(ts.ReadAll, vbCrLf)
  ts.Close
  VsText = Mid$(Ary(VrsIdx), 8)
'
' read the KJV translation verse index
'
  VsIndex = Split(Mid$(KJVidxAry(VrsIdx), 8), ",")
  sAry = Split(Mid$(KJVidxAry(VrsIdx), 8), ",")
'
' read the KJV translation word list
'
  VsWords = Split(Mid$(KJVwrdAry(VrsIdx), 8), ",")
'
' strip the punctuation from a copy of the current verse
'
  S = VsText
  For I = 1 To Len(S)
    Select Case Mid$(S, I, 1)
      Case "A" To "Z", "a" To "z"
      Case Else
        Mid$(S, I, 1) = " "
    End Select
  Next I
  S = " " & S & " "
'
' mark words that correspond in the original verse text
'
  For I = 0 To UBound(VsWords)
    If VsWords(I) <> "*" Then
      T = VsWords(I)
      J = InStr(1, T, "'")
      If J Then Mid$(T, J, 1) = " "
      J = InStr(1, T, "-")
      If J Then Mid$(T, J, 1) = " "
      J = InStr(1, S, " " & T & " ")
      S = Left$(S, J) & "[" & VsWords(I) & "]" & Mid$(S, J + Len(T) + 1)
      VsText = Left$(VsText, J - 1) & "[" & VsWords(I) & "]" & Mid$(VsText, J + Len(T))
    End If
    OT = OT & " " & VsWords(I)
  Next I
'
' show the original verse, marked with found words
'
  With Me.rtbVerse
    .Text = vbNullString
    J = 0
    Do While Len(VsText) <> 0
      I = InStr(1, VsText, "[")
      If I = 0 Then
        S = VsText
        VsText = vbNullString
        J = Len(.Text)
      Else
        S = Left$(VsText, I - 1)
        If Len(S) <> 0 Then
          .SelStart = J
          .SelText = S
          .SelStart = J
          .SelLength = Len(S)
          .SelBold = False
          .SelColor = vbBlack
          J = Len(.Text)
        End If
        VsText = Mid$(VsText, I + 1)
        I = InStr(1, VsText, "]")
        S = Left$(VsText, I - 1)
      End If
      .SelStart = J
      .SelText = S
      .SelStart = J
      .SelLength = Len(S)
      .SelBold = True
      .SelColor = vbBlue
      VsText = Mid$(VsText, I + 1)
      J = Len(.Text)
    Loop
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelLength = 0
  End With
'
' display the KJV keys text
'
  With Me.txtKeys
    .Text = Trim$(OT)
    If Len(Grk(VrsIdx)) > 7 Then
      I = vbBlue
    Else
      I = vbMagenta
    End If
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelColor = I
    .SelFontSize = FntSize
    .SelLength = 0
  End With
'
' build the text hsoing usage in the source Greek order of wording
'
  S = Mid$(GrkBBL(VrsIdx), 8)
  If Len(S) = 0 Then
    Me.rtbOrder.Text = vbNullString
  Else
    Ary = Split(S, " ")
    ReDim tAry(UBound(Ary)) As String
    For I = 0 To UBound(Ary)
      tAry(I) = "*"
      For J = 0 To UBound(sAry)
        If Len(sAry(J)) <> 0 Then
          If sAry(J) = Ary(I) Then
            tAry(I) = VsWords(J)
            sAry(J) = vbNullString
            Exit For
          End If
        End If
      Next J
    Next I
    With Me.rtbOrder
      .Text = Join(tAry, " ")
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelFontSize = FntSize
      .SelStart = 0
    End With
  End If
'
' show the definition of all Greek words used
'
  ShowDef
  
  Me.Enabled = True
  Me.cboBk.ListIndex = Bk - 1
  Me.cboChp.ListIndex = Chp - 1
  Me.cboVrs.ListIndex = Vrs - 1
  Me.cmdGo.Enabled = False
  Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize the form
'*******************************************************************************
Private Sub Form_Resize()
  Dim I As Long
  Static Resizing As Boolean
  
  If Resizing Then Exit Sub
  If Me.WindowState = vbMinimized Then Exit Sub
  Resizing = True
  
  On Error Resume Next
  If Me.Height < 8100 Then Me.Height = 8100
  If Me.Width < 6300 Then Me.Width = 6500
  
  If Me.Width > ScrW Then Me.Width = ScrW
  If Me.Height > ScrH Then Me.Height = ScrH
  If Me.Top < ScrT Then Me.Top = ScrT
  If Me.Left < ScrL Then Me.Left = ScrL
  
  If Me.Top + Me.Height > ScrT + ScrH Then Me.Top = ScrH - Me.Height
  If Me.Left + Me.Width > ScrW Then Me.Left = ScrW - Me.Width
  On Error GoTo 0
  
  Me.rtbVerse.Width = Me.ScaleWidth - Me.rtbVerse.Left * 2
  Me.txtKeys.Width = Me.rtbVerse.Width
  Me.rtbDef.Width = Me.rtbVerse.Width
  Me.rtbOrder.Width = Me.rtbVerse.Width
  Me.cmdClose.Left = Me.ScaleWidth - Me.cmdClose.Width - Me.rtbVerse.Left
  Me.cmdClose.Top = Me.ScaleHeight - Me.cmdClose.Height - 120 - Me.StatusBar1.Height
  Me.cmdCopy.Top = Me.cmdClose.Top
  Me.Label1(4).Top = Me.cmdClose.Top - Me.Label1(4).Height - 60
  Me.Label1(4).Left = Me.ScaleWidth - Me.Label1(4).Width - Me.rtbVerse.Left
  Me.Label1(3).Left = Me.ScaleWidth - Me.Label1(3).Width - Me.rtbVerse.Left
  Me.Label1(6).Left = Me.ScaleWidth - Me.Label1(6).Width - Me.rtbVerse.Left
  Me.rtbDef.Height = Me.Label1(4).Top - Me.rtbDef.Top
  Me.cmdNext.Left = Me.ScaleWidth - Me.cmdNext.Width - Me.rtbVerse.Left
  Me.cmdPrevious.Left = Me.cmdNext.Left - Me.cmdPrevious.Width - 120
  
  Me.picBCV.Top = Me.cmdClose.Top
  I = Me.cmdCopy.Left + Me.cmdCopy.Width
  Me.picBCV.Left = (Me.cmdClose.Left - I - Me.picBCV.Width) / 2 + I
  
  If Not IsLoading Then
    SaveSetting App.Title, "Settings", "OrgKJVState", CStr(Me.WindowState)
    If Me.WindowState = vbNormal Then
      SaveSetting App.Title, "Settings", "OrgKJVWidth", CStr(Me.Width)
      SaveSetting App.Title, "Settings", "OrgKJVHeight", CStr(Me.Height)
      SaveSetting App.Title, "Settings", "OrgKJVLeft", CStr(Me.Left)
      SaveSetting App.Title, "Settings", "OrgKJVTop", CStr(Me.Top)
    End If
  End If
  
  Resizing = False
End Sub

'*******************************************************************************
' Subroutine Name   : rtbDef_GotFocus
' Purpose           : Avoid displaying the cursor in the text controls
'*******************************************************************************
Private Sub rtbDef_GotFocus()
  Me.txtFocus.SetFocus
End Sub

Private Sub rtbOrder_GotFocus()
  Me.txtFocus.SetFocus
End Sub

Private Sub rtbVerse_GotFocus()
  Me.txtFocus.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : txtKeys_Click
' Purpose           : Highlight the selected work and update indexing for it
'*******************************************************************************
Private Sub txtKeys_Click()
  Dim I As Long, K As Long, SL As Long, Idx As Long, J As Long
  Dim S As String, Ary() As String
  
  S = Me.txtKeys.Text                           'grab text
  With Me.txtKeys
    If CBool(.SelLength) Then Exit Sub
    SL = .SelStart + 2                  'get cursor position
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
    .SelStart = K   'ensure selection set
    .SelLength = I - K - 1
    
    I = 1
    S = Me.txtKeys.Text
    Do While K <> 0
      If Mid$(S, K, 1) = " " Then I = I + 1
      K = K - 1
    Loop
    WordIndex = I
    
    K = 0
    S = Me.rtbDef.Text
    Do While I > 0
      K = InStr(K + 1, S, "KJV word:")
      I = I - 1
    Loop
    I = InStr(K, S, vbCrLf)
  End With
  With Me.rtbDef
    .SelStart = Len(S)
    .SelStart = K - 1
    .SelLength = I - K + 1
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : PointWord
' Purpose           : Point to a word in the key list, and update the definition control
'*******************************************************************************
Private Sub PointWord()
  Dim S As String
  Dim I As Long, J As String, K As Long
  
  With Me.txtKeys
    I = 0
    J = WordIndex - 1
    S = .Text & " "
    Do While J > 0
      I = I + 1
      If Mid$(S, I, 1) = " " Then J = J - 1
    Loop
    
    J = I + 1
    Do
      J = J + 1
    Loop Until Mid$(S, J, 1) = " "
    .SelStart = I   'ensure selection set
    .SelLength = J - I - 1
    
    I = WordIndex
    K = 0
    S = Me.rtbDef.Text
    Do While I > 0
      K = InStr(K + 1, S, "KJV word:")
      I = I - 1
    Loop
    I = InStr(K, S, vbCrLf)
  End With
  With Me.rtbDef
    .SelStart = Len(S)
    .SelStart = K - 1
    .SelLength = I - K + 1
  End With
End Sub

Private Sub txtKeys_GotFocus()
  Me.txtFocus.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : ShowDef
' Purpose           : Display the list of key term definitions
'*******************************************************************************
Private Sub ShowDef()
  Dim Idx As Long, I As Long, SL As Long, SS As Long, J As Long
  Dim S As String, Ary() As String, T As String, TT As String
  
  frmGrkXlate.InitMerge                            'initialize the merge text data
'
' now process each Greek word
'
  For Idx = 0 To UBound(VsWords)
    I = CLng(VsIndex(Idx))
    Ary = Split(DefRef(I), vbTab)
    With frmGrkXlate
      SL = Len(.rtbMerge.Text)
      .AddMergeCrlf "KJV word:  '" & VsWords(Idx) & "'", , FntSize, True
      With .rtbMerge
        SS = Len(.Text)
        .SelStart = SL
        .SelLength = SS - SL + 1
        .SelColor = vbBlue
      End With
      .AddMerge "(" & Ary(1) & ")    ", "Symbol", FntSize, True
      .AddMerge Ary(2), , FntSize, True, True
      S = "    (" & Ary(3) & ")    Strong's Reference # " & Ary(5) & vbCrLf
      If Len(Ary(4)) <> 0 Then S = S & Ary(4) & vbCrLf
      .AddMergeCrlf S, , FntSize
      '
      ' display the description contents
      '
      S = Ary(6)
      I = InStr(1, S, "\")                          'conver "\" to vbCrLf
      Do While I
        S = Left$(S, I - 1) & vbCrLf & Mid$(S, I + 1)
        I = InStr(I + 2, S, "\")
      Loop
'
' process list and transcribe any Greek words
'
      J = InStr(1, S, "{")
      Do While J <> 0
        .AddMerge Left$(S, J - 1), , FntSize
        I = InStr(J + 1, S, "}")
        If I = 0 Then Exit Do                 'no match
        TT = Mid$(S, J + 1, I - J - 1)
        .AddMerge TT, "Symbol", FntSize, True
        S = Mid$(S, I + 1)
        J = InStr(1, S, "{")
      Loop
      If Len(S) <> 0 Then
        .AddMerge S, , FntSize
      End If
      
      .AddMergeCrlf vbCrLf & vbCrLf & "Current Synonyms:", , FntSize   'prepart for synonyms
      Ary = Split(WordRef(CLng(Ary(5))), vbTab)     'grab list of words from word list
      Ary = Split(Ary(2), ",")                      'grab words
      .AddMergeCrlf Join(Ary, ", "), , FntSize, , True 'sent it to merge data
      
      S = VsWords(Idx) & vbTab
      I = Len(S)
      For J = 1 To UBound(KJVAry) - 1
        If StrComp(Left$(KJVAry(J), I), S, vbTextCompare) = 0 Then
          .AddMerge vbCrLf & "Modern meaning of KJV word (", , FntSize, , True
          .AddMerge VsWords(Idx), , FntSize, True, True
          .AddMerge "):  ", , FntSize, , True
          .AddMergeCrlf Mid$(KJVAry(J), I + 1) & ".", , FntSize, True, True
          Exit For
        End If
      Next J
      '
      ' if we just processed the last entry, then add a separator
      '
      .AddMergeSep
    End With
  Next Idx
'
' set the indenting
'
  SL = Len(Me.rtbDef.Text)
  With Me.lblWidth
    .FontSize = FntSize
    .Caption = String$(7, 32)
    frmGrkXlate.SetIndent 0, SL, .Width                         'set hanging indent
  End With
  With Me.rtbDef
    LockWindowUpdate .hwnd
    .TextRTF = frmGrkXlate.rtbMerge.TextRTF
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelFontSize = FntSize
    .SelLength = 0
  End With
  LockWindowUpdate 0
  frmGrkXlate.InitMerge                                         'then clear it
End Sub

'*******************************************************************************
' Subroutine Name   : SetVerseCount
' Purpose           : Rebuild the verse list for the current chapter
'*******************************************************************************
Private Sub SetVerseCount()
  Dim I As Long
  Dim S As String
  
  S = Format$(Me.cboBk.ListIndex + 1, "00") & Format$(Me.cboChp.ListIndex + 1, "00")
  I = FindExactMatch(frmGrkXlate.lstGrk, S & "01")   'find verse 1
  With Me.cboVrs
    .Clear                                  'remove any old data
    Do
      .AddItem CStr(.ListCount + 1)         'add an item
    Loop While Left$(frmGrkXlate.lstGrk.List(I + .ListCount), 4) = S
    .ToolTipText = "Select Verse 1 - " & CStr(.ListCount)
    .ListIndex = 0                          'force verse 1 (-1)
  End With
End Sub

