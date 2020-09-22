VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmView 
   Caption         =   "View Bible"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "frmView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   11010
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   6900
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6780
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   9060
      Top             =   6780
   End
   Begin VB.CommandButton cmdFindInText 
      Height          =   315
      Left            =   6240
      Picture         =   "frmView.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Search for a word or phrase in the text below from the current cursor position (Ctrl-F)"
      Top             =   60
      Width           =   375
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
      Height          =   315
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Find next match in text (F3)"
      Top             =   60
      Width           =   375
   End
   Begin VB.ComboBox cboVerse 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Go to and highlight the selected V of the current Book and Chapter"
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CheckBox chkEnableEdit 
      Caption         =   "&Enable Editing"
      Height          =   315
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Click to enable editing to add personal notes"
      Top             =   60
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Enabled         =   0   'False
      Height          =   315
      Left            =   9120
      TabIndex        =   10
      ToolTipText     =   "Save modifications"
      Top             =   60
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cboChapter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Select a Chapter of the Book, and go to the first verse of the chapter"
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ComboBox cboBook 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmView.frx":0E54
      Left            =   720
      List            =   "frmView.frx":0E56
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select a Book and go to the first chapter of this book"
      Top             =   60
      Visible         =   0   'False
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6315
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11139
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmView.frx":0E58
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6960
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   11853
            Text            =   "If a verse does not actually exist in the Greek text, it is displayed in a really obnoxious magenta"
            TextSave        =   "If a verse does not actually exist in the Greek text, it is displayed in a really obnoxious magenta"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   7011
            Text            =   "Bible Base"
            TextSave        =   "Bible Base"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblVerse 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Verse:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblChapter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Chapter:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblBook 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Book:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' API Stuff
'
Private Const CB_SHOWDROPDOWN = &H14F
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
' local stuff
'
Private Const MyPersonalNotes As String = "My Personal Notes:"
Private Unloading As Boolean        'true when unloading form
Private BuildVerse As Boolean       'true when building the verse combo list
Private IsTxt As Boolean            'true if the file was TXT
Private UserBible As String         'name of Bible this text based upon
Private LastSch As String           'last search text

Private BookDropHandler As clscboFullDrop   'handle fulldrop on book list
Private ChapDropHandler As clscboFullDrop   'handle fulldrop on chapter list
Private VerseDropHandler As clscboFullDrop  'handle fulldrop on verse list

'*******************************************************************************
' Subroutine Name   : cmdClose_Click
' Purpose           : Close form
'*******************************************************************************
Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Idx As Long
  Dim Ary() As String, Txt As String

  Me.cmdClose.Left = -2440
'
' get user bible
'
  ShowingBible = True
  frmGrkXlate.mnuFileCreateBible.Enabled = False
  Me.RichTextBox1.BackColor = cVLight
  Me.cmdFindInText.BackColor = cLight
  Me.cmdFindNext.BackColor = cLight
  Me.cmdFindNext.Enabled = False
  Me.cmdFindInText.Visible = False
  Me.cmdFindNext.Visible = False
'
' show business
'
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  Me.Top = CLng(GetSetting(App.Title, "Settings", "BBLViewerT", "0"))
  Me.Left = CLng(GetSetting(App.Title, "Settings", "BBLViewerL", "0"))
  Me.Width = CLng(GetSetting(App.Title, "Settings", "BBLViewerW", CStr(Me.Width)))
  Me.Height = CLng(GetSetting(App.Title, "Settings", "BBLViewerH", CStr(Me.Height)))
  Me.WindowState = CLng(GetSetting(App.Title, "Settings", "BBLViewer", "2"))
  Me.Show
  DoEvents
'
' get user bible
'
  UserBible = GetSetting(App.Title, "Settings", "SaveBible", vbNullString)
  Txt = GetSetting(App.Title, "Settings", "BibleBase", vbNullString)
  If Len(Txt) <> 0 Then Txt = "Bible based upon: " & Txt
  Me.StatusBar1.Panels(2).Text = Txt
'
' add class to force comboboxes to do a full drop
'
  Set BookDropHandler = New clscboFullDrop
  Set ChapDropHandler = New clscboFullDrop
  Set VerseDropHandler = New clscboFullDrop
  '*** Comment out following 3 lines if you are debugging this form code,
  '*** as otherwise the WndProc handler will hange the VB IDE if you try to STOP
  '*** (not step through, this is OK) this code with this code active
  BookDropHandler.hwnd = Me.cboBook.hwnd
  ChapDropHandler.hwnd = Me.cboChapter.hwnd
  VerseDropHandler.hwnd = Me.cboVerse.hwnd
'
' load the bible file
'
  If Len(UserBible) <> 0 Then
    If Not Fso.FileExists(UserBible) Then UserBible = vbNullString
  End If
  If Len(UserBible) = 0 Then
    UserBible = AddSlash(App.Path) & "\MyBible.rtf"
    If Not Fso.FileExists(UserBible) Then UserBible = vbNullString
  End If
  If Len(UserBible) = 0 Then
    UserBible = AddSlash(App.Path) & "\MyBible.txt"
    If Not Fso.FileExists(UserBible) Then UserBible = vbNullString
  End If
  
  If Len(UserBible) <> 0 Then
    DoEvents
    On Error Resume Next
    Me.RichTextBox1.LoadFile UserBible, rtfRTF    'first try as RTF
    If Err.Number <> 0 Then
      IsTxt = True
      Err.Clear
      Me.RichTextBox1.LoadFile UserBible, rtfText 'if that fails, try flat text
    End If
    If Err.Number <> 0 Then UserBible = vbNullString
    If Len(UserBible) = 0 Then
      Screen.MousePointer = vbDefault
      Me.RichTextBox1.Text = vbNullString
      MessageBox Me, "Cannot seem to load the Bible file. It may be corrupted.", _
                 vbOKOnly Or vbExclamation, "Cannot Open Bible File"
      Me.Timer1.Enabled = True  'use the timer to prevent parent processing error
      Exit Sub
    End If
  Else
    Screen.MousePointer = vbDefault
    Me.RichTextBox1.Text = vbNullString
    MessageBox Me, "Cannot seem to find the Bible file. It may not exist." & vbCrLf & _
                   "You can create it from the 'Write complete Bible file' " & vbCrLf & _
                   "option under the 'Bible' menu, in the main program.", _
               vbOKOnly Or vbExclamation, "Cannot Open Bible File"
    Me.Timer1.Enabled = True    'use the timer to prevent parent processing error
    Exit Sub
  End If
  On Error GoTo 0
'
' show where the file is located
'
  Me.Caption = Me.Caption & " - " & UserBible
'
' build the book combo list
'
  For Idx = 1 To 27
    Ary = Split(Books(Idx), ",")
    Me.cboBook.AddItem Ary(3)
  Next Idx
'
' select Matthew (book1 -1) and force a rebuild of the chapter list
'
  Me.cboBook.ListIndex = 0
'
' disable the saving and editing options for now
'
  Me.cmdSave.Enabled = False
  Me.chkEnableEdit.Value = vbUnchecked
  Me.chkEnableEdit.BackColor = cMedium
'
' give the user some options
'
  Me.lblBook.Visible = True
  Me.lblChapter.Visible = True
  Me.lblVerse.Visible = True
  Me.cboBook.Visible = True
  Me.cboChapter.Visible = True
  Me.cboVerse.Visible = True
  Me.chkEnableEdit.Visible = True
  Me.cmdSave.Visible = True
  Me.cmdFindInText.Visible = True
  Me.cmdFindNext.Visible = True
'
' no longer busy
'
  Screen.MousePointer = vbDefault
  Me.Enabled = True
  Me.RichTextBox1.SetFocus
  Me.Show
  frmGrkXlate.mnuWinBible.Enabled = True
  frmGrkXlate.CheckWin
  ViewBible = True
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Retile the background as needed
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
End Sub

'*******************************************************************************
' Subroutine Name   : Form_QueryUnload
' Purpose           : If the file is dirty, allow the user the option to save it
'*******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then
    Unloading = True
    If Me.cmdSave.Enabled Then
      bCancel = False
      Select Case MessageBox(Me, "This file has changed. Save these changes?", _
                  vbYesNoCancel Or vbQuestion, "Save Changes")
        Case vbCancel
          Cancel = 1
          Exit Sub
        Case vbYes
          Me.cmdSave.Value = True
          If bCancel Then Cancel = 1
      End Select
    End If
  End If
  Unloading = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize the form
'*******************************************************************************
Private Sub Form_Resize()
  Static Resizing As Boolean
  
  If Me.WindowState = vbMinimized Then
    frmGrkXlate.ZOrder 0
    frmGrkXlate.SetFocus
    Exit Sub
  End If
  If Resizing Then Exit Sub
  Resizing = True
  If Me.Width < 10440 Then Me.Width = 10440
  If Me.Height < 4000 Then Me.Height = 4000
  Me.RichTextBox1.Width = Me.ScaleWidth - Me.RichTextBox1.Left * 2
  Me.RichTextBox1.Height = Me.ScaleHeight - Me.StatusBar1.Height - Me.RichTextBox1.Top - 120
  Me.cmdSave.Left = Me.RichTextBox1.Left + Me.RichTextBox1.Width - Me.cmdSave.Width
  Me.chkEnableEdit.Left = Me.cmdSave.Left - Me.chkEnableEdit.Width - 120
  Resizing = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Release allocated resources
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set VerseDropHandler = Nothing
  Set ChapDropHandler = Nothing
  Set BookDropHandler = Nothing
  ShowingBible = False
  frmGrkXlate.mnuFileCreateBible.Enabled = True
  frmGrkXlate.mnuWinBible.Enabled = False
  frmGrkXlate.CheckWin
  ViewBible = False
  If Me.WindowState <> vbMinimized Then
    SaveSetting App.Title, "Settings", "BBLViewer", CStr(Me.WindowState)
    If Me.WindowState = vbNormal Then
      SaveSetting App.Title, "Settings", "BBLViewerW", CStr(Me.Width)
      SaveSetting App.Title, "Settings", "BBLViewerH", CStr(Me.Height)
      SaveSetting App.Title, "Settings", "BBLViewerL", CStr(Me.Left)
      SaveSetting App.Title, "Settings", "BBLViewerT", CStr(Me.Top)
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyDown
' Purpose           : Keyboard support of Find and Find Next buttons
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = vbCtrlMask Then        'CTRL key?
    If KeyCode = vbKeyF Then        'and F key?
      KeyCode = 0                   'disable further processing on it
      Me.cmdFindInText.Value = True 'process FIND command
    End If
  ElseIf Shift = 0 Then             'no special keys
    If KeyCode = vbKeyF3 Then       'F3 key?
      If Me.cmdFindNext.Enabled Then  'can do Find Next?
        KeyCode = 0                 'yes, so disable further key processing
        Me.cmdFindNext.Value = True 'and find the next match
      End If
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cboBook_Click
' Purpose           : A new book selected, rebuild the chapter list
'*******************************************************************************
Private Sub cboBook_Click()
  Dim Idx As Long, Cnt As Long
  Dim Ary() As String
  
  If BuildVerse Then Exit Sub     'do not process if the verse list is being built
  
  Ary = Split(Books(Me.cboBook.ListIndex + 1), ",") 'get selected book data
  Me.cboChapter.Clear                               'init combo for new data
  For Idx = 1 To CLng(Ary(4))
    Me.cboChapter.AddItem CStr(Idx)                 'build new data
  Next Idx
  Me.cboChapter.ListIndex = 0       'select chapter 1 (-1), force verse list rebuild
End Sub

'*******************************************************************************
' Subroutine Name   : cboChapter_Click
' Purpose           : A chapter selected, find it in the text
'*******************************************************************************
Private Sub cboChapter_Click()
  Dim S As String, Txt As String
  Dim Idx As Long
  
  If BuildVerse Then Exit Sub     'do not process if the verse list is being built
  SetVerseCount                   'build a verse list for this chapter
  DoEvents
  With Me.RichTextBox1
    LockWindowUpdate .hwnd
    S = Me.cboBook.Text & ", Chapter " & Me.cboChapter.Text
    Txt = .Text
    Idx = InStr(1, Txt, S)        'find it
    If Me.cboChapter.Text = "1" Then  'if chapter 1, then find book heading
      Idx = InStrRev(Txt, UCase$(Me.cboBook.Text), Idx - 1)
    End If
    .SelStart = Len(Txt)          'go to end of text
    .SelStart = Idx - 1           'slide back up to find point
    LockWindowUpdate 0
    On Error Resume Next          'trap error on initial load (before form Paint)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cboVerse_Click
' Purpose           : Go to and highlight the selected verse
'*******************************************************************************
Private Sub cboVerse_Click()
  Dim S As String, Txt As String, T As String, TT As String
  Dim Idx As Long, I As Long, J As Long, K As Long
  
  If BuildVerse Then Exit Sub     'do not process if the verse list is being built
'
' first duplicate search for the chapter
'
  DoEvents
  With Me.RichTextBox1
    LockWindowUpdate .hwnd        'prevent control flash
    S = Me.cboBook.Text & ", Chapter " & Me.cboChapter.Text
    Txt = .Text
    Idx = InStr(1, Txt, S)        'find it
    .SelStart = Len(Txt)          'first go to the end
    .SelStart = Idx               'then "scroll" up to it
'
' now seek verse data
'
    T = "(" & CStr(Me.cboVerse.ListIndex + 1) & ")" 'verse to find
    TT = "(" & CStr(Me.cboVerse.ListIndex + 2) & ")" 'hopeful end of verse
    Idx = InStr(Idx, Txt, T)      'find verse
    I = InStr(Idx + Len(T), Txt, TT) 'find end of verse
    J = InStr(Idx + Len(T), Txt, vbCrLf) 'find end of paragraph
    If J = 0 Then J = Len(Txt) + 1  'if for some wierd reason no terminating CRLF
    K = InStr(Idx + Len(T), Txt, MyPersonalNotes) 'find any note data
    If K = 0 Then K = Len(Txt) + 1
    If I = 0 Then                 'if next verse not found
      I = J                       'use end of paragraph
    Else
      If J < I Then I = J         'back up to paragraph mark
    End If
    If K < I Then I = K           'back up to note data
    .SelStart = Idx - 1           'slide back up to find point
    .SelLength = I - Idx
    Txt = .SelText                'remove any trailing spaces
    .SelLength = Len(RTrim$(Txt))
    LockWindowUpdate 0            'released locked control
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : SetVerseCount
' Purpose           : Rebuild the verse list for the current chapter
'*******************************************************************************
Private Sub SetVerseCount()
  Dim I As Long
  Dim S As String
  
  S = Format$(Me.cboBook.ListIndex + 1, "00") & Format$(Me.cboChapter.ListIndex + 1, "00")
  I = FindExactMatch(frmGrkXlate.lstGrk, S & "01")   'find verse 1
  With Me.cboVerse
    .Clear                                  'remove any old data
    Do
      .AddItem CStr(.ListCount + 1)         'add an item
    Loop While Left$(frmGrkXlate.lstGrk.List(I + .ListCount), 4) = S
    BuildVerse = True                       'prevent "GO TO" on verse this time
    .ListIndex = 0                          'force verse 1 (-1)
    BuildVerse = False
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdFindInText_Click
' Purpose           : Find text in the Notes panel
'*******************************************************************************
Private Sub cmdFindInText_Click()
  Dim Sch As String, Text As String
  Dim Idx As Long
  
  Sch = InputMsgBox(Me, "Enter word or phrase to find:", "Search For Text", LastSch)
  If Len(Sch) = 0 Then Exit Sub
  LastSch = Sch
  With Me.RichTextBox1
    Text = .Text
    Idx = .SelStart + .SelLength + 1
    Idx = InStr(Idx, Text, Sch, vbTextCompare)
    If Idx <> 0 Then
      .SelStart = Idx - 1
      .SelLength = Len(Sch)
      Call FindBkChVs(Text, Idx)
      Me.cmdFindNext.Enabled = True
      Me.cmdFindNext.SetFocus
    Else
      Me.cmdFindNext.Enabled = False
      MessageBox Me, "Search text not found: " & Sch, _
                 vbOKOnly Or vbExclamation, "Text Not Found"
      Me.RichTextBox1.SetFocus
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
  With Me.RichTextBox1
    Text = .Text
    Idx = .SelStart + .SelLength + 1
    Idx = InStr(Idx, Text, LastSch, vbTextCompare)
    If Idx <> 0 Then
      .SelStart = Idx - 1
      .SelLength = Len(LastSch)
      Call FindBkChVs(Text, Idx)
      Me.cmdFindNext.Enabled = True
    Else
      Me.cmdFindNext.Enabled = False
      MessageBox Me, "Search text not found: " & LastSch, vbOKOnly Or vbExclamation, "Text Not Found"
      Me.RichTextBox1.SetFocus
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : FindBkChVs
' Purpose           : Support routine that locates the sought book/chapter/verse
'*******************************************************************************
Private Sub FindBkChVs(Text As String, ByVal Index As Long)
  Dim Idx As Long, B As Long, C As Long, V As Long, Ccnt As Long, Vcnt As Long
  Dim oIdx As Long, I As Long
  Dim BCV As String, Ary() As String, S As String, T As String
  
  BuildVerse = True                       'prevent "GO TO" on verse this time
  oIdx = 1
  For B = 1 To 27
    Ary = Split(Books(B), ",")            'get book information
    S = Ary(3) & ", Chapter "             'get title
    Ccnt = CLng(Ary(4))                   'get chapter count
    For C = 1 To Ccnt
      T = S & CStr(C)
      Idx = InStr(oIdx, Text, T)
      If Idx > Index Then
        Me.cboBook.ListIndex = B - 1
'
' rebuild chapter list
'
        With Me.cboChapter
          .Clear                          'remove old list
          For I = 1 To Ccnt
            .AddItem CStr(I)              'build new list
          Next I
        End With
        
        If C > 1 Then
          Me.cboChapter.ListIndex = C - 2
        Else
          Me.cboChapter.ListIndex = 0
        End If
'
' rebuild verse count
'
        With Me.cboVerse
          .Clear                                    'clear old list
          Vcnt = 0                                  'init to 0 verses
          T = Format$(Me.cboBook.ListIndex + 1, "00") & _
              Format$(Me.cboChapter.ListIndex + 1, "00") 'selected book and chapter
          I = FindExactMatch(frmGrkXlate.lstGrk, T & "01")   'find verse 1
          Do
            Vcnt = Vcnt + 1                         'find consecutive verse
            .AddItem CStr(Vcnt)                     'add item to list
          Loop While Left$(Grk(I + Vcnt), 4) = T
        End With
'
' find target Verse
'
        Idx = oIdx
        For V = 1 To Vcnt
          T = "(" & CStr(V) & ") "
          Idx = InStr(Idx, Text, T)
          If Idx > Index Then
            If V > 1 Then
              Me.cboVerse.ListIndex = V - 2
            Else
              Me.cboVerse.ListIndex = 0
            End If
            Exit For
          End If
        Next V
        BuildVerse = False
        Exit Sub
      End If
      oIdx = Idx
    Next C
  Next B
  BuildVerse = False
End Sub

'*******************************************************************************
' Subroutine Name   : chkEnableEdit_Click
' Purpose           : Toggle enabling editing in this text
'*******************************************************************************
Private Sub chkEnableEdit_Click()
  If Me.chkEnableEdit.Value = vbChecked Then      'enabling editing
    Me.RichTextBox1.Locked = False
    Me.RichTextBox1.BackColor = vbWhite
    Me.chkEnableEdit.Caption = "&Disable Editing"
    Me.chkEnableEdit.ToolTipText = "Click to disable editing for adding personal notes"
  Else                                            'disable editing
    Me.RichTextBox1.Locked = True
    Me.RichTextBox1.BackColor = cVLight
    Me.chkEnableEdit.Caption = "&Enable Editing"
    Me.chkEnableEdit.ToolTipText = "Click to enable editing to add personal notes"
  End If
  Me.RichTextBox1.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSave_Click
' Purpose           : Save changes to teh file
'*******************************************************************************
Private Sub cmdSave_Click()
  On Error Resume Next
  If IsTxt Then                                   'was it loaded as text?
    Me.RichTextBox1.SaveFile UserBible, rtfText   'yes, so save as text
  Else
    Me.RichTextBox1.SaveFile UserBible, rtfRTF    'else Rich Text Format
  End If
  If Err.Number <> 0 Then                         'did we stub our toe?
    If Unloading Then                             'are we unloading?
      If MessageBox(Me, "Cannot save the file. " & _
                        "It may be read-only or it is opened by another application.", _
                        vbOKCancel Or vbQuestion, _
                        "Cannot Update the File") = vbCancel Then
        bCancel = True
        Exit Sub
      End If
    Else
      MessageBox Me, "Cannot save the file. " & _
                     "It may be read-only or it is opened by another application.", _
                     vbOKOnly Or vbExclamation, _
                     "Cannot Update the File"
    End If
  End If
  Me.cmdSave.Enabled = False
  Me.RichTextBox1.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : RichTextBox1_Change
' Purpose           : Enable the Save button if the text has changed
'*******************************************************************************
Private Sub RichTextBox1_Change()
  If Not Me.cmdSave.Enabled Then
    Me.cmdSave.Enabled = Me.chkEnableEdit.Value = vbChecked
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : RichTextBox1_KeyPress
' Purpose           : Force user-entered data to be set in the Personal Notes color
'*******************************************************************************
Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
  If Me.chkEnableEdit.Value = vbChecked Then
    If Me.RichTextBox1.SelColor <> PNotesColor Then
      Me.RichTextBox1.SelColor = PNotesColor
    End If
  End If
End Sub

'***************************************************************
' ShowDropdown(): force Combobox dropdown list to display
'EXAMPLE: ShowDropdown Combo1
'***************************************************************
Public Sub ShowDropdown(cboThis As ComboBox)
  Call SendMessageLong(cboThis.hwnd, CB_SHOWDROPDOWN, True, ByVal 0&)
End Sub

'***************************************************************
' Various routines to force comboboxes to drop down
'***************************************************************
Private Sub lblBook_Click()
  ShowDropdown Me.cboBook
End Sub

Private Sub lblChapter_Click()
  ShowDropdown Me.cboChapter
End Sub

Private Sub LblVerse_Click()
  ShowDropdown Me.cboVerse
End Sub

Private Sub cboBook_GotFocus()
  ShowDropdown Me.cboBook
End Sub

Private Sub cboChapter_GotFocus()
  ShowDropdown Me.cboChapter
End Sub

Private Sub cboVerse_GotFocus()
  ShowDropdown Me.cboVerse
End Sub

'*******************************************************************************
' Subroutine Name   : Timer1_Timer
' Purpose           : Prevent processing error if the form is unloaded from within the
'                   : LOAD event
'*******************************************************************************
Private Sub Timer1_Timer()
  Me.Timer1.Enabled = False
  Unload Me
End Sub
