VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmShowVineList 
   Caption         =   "View Vine Word List"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   Icon            =   "frmShowVineList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8835
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
      Left            =   8415
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Find next match in List"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdFindInText 
      Height          =   315
      Left            =   8040
      Picture         =   "frmShowVineList.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Search for a word or phrase in the text below (Ctrl-S)"
      Top             =   0
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2820
      Top             =   6360
   End
   Begin VB.PictureBox picVBar 
      Height          =   5835
      Left            =   3540
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5775
      ScaleWidth      =   60
      TabIndex        =   12
      ToolTipText     =   "Drag to resize"
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CheckBox chkViewFull 
      Caption         =   "Check1"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   6960
      Width           =   195
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Copy Data to the clipboard"
      Top             =   6180
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdFind 
      Height          =   315
      Left            =   2340
      Picture         =   "frmShowVineList.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Find a word in the list (Ctrl-F)"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox chkBreakup 
      Height          =   195
      Left            =   180
      TabIndex        =   6
      ToolTipText     =   "Break up list to individual words"
      Top             =   6720
      Width           =   195
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   6240
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   3255
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   7275
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   159
            MinWidth        =   26
            Object.Tag             =   ""
            Object.ToolTipText     =   "Double-Click to save this text to the clipboard"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbData 
      Height          =   5835
      Left            =   3780
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   10292
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmShowVineList.frx":316E
   End
   Begin VB.Label lblWordRef 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double-click Words, Strong #'s, or underlined verses to access them"
      Height          =   195
      Left            =   3840
      TabIndex        =   15
      ToolTipText     =   "Double-click Words, Strong #'s, or underlined verses to access them"
      Top             =   60
      Width           =   4845
   End
   Begin VB.Label lblViewFull 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&View Data Full-Screen"
      Height          =   195
      Left            =   420
      TabIndex        =   7
      ToolTipText     =   "View Data Full-Screen"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double-Click a word in the list to display its data"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   3345
   End
   Begin VB.Label lblBreakup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Break up list to individual words (and sort)"
      Height          =   195
      Left            =   420
      TabIndex        =   5
      ToolTipText     =   "Break up list to individual words"
      Top             =   6780
      Width           =   2925
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblWidth"
      Height          =   195
      Left            =   3480
      TabIndex        =   9
      Top             =   6780
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmShowVineList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Loading As Boolean    'true when loading form
Private SaveW As Long, SaveH As Long, SaveT As Long, SaveL As Long
Private vBar As Boolean       'true if the vertical bar is being moved
Private LastSch As String

'*******************************************************************************
' Subroutine Name   : cmdFindInText_Click
' Purpose           : Find text in the Notes panel
'*******************************************************************************
Private Sub cmdFindInText_Click()
  Dim Sch As String, Text As String
  Dim Idx As Long
  
  If Me.rtbData.SelLength <> 0 Then
    Text = Me.rtbData.SelText
  Else
    Text = LastSch
  End If
  Sch = InputMsgBox(Me, "Enter word or phrase to find:", "Search For Text", Text)
  If Len(Sch) = 0 Then Exit Sub
  LastSch = Sch
  With Me.rtbData
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
  With Me.rtbData
    Text = .Text
    Idx = .SelStart + .SelLength + 1
    Idx = InStr(Idx, Text, LastSch)
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not Me.List1.Enabled Then Me.List1.Enabled = True
End Sub

Private Sub Form_Load()
  Dim T As String
  Dim Lf As Long, Wd As Long, Tp As Long, Ht As Long, I As Long
  
  Loading = True                        'indicate that we are loading the form
  ShowVineList = True                   'indicate that the list is being displayed
  Set Me.lblWidth.Font = Me.List1.Font  'set sizing fontsize to text field
  Me.chkBreakup.Value = CLng(GetSetting(App.Title, "Settings", "BreakVineList", "0"))
  
  Me.cmdFind.BackColor = cLight
  Me.List1.BackColor = cVLight
  Me.cmdFindInText.BackColor = cLight
  Me.cmdFindNext.BackColor = cLight
  Me.cmdFindInText.Visible = False
  Me.cmdFindNext.Visible = False
  Me.cmdFindNext.Enabled = False
'
' initialize some start-up states
'
  Me.picVBar.BorderStyle = 0
  Me.picVBar.Width = 120
  Me.picVBar.Visible = False
  Me.rtbData.Visible = False
  Me.lblWordRef.Visible = False
  Me.cmdCopy.Visible = False
  
  Loading = False                       'allow updates
  Call chkBreakup_Click                 'load appropriate list
'
' set initial sizing
'
  If WordWidth = 0 Then
    Call GetScreenWorkArea(Lf, Wd, Tp, Ht)
    WordTop = Tp
    WordLeft = Lf
    WordHeight = Ht
  End If
  Me.Top = WordTop
  Me.Left = WordLeft
  Me.Height = WordHeight
'
' see if we can set our index to a match for the currently selected word
'
  With frmGrkXlate
    I = .lstWords.ListIndex
    If .lstWords.ListCount > 0 And I <> -1 Then
      T = .lstWords.List(.lstWords.ListIndex)       'get current active synonym
      I = FindExactMatch(.lstVine, T)               'can we find it in our list?
    End If
    If I <> -1 Then
      If Me.chkBreakup.Value = vbUnchecked Then   'yes. List not broken up?
        I = CLng(.lstVineNum.List(I)) - 1         'if not, get reference index
      End If
      Me.List1.ListIndex = I                      'set selection position
    End If
  End With
  
  Me.Timer1.Enabled = True
  Screen.MousePointer = vbDefault                 'show that we are no longer busy
  Me.Show
  DoEvents
  frmGrkXlate.mnuWinVine.Enabled = True
  frmGrkXlate.CheckWin
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Show a nice background
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Size form as needed
'*******************************************************************************
Private Sub Form_Resize()
  Dim Tp As Long, Lf As Long, Wt As Long, Ht As Long
  Static Resizing As Boolean
  
  If Resizing Then Exit Sub
  If Me.WindowState = vbMinimized Then
    frmGrkXlate.mnuViewVine.Enabled = True
    frmGrkXlate.Toolbar1.Buttons(6).Enabled = True
    frmGrkXlate.ZOrder 0
    frmGrkXlate.SetFocus
    Exit Sub
  End If
  Resizing = True
  
  If VineWidth <> 0 Then                      'if we have data to set
    Me.Width = VineWidth
    Me.Height = VineHeight
    VineWidth = 0
    VineHeight = 0
  End If
'
' if we are viewing full-screen data
'
  If Me.chkViewFull.Value = vbUnchecked Then  'not viewing...
    If Me.rtbData.Visible Then                'but data is visible...
      Me.Left = SaveL                         'then reset sizing...
      Me.Top = SaveT
      Me.Width = SaveW
      Me.Height = SaveH
    Else
      SaveL = Me.Left                         'else save current sizing
      SaveT = Me.Top
      SaveW = Me.Width
      SaveH = Me.Height
    End If
  Else
    If Not Me.rtbData.Visible Then            'want full, but not yet displayed?
      GetScreenWorkArea Lf, Wt, Tp, Ht        'yes, get full work area
      Me.Left = Lf                            'stuff sizing
      Me.Top = Tp
      Me.Width = Wt
      Me.Height = Ht
    End If
  End If
  
  frmGrkXlate.mnuViewVine.Enabled = False
  frmGrkXlate.Toolbar1.Buttons(6).Enabled = False
  If Me.Height < 4320 Then Me.Height = 4320
  If Me.chkViewFull.Value = vbUnchecked Then  'if not full data...
    If Me.Width < 4320 Then Me.Width = 4320
    Me.List1.Width = Me.ScaleWidth - Me.List1.Left * 2
    Me.rtbData.Visible = False
    Me.lblWordRef.Visible = False
    Me.cmdCopy.Visible = False
    Me.picVBar.Visible = False
  Else
    If Me.Width < 8000 Then Me.Width = 8000
    With Me.List1
      If vBar Then
        .Width = Me.picVBar.Left - .Left
        vBar = False
      Else
        Me.picVBar.Left = .Left + .Width
      End If
    End With
    With Me.rtbData
      .Top = Me.picVBar.Top
      .Left = Me.picVBar.Left + Me.picVBar.Width
      .Width = Me.ScaleWidth - .Left
      .BackColor = clBlue
      Me.cmdCopy.Left = .Left
      .Visible = True
      Me.cmdCopy.Visible = True
      Me.cmdCopy.Enabled = False
      Me.picVBar.Visible = True
      Me.lblWordRef.Left = .Left
      Me.lblWordRef.Visible = True
    End With
  End If
  
  Me.cmdClose.Left = Me.ScaleWidth - Me.List1.Left - Me.cmdClose.Width
  Me.cmdClose.Top = Me.ScaleHeight - Me.StatusBar1.Height - Me.cmdClose.Height - 60
  Me.chkBreakup.Top = Me.cmdClose.Top
  Me.lblBreakup.Top = Me.chkBreakup.Top
  Me.chkViewFull.Top = Me.chkBreakup.Top + Me.chkBreakup.Height + 30
  Me.lblViewFull.Top = Me.chkViewFull.Top
  With Me.List1
    .Height = Me.cmdClose.Top - .Top - 30
    Me.picVBar.Top = .Top
    Me.picVBar.Height = .Height
  End With
  Me.rtbData.Height = Me.List1.Height
  Me.cmdCopy.Top = Me.cmdClose.Top
  Me.cmdFind.Top = 0
  Me.cmdFind.Left = Me.List1.Left + Me.List1.Width - Me.cmdFind.Width
  Me.cmdFindNext.Left = Me.ScaleWidth - Me.cmdFindInText.Width
  Me.cmdFindInText.Left = Me.cmdFindNext.Left - Me.cmdFindInText.Width
  Resizing = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : When unloading, remove evidence
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  ShowVineList = False
  frmGrkXlate.mnuViewVine.Enabled = True
  frmGrkXlate.Toolbar1.Buttons(6).Enabled = True
  SaveSetting App.Title, "Settings", "BreakVineList", CStr(Me.chkBreakup.Value)
  If Me.WindowState <> vbMinimized Then
    If Me.chkViewFull.Value = vbUnchecked Then
      VineTop = Me.Top
      VineLeft = Me.Left
      VineWidth = Me.Width
      VineHeight = Me.Height
    Else
      VineTop = SaveT
      VineLeft = SaveL
      VineWidth = SaveW
      VineHeight = SaveH
    End If
  End If
  frmGrkXlate.mnuWinVine.Enabled = False
  frmGrkXlate.CheckWin
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyDown
' Purpose           : Allow Ctrl-F for invoking FIND
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = vbCtrlMask Then    'CTRL?
    If KeyCode = 17 Then
      Me.List1.Enabled = False
      KeyCode = 0
    End If
    If KeyCode = 70 Then        'F?
      Me.cmdFind.Value = True
      KeyCode = 0
      Me.List1.Enabled = True
    ElseIf KeyCode = 83 Then    'S?
      KeyCode = 0
      Me.cmdFindInText.Value = True
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : chkViewFull_Click
' Purpose           : User wnat to enable/disable full data view
'*******************************************************************************
Private Sub chkViewFull_Click()
  Call SaveSetting(App.Title, "Settings", "ViewVineFull", CStr(Me.chkViewFull.Value))
  Me.cmdFindInText.Visible = Me.chkViewFull.Value = vbChecked
  Me.cmdFindNext.Visible = Me.cmdFindInText.Visible
  Me.lblWordRef.Visible = Me.cmdFindInText.Visible
  Call Form_Resize
  If Me.chkViewFull.Value = vbChecked Then              'if starting to view full data
    Me.lblInfo.Caption = "Select a row  in the list to display its data"
    If Me.List1.ListIndex = -1 Then
      Me.List1.ListIndex = 0
      Exit Sub
    End If
    Call List1_DblClick                                 'update master text
    Me.rtbData.Text = vbNullString                      'ensure unscrolled
    Me.rtbData.TextRTF = frmGrkXlate.rtbNotes.TextRTF   'stuff new data
    Me.rtbData.BackColor = frmGrkXlate.rtbNotes.BackColor
    Me.cmdCopy.Enabled = True                           'allow copying
  Else
    Me.lblInfo.Caption = "Double-Click a row  in the list to display its data"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lblViewFull_Click
' Purpose           : Support the checkbox from the label
'*******************************************************************************
Private Sub lblViewFull_Click()
  With Me.chkViewFull
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = vbChecked
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopy_Click
' Purpose           : Copy text data to the clipboard
'*******************************************************************************
Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetText Me.rtbData.Text
  Clipboard.SetText Me.rtbData.TextRTF, vbCFRTF
  Me.List1.SetFocus
  Me.cmdCopy.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : chkBreakup_Click
' Purpose           : Process user option on list
'*******************************************************************************
Private Sub chkBreakup_Click()
  Dim Idx As Integer
  Dim Ary() As String
  
  If Loading Then Exit Sub                    'ignore if we are loading the form
  Me.Caption = "View Vine Word List"
  Me.StatusBar1.Panels(1).Text = vbNullString
  Me.List1.Clear                              'ensure the display list is cleared
  If Me.chkBreakup.Value = vbChecked Then     'if breaking up list, load internal list
    With frmGrkXlate.lstVine
      For Idx = 0 To .ListCount - 1
        Me.List1.AddItem .List(Idx)
      Next Idx
    End With
  Else
    For Idx = 1 To UBound(Vine) - 1           'else build master list
      Ary = Split(Vine(Idx), vbTab)
      Me.List1.AddItem Ary(1)
    Next Idx
  End If
  
  Me.List1.ListIndex = -1
End Sub

'*******************************************************************************
' Subroutine Name   : lblBreakup_Click
' Purpose           : Allow label to alter checkbox
'*******************************************************************************
Private Sub lblBreakup_Click()
  With Me.chkBreakup
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = vbChecked
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClose_Click
' Purpose           : Close down list form and erase traces
'*******************************************************************************
Private Sub cmdClose_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdFind_Click
' Purpose           : Find a word or words in the vine index
'*******************************************************************************
Private Sub cmdFind_Click()
  Dim S As String, T As String, Ary() As String, TT As String
  Dim I As Long, J As Long, Chk As Long, Idx As Long
  
  T = Trim$(InputMsgBox(Me, "Enter an English word or Reference # to find:", _
            "Find Word", , _
            "Check the full Vine database text if the word is not found in the word list.", True))
  If Len(T) = 0 Then Exit Sub
  Chk = CLng(GetSetting(App.Title, "Settings", "CheckVineAll", CStr(vbChecked)))
'
' if numeric, get the actual databased index and the search word
'
  With frmGrkXlate
    If IsNumeric(T) Then
      I = FindExactMatch(.lstVineNum, T)
      If I <> -1 Then T = .lstVine.List(I)    'get search word
    Else
      I = FindMatch(.lstVine, LCase$(T)) 'find word
    End If
'
' if not found (I=-1) and the user wants to check the full database...
'
    If I = -1 Then
      If Chk = vbChecked Then
        Screen.MousePointer = vbHourglass       'show that we are busy
        Me.Enabled = False
        DoEvents
        
        TT = " " & LCase$(T) & " "              'test string
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
            I = FindExactMatch(.lstVineNum, CStr(Idx))  'found it, get the ref index
            If I <> -1 Then
              Exit For            'exit scan
            End If
          End If
        Next Idx
        Screen.MousePointer = vbDefault
        Me.Enabled = True
      End If
    End If
    
    If I = -1 Then
      MessageBox Me, "The word """ & T & """ was not found.", vbOKOnly Or vbExclamation, "Word not found"
      Exit Sub
    End If
'
' found the data, so display it
'
    If Me.chkBreakup.Value = vbUnchecked Then
      I = CLng(.lstVineNum.List(I)) - 1
    End If
    Me.List1.ListIndex = I
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lblRef_DblClick
' Purpose           : Save text in reference box to the clipboard
'*******************************************************************************
Private Sub lblRef_DblClick()
  Clipboard.Clear
  Clipboard.SetText Me.StatusBar1.Panels(1).Text
End Sub

'*******************************************************************************
' Subroutine Name   : List1_Click
' Purpose           : When list is clicked, show the reference index
'*******************************************************************************
Private Sub List1_Click()
  Dim S As String
  
  If Me.List1.ListIndex = -1 Then                       'if nothing selected
    S = vbNullString
  Else
    If Me.chkBreakup.Value = vbChecked Then             'something, so show ref #
      S = "# " & frmGrkXlate.lstVineNum.List(Me.List1.ListIndex)
    Else
      S = "# " & CStr(Me.List1.ListIndex + 1)
    End If
    Me.StatusBar1.Panels(1).Text = Me.List1.List(Me.List1.ListIndex)
  End If
  Me.Caption = "View Vine Word List" & S
'
' if we are set to FULL VIEW, copy the Notes panel data to our own RTB
'
  If Me.chkViewFull.Value = vbChecked Then
    Call List1_DblClick                                 'treat a click like a double
    Me.rtbData.Text = vbNullString                      'adjust scrolling
    Me.rtbData.TextRTF = frmGrkXlate.rtbNotes.TextRTF   'grab data
    Me.rtbData.BackColor = frmGrkXlate.rtbNotes.BackColor 'reflect background color
    Me.cmdCopy.Enabled = True                             'enable the Copy button
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : List1_DblClick
' Purpose           : Display data in the Notes panel if a line is double-clicked
'*******************************************************************************
Private Sub List1_DblClick()
  Dim I As Long
  
  I = Me.List1.ListIndex                                'grab the selection
  With frmGrkXlate
    If Me.chkBreakup.Value = vbUnchecked Then
      I = FindExactMatch(.lstVineNum, CStr(I + 1))      'ensure we have the right ref#
    End If
    .DisplayVine I, .lstVine.List(I)                    'display the Vine data for it
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : List1_MouseMove
' Purpose           : If the selection text is too wide for the lsitbox, display it in
'                   : a tooltip
'*******************************************************************************
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim S As String
  Static LastTip As String
  
  S = GetStringFromMouseMove(Me.List1, x, y)
  If S <> LastTip Then
    LastTip = S
    Me.lblWidth.Caption = S
    If Me.lblWidth.Width < Me.List1.Width - 240 Then S = vbNullString
    If Me.List1.ToolTipText <> S Then Me.List1.ToolTipText = S
    Me.List1.ToolTipText = S
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : picVBar_MouseDown
' Purpose           : User is initiation a frame resize event
'*******************************************************************************
Private Sub picVBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Shift = 0 Then
     vBar = True                          'indicate the bar is being moved
     Me.picVBar.BackColor = cdGray        'reflect this visually with a color change
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : picVBar_MouseMove
' Purpose           : The User is dragging the bar
'*******************************************************************************
Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim Lft As Long, Lf As Long
  
  If Not vBar Then Exit Sub                 'ignore if we are not dragging
  With Me.picVBar
    Lft = .Left - (.Width \ 2 - CLng(x))    'update position
    Lf = Me.ScaleWidth - 3000               'get rightward limit
    If (Lft + .Width) > Lf Then Lft = Lf    'do not exceed it
    Lf = 4000                               'likewise leftward
    If Lft < 4000 Then Lft = 4000
    If .Left <> Lft Then .Left = Lft        'if a change, update it
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : picVBar_MouseUp
' Purpose           : The user is done moving the bar, so update the screen
'*******************************************************************************
Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.picVBar.BackColor = cLight
  Call Form_Resize
  Me.picVBar.Refresh
End Sub

'*******************************************************************************
' Subroutine Name   : picVBar_Paint
' Purpose           : Update the theme for the bar
'*******************************************************************************
Private Sub picVBar_Paint()
  If vBar Then Exit Sub                     'do not bother if resizing
  PaintTilePicBackground Me.picVBar, frmGrkXlate.picTile(Background)   'repaint background
  frmGrkXlate.Vbarz Me.picVBar
End Sub

'*******************************************************************************
' Subroutine Name   : rtbData_DblClick
' Purpose           : If the user double-clicks some data, see if we can find a
'                   : reference to it in the database
'*******************************************************************************
Private Sub rtbData_DblClick()
  Dim S As String, T As String, Ary() As String
  Dim I As Long, J As Long, K As Long
  
  If CheckUndl(Me.rtbData) Then Exit Sub       'user selected underlined text
  Screen.MousePointer = vbHourglass             'show busy
  Me.Enabled = False
  DoEvents
  
  With Me.rtbData
    T = .Text                                   'grab the data
    I = .SelStart + 1                           'set the cursor point
    J = InStrRev(T, " ", I)                     'find a leading space
    If J = 0 Then J = 1                         'none
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
      On Error Resume Next
      I = CLng(T)                               'see if it is a fluke
      If Err.Number <> 0 Then Exit Sub
      On Error GoTo 0
      
      For I = 1 To UBound(DefRef) - 1
        S = DefRef(I)
        If Len(S) <> 0 Then
          Ary = Split(DefRef(I), vbTab)
          If T = Ary(5) Then                    'found Strong #?
            frmGrkXlate.ShowDefRef CStr(I), False 'yes, so show the data
            Me.rtbData.Text = vbNullString
            Me.rtbData.TextRTF = frmGrkXlate.rtbNotes.TextRTF
            Me.rtbData.BackColor = frmGrkXlate.rtbNotes.BackColor
            Exit For                            'all done
          End If
        End If
      Next I
    ElseIf Left$(T, 1) <> "#" Then              'not a Vine reference #
      I = InStr(1, T, "_")                      'strip underscore connectors
      Do While I <> 0
        Mid$(T, I, 1) = " "
        I = InStr(1, T, "_")
      Loop
      
      I = FindExactMatch(frmGrkXlate.lstVine, T)  'see if the words matches
      If I <> -1 Then
        frmGrkXlate.DisplayVine I, T            'found data, so display it
        Me.rtbData.Text = vbNullString
        Me.rtbData.TextRTF = frmGrkXlate.rtbNotes.TextRTF
        Me.rtbData.BackColor = frmGrkXlate.rtbNotes.BackColor
      End If
    End If
  End With
  
  DoEvents
  Me.Enabled = True
  Screen.MousePointer = vbDefault
End Sub

Private Sub Timer1_Timer()
  Me.Timer1.Enabled = False   '1-shot deal after form load
  Me.chkViewFull.Value = CLng(GetSetting(App.Title, "Settings", "ViewVineFull", "0"))
  Me.cmdFindInText.Visible = Me.chkViewFull.Value = vbChecked
  Me.cmdFindNext.Visible = Me.cmdFindInText.Visible
  Me.lblWordRef.Visible = Me.cmdFindInText.Visible
End Sub
