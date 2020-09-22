VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShowWords 
   Caption         =   "View Synonym List"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   Icon            =   "frmShowWords.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8025
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
      Left            =   7515
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Find next match in List"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdFindInText 
      Height          =   315
      Left            =   7140
      Picture         =   "frmShowWords.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Search for a word or phrase in the text below (Ctrl-S)"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picVBar 
      Height          =   5835
      Left            =   5220
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5775
      ScaleWidth      =   60
      TabIndex        =   14
      ToolTipText     =   "Drag to resize"
      Top             =   300
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   "Copy Data to the clipboard"
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rtbData 
      Height          =   5835
      Left            =   5400
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   10292
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmShowWords.frx":09CC
   End
   Begin VB.CheckBox chkViewFull 
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   6840
      Width           =   195
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   7590
      Width           =   8025
      _ExtentX        =   14155
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5820
      Top             =   7020
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5835
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   10292
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Strong #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Words"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7500
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   6240
      Width           =   975
   End
   Begin VB.CheckBox chkBreakup 
      Height          =   195
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Break up list to individual words"
      Top             =   6600
      Width           =   195
   End
   Begin VB.CommandButton cmdFind 
      Height          =   315
      Left            =   4860
      Picture         =   "frmShowWords.frx":0A4E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Find a word in the list (Ctrl-F)"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblWordRef 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double-click Words or Strong Numbers to access them"
      Height          =   195
      Left            =   5400
      TabIndex        =   17
      ToolTipText     =   "Double-click Words or Strong Numbers to access them"
      Top             =   60
      Width           =   3885
   End
   Begin VB.Label lblViewFull 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&View Data Full-Screen"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "View Data Full-Screen"
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblWidth"
      Height          =   195
      Left            =   5700
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblBreakup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Break up list to individual words (and sort)"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Break up list to individual words"
      Top             =   6600
      Width           =   2925
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double-Click a row  in the list to display its data"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   3300
   End
End
Attribute VB_Name = "frmShowWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Loading As Boolean    'true when loading form
Private DidLists As Boolean   'true if list2/3 have been built
Private SelItem As Long       'selected item index in listview
Private SaveW As Long, SaveH As Long, SaveT As Long, SaveL As Long
Private vBar As Boolean
Private LastSch As String
Private MyToolTips As clsToolTip

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

Private Sub Form_Load()
  Dim T As String, Ary() As String
  Dim Idx As Long, Lf As Long, Wd As Long, Tp As Long, Ht As Long
  
  Me.Enabled = False
  DoEvents
  Loading = True                        'indicate that we are loading the form
  ShowWordList = True                   'indicate that the list is being displayed
  Set Me.lblwidth.Font = Me.ListView1.Font  'set sizing fontsize to text field
  Me.chkBreakup.Value = CLng(GetSetting(App.Title, "Settings", "BreakWordList", "0"))
  frmGrkXlate.Favorz = vbNullString
  
  Me.picVBar.BorderStyle = 0
  Me.picVBar.Width = 120
  Me.picVBar.Visible = False
  Me.rtbData.Visible = False
  Me.lblWordRef.Visible = False
  Me.cmdCopy.Visible = False
'
' set some backgrounds
'
  Me.cmdFind.BackColor = cLight
  Me.ListView1.BackColor = cVLight
  Me.cmdFindInText.BackColor = cLight
  Me.cmdFindNext.BackColor = cLight
  Me.cmdFindInText.Visible = False
  Me.cmdFindNext.Visible = False
  Me.cmdFindNext.Enabled = False
  
  Loading = False
  Call chkBreakup_Click                     'load appropriate list
  If WordWidth = 0 Then                     'if width not set...
    Call GetScreenWorkArea(Lf, Wd, Tp, Ht)  'get screen work area sizing
    WordTop = Tp                            'top/left and full work area height
    WordLeft = Lf
    WordHeight = Ht
  End If
  
  Me.Top = WordTop                          'set window to it
  Me.Left = WordLeft
  Me.Height = WordHeight
'
' initially, display currently selected word
'
  If Me.chkBreakup.Value = vbUnchecked Then 'if a list
    Ary = Split(WordRef(Strong), vbTab)     'grab words
    Ary = Split(Ary(2), ",")                'split them apart
    T = Join(Ary, ", ")                     'rejoin with inserted spaces after each comma
  Else
    With frmGrkXlate.lstWords
      If .ListCount > 0 And .ListIndex <> -1 Then
        T = .List(.ListIndex)       'get current active synonym
      Else
        T = vbNullString
      End If
    End With
  End If
'
' now scan current list for a match
'
  If Len(T) <> 0 Then
    With Me.ListView1.ListItems
      For Idx = 1 To .Count
        If StrComp(T, .Item(Idx).SubItems(1), vbTextCompare) = 0 Then
          SelItem = Idx
          Me.Timer1.Enabled = True          'use a timer because we are not displayed yet
          Exit For
        End If
      Next Idx
    End With
  Else
    SelItem = 0
  End If
'
' set custom tootip for treeview
'
  Set MyToolTips = New clsToolTip
  With MyToolTips
    .Create Me               'create object
    .MaxTipWidth = 1440 * 2  'width max = 2 inches
    .DelayTime(ttDelayShow) = 20 * 1000 'set to 20 seconds
    .SetFont , 8
    .AddTool Me.ListView1
    .ToolText(Me.ListView1) = vbNullString
  End With
  Me.Show
  frmGrkXlate.mnuWinStrong.Enabled = True
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
    frmGrkXlate.mnuViewSyn.Enabled = True   'enable main menu option if we are minimized
    frmGrkXlate.Toolbar1.Buttons(7).Enabled = True
    frmGrkXlate.ZOrder 0
    frmGrkXlate.SetFocus
    Exit Sub
  End If
  Resizing = True
  
  If WordWidth <> 0 Then                    'if we have data to set
    Me.Width = WordWidth                    'set width and height
    Me.Height = WordHeight
    WordWidth = 0                           'reset to prevent redundancy
    WordHeight = 0
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
  
  frmGrkXlate.mnuViewSyn.Enabled = False      'dispable activator on main menu
  frmGrkXlate.Toolbar1.Buttons(7).Enabled = False
  If Me.Height < 4320 Then Me.Height = 4320
  If Me.chkViewFull.Value = vbUnchecked Then  'if not full data...
    If Me.Width < 4320 Then Me.Width = 4320
    Me.ListView1.Width = Me.ScaleWidth - Me.ListView1.Left * 2
    Me.rtbData.Visible = False
    Me.lblWordRef.Visible = False
    Me.cmdCopy.Visible = False
    Me.picVBar.Visible = False
  Else
    If Me.Width < 8000 Then Me.Width = 8000
    With Me.ListView1
      If vBar Then
        .Width = Me.picVBar.Left - .Left
        vBar = False
      Else
        Me.picVBar.Left = .Left + .Width
      End If
      Me.picVBar.Top = .Top
      Me.picVBar.Height = .Height
    End With
    With Me.rtbData
      .Top = Me.picVBar.Top
      .Left = Me.picVBar.Left + Me.picVBar.Width
      .Width = Me.ScaleWidth - .Left
      .BackColor = cLight
      Me.cmdCopy.Left = .Left
      .Visible = True
      Me.cmdCopy.Visible = True
      Me.cmdCopy.Enabled = False
      Me.picVBar.Visible = True
      Me.lblWordRef.Left = .Left
      Me.lblWordRef.Visible = True
    End With
  End If
'
' shift form data around...
'
  Me.ListView1.ColumnHeaders(2).Width = Me.ListView1.Width - Me.ListView1.ColumnHeaders(1).Width - 360
  Me.cmdClose.Left = Me.ScaleWidth - Me.ListView1.Left - Me.cmdClose.Width
  Me.cmdClose.Top = Me.ScaleHeight - Me.StatusBar1.Height - Me.cmdClose.Height - 60
  Me.chkBreakup.Top = Me.cmdClose.Top
  Me.lblBreakup.Top = Me.chkBreakup.Top
  Me.chkViewFull.Top = Me.chkBreakup.Top + Me.chkBreakup.Height + 30
  Me.lblViewFull.Top = Me.chkViewFull.Top
  With Me.ListView1
    .Height = Me.cmdClose.Top - .Top - 30
    Me.picVBar.Top = .Top
    Me.picVBar.Height = .Height
  End With
  Me.rtbData.Height = Me.ListView1.Height
  Me.cmdCopy.Top = Me.cmdClose.Top
  Me.cmdFind.Top = 0
  Me.cmdFind.Left = Me.ListView1.Left + Me.ListView1.Width - Me.cmdFind.Width
  Me.cmdFindNext.Left = Me.ScaleWidth - Me.cmdFindInText.Width
  Me.cmdFindInText.Left = Me.cmdFindNext.Left - Me.cmdFindInText.Width
  Resizing = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : When unloading, remove evidence
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  ShowWordList = False
  frmGrkXlate.mnuViewSyn.Enabled = True
  frmGrkXlate.Toolbar1.Buttons(7).Enabled = True
  SaveSetting App.Title, "Settings", "BreakWordList", CStr(Me.chkBreakup.Value)
  If Me.WindowState <> vbMinimized Then
    If Me.chkViewFull.Value = vbUnchecked Then
      WordTop = Me.Top
      WordLeft = Me.Left
      WordWidth = Me.Width
      WordHeight = Me.Height
    Else
      WordTop = SaveT
      WordLeft = SaveL
      WordWidth = SaveW
      WordHeight = SaveH
    End If
  End If
  Set MyToolTips = Nothing
  frmGrkXlate.mnuWinStrong.Enabled = False
  frmGrkXlate.CheckWin
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyDown
' Purpose           : Allow Ctrl-F for invoking FIND
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = vbCtrlMask Then    'CTRL?
    If KeyCode = 70 Then        'F?
      KeyCode = 0
      Me.cmdFind.Value = True
    ElseIf KeyCode = 83 Then    'S?
      KeyCode = 0
      Me.cmdFindInText.Value = True
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ListView1_MouseMove
' Purpose           : Displat text in tooltip if longer than displayed
'*******************************************************************************
Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String
  Dim lItm As MSComctlLib.ListItem
  Static LastTip As String
  
  Set lItm = Me.ListView1.HitTest(X, Y)   'anything hovered over?
  If lItm Is Nothing Then Exit Sub        'no, so ignore
  S = lItm.SubItems(1)                    'get text
  If S <> LastTip Then
    LastTip = S
    Me.lblwidth.Caption = S                 'test width
    If Me.lblwidth.Width < Me.ListView1.ColumnHeaders(2).Width - 240 Then S = vbNullString
    MyToolTips.ToolText(Me.ListView1) = S
  End If
End Sub

Private Sub picVBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 0 Then
     vBar = True
     Me.picVBar.BackColor = cdGray
  End If
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Lft As Long, Lf As Long
  
  If Not vBar Then Exit Sub
  With Me.picVBar
    Lft = .Left - (.Width \ 2 - CLng(X))
    Lf = Me.ScaleWidth - 3000
    If (Lft + .Width) > Lf Then Lft = Lf
    Lf = 4000
    If Lft < 4000 Then Lft = 4000
    If .Left <> Lft Then .Left = Lft
  End With
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.picVBar.BackColor = cLight
  Call Form_Resize
  Me.picVBar.Refresh
End Sub

Private Sub picVBar_Paint()
  PaintTilePicBackground Me.picVBar, frmGrkXlate.picTile(Background)   'repaint background
  frmGrkXlate.Vbarz Me.picVBar    'give it a 3D effect
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

'*******************************************************************************
' Subroutine Name   : Timer1_Timer
' Purpose           : Timer to allow screen refresh before displaying selection
'*******************************************************************************
Private Sub Timer1_Timer()
  Me.Timer1.Enabled = False   '1-shot deal after form load
  If SelItem <> 0 Then
    Me.ListView1.SetFocus                                 'set focus so that selection is displayed
    Me.ListView1.ListItems.Item(SelItem).EnsureVisible    'make sure it can be seen
    Me.ListView1.ListItems.Item(SelItem).Selected = True  'select the item
    Call ListView1_ItemClick(Me.ListView1.SelectedItem)   'process periphery services
    SelItem = 0                                           'disable index
  End If
  Me.chkViewFull.Value = CLng(GetSetting(App.Title, "Settings", "ViewSynFull", "0"))
  Me.cmdFindInText.Visible = Me.chkViewFull.Value = vbChecked
  Me.cmdFindNext.Visible = Me.cmdFindInText.Visible
  Me.lblWordRef.Visible = Me.cmdFindInText.Visible
End Sub

'*******************************************************************************
' Subroutine Name   : chkBreakup_Click
' Purpose           : Process user option on list
'*******************************************************************************
Private Sub chkBreakup_Click()
  Dim Idx As Integer, I As Long
  Dim Ary() As String, S As String, T As String
  Dim Itm As MSComctlLib.ListItem
  
  If Loading Then Exit Sub                    'ignore if we are loading the form
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  DoEvents
  Me.Caption = "View Synonym List"
  Me.StatusBar1.Panels(1).Text = vbNullString
  With Me.ListView1.ListItems
    .Clear
    For Idx = 1 To UBound(WordRef)
      S = WordRef(Idx)
      If Len(S) <> 0 Then
        Ary = Split(S, vbTab)
        S = Ary(2)
        T = Format(CLng(Ary(0)), "0000")
        If Len(S) <> 0 Then                   'ignore blank (Hebrew) lines
          Ary = Split(S, ",")
          If Me.chkBreakup.Value = vbChecked Then
            For I = 0 To UBound(Ary)
              Set Itm = .Add(, , T)
              Itm.SubItems(1) = Ary(I)
              If Not DidLists Then
                Me.List1.AddItem Ary(I)
                Me.List2.AddItem T
              End If
            Next I
          Else
            Set Itm = .Add(, , T)
            Itm.SubItems(1) = Join(Ary, ", ")
            If Not DidLists Then
              For I = 0 To UBound(Ary)
                Me.List1.AddItem Ary(I)
                Me.List2.AddItem T
              Next I
            End If
          End If
        End If
      End If
    Next Idx
    .Item(1).EnsureVisible
    DidLists = True
  End With
  With Me.ListView1
    .SortKey = 0
    .Sorted = True
    .SortKey = 1
  End With
  Me.Enabled = True
  Screen.MousePointer = vbDefault
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
  Dim I As Long, J As Long, Chk As Long
  
  T = Trim$(InputMsgBox(Me, "Enter an English word or Strong # to find:", "Find Word", , _
                        "Check the Vine database list if the word is not found in the synonym list."))
  If Len(T) = 0 Then Exit Sub
  Chk = CLng(GetSetting(App.Title, "Settings", "CheckVineAll", CStr(vbChecked)))
'
' if numeric, get the actual databased index and the search word
'
  If IsNumeric(T) Then
    I = CLng(T)
    If I >= 0 Or I <= UBound(WordRef) Then
      S = WordRef(I)
      If Len(S) <> 0 Then
        Ary = Split(S, vbTab)
        S = Ary(2)
      End If
      If Len(S) = 0 Then I = -1
    Else
      I = -1
    End If
    If I = -1 Then
      MessageBox Me, "Cannot find selected Strong Reference #: " & T, vbExclamation Or vbOKOnly, "Strong # Not Found"
      Exit Sub
    End If
  Else
    If Me.chkBreakup.Value = vbChecked Then
      With Me.ListView1.ListItems
        For I = 1 To .Count
          If StrComp(.Item(I).SubItems(1), T) = 0 Then Exit For
        Next I
        If I <= .Count Then
          Me.ListView1.SetFocus
          .Item(I).EnsureVisible
          .Item(I).Selected = True
          Call ListView1_ItemClick(Me.ListView1.SelectedItem)
          Exit Sub
        End If
        I = -1
      End With
    Else
      I = FindExactMatch(Me.List1, T)
      If I = -1 Then I = FindMatch(Me.List1, T)
    End If
    If I <> -1 Then I = CLng(Me.List2.List(I))
  End If
'
' if not found
'
  If I = -1 Then
    If Chk = vbChecked Then                                   'check Vine database
      With frmGrkXlate
        I = FindExactMatch(.lstVine, LCase$(T))               'find word in Vine list
        If I <> -1 Then                                       'found it?
          I = CLng(.lstVineNum.List(I)) - 1                   'yes get main list index
          .DisplayVine I, UCase$(T)                           'put to main form
          .cmdVine.Enabled = True                             'enable some standard stuff
          .cmdAnalysis.Enabled = True
          .cmdCopyDef.Enabled = True
          Me.rtbData.Text = vbNullString                      'adjust scrolling
          Me.rtbData.TextRTF = frmGrkXlate.rtbNotes.TextRTF   'grab data
          Me.rtbData.BackColor = frmGrkXlate.rtbNotes.BackColor 'reflect background color
          Me.cmdCopy.Enabled = True                             'enable the Copy button
          Exit Sub
        End If
      End With
    End If
    MessageBox Me, "Cannot find selected word: " & T, vbExclamation Or vbOKOnly, "Strong # Not Found"
    Exit Sub
  Else
    T = Format(I, "0000")
    With Me.ListView1.ListItems
      For I = 1 To .Count
        If .Item(I).Text = T Then
          With .Item(I)
            Me.ListView1.SetFocus
            .EnsureVisible
            .Selected = True
            Call ListView1_ItemClick(Me.ListView1.SelectedItem)
            Exit For
          End With
        End If
      Next I
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ListView1_ItemClick
' Purpose           : When list is clicked, show the reference index
'*******************************************************************************
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Dim S As String

  Me.Caption = "View Synonym List # " & Item.Text
  Me.StatusBar1.Panels(1).Text = Item.SubItems(1)
  If Me.chkViewFull.Value = vbChecked Then
    Call ListView1_DblClick
    Me.rtbData.Text = vbNullString
    Me.rtbData.TextRTF = frmGrkXlate.rtbNotes.TextRTF
    Me.rtbData.BackColor = frmGrkXlate.rtbNotes.BackColor
    Me.cmdCopy.Enabled = True
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ListView1_DblClick
' Purpose           : Show Strong reference data for selected item
'*******************************************************************************
Private Sub ListView1_DblClick()
  Dim I As Long
  Dim S As String, Ary() As String, Stng As String
  
  Stng = CStr(CLng(Me.ListView1.SelectedItem.Text))
  For I = 1 To UBound(DefRef) - 1
    S = DefRef(I)
    If Len(S) <> 0 Then
      Ary = Split(DefRef(I), vbTab)
      If Stng = Ary(5) Then
        With frmGrkXlate
         .ShowDefRef CStr(I), False
          .cmdVine.Enabled = True
          .cmdAnalysis.Enabled = True
          .cmdCopyDef.Enabled = True
          Exit For
        End With
      End If
    End If
  Next I
End Sub

'*******************************************************************************
' Subroutine Name   : ListView1_ColumnClick
' Purpose           : Sort columns as needed
'*******************************************************************************
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  With Me.ListView1
    If ColumnHeader.Index = 2 Then
      If .SortKey = 0 Then
        .SortOrder = lvwAscending
        .SortKey = 1
        .SortOrder = lvwAscending
      Else
        If .SortOrder = lvwAscending Then
          .SortOrder = lvwDescending
        Else
          .SortOrder = lvwAscending
        End If
        .SortKey = 0
        .SortKey = 1
      End If
    Else
      If .SortKey = 1 Then
        .SortKey = 0
        .SortOrder = lvwAscending
      Else
        If .SortOrder = lvwAscending Then
          .SortOrder = lvwDescending
        Else
          .SortOrder = lvwAscending
        End If
      End If
    End If
    .ListItems.Item(1).EnsureVisible
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : StatusBar1_DblClick
' Purpose           : Save the Panel data to the clipboard
'*******************************************************************************
Private Sub StatusBar1_DblClick()
  Clipboard.Clear
  Clipboard.SetText Me.StatusBar1.Panels(1).Text
End Sub

'*******************************************************************************
' Subroutine Name   : chkViewFull_Click
' Purpose           : User wnat to enable/disable full data view
'*******************************************************************************
Private Sub chkViewFull_Click()
  Call SaveSetting(App.Title, "Settings", "ViewSynFull", CStr(Me.chkViewFull.Value))
  Me.cmdFindInText.Visible = Me.chkViewFull.Value = vbChecked
  Me.cmdFindNext.Visible = Me.cmdFindInText.Visible
  Me.lblWordRef.Visible = Me.cmdFindInText.Visible
  Call Form_Resize
  If Me.chkViewFull.Value = vbChecked Then              'if starting to view full data
    Me.lblInfo.Caption = "Select a row  in the list to display its data"
    Call ListView1_DblClick                             'update master text
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
  Me.ListView1.SetFocus
  Me.cmdCopy.Enabled = False
End Sub

