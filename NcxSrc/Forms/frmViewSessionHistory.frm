VERSION 5.00
Begin VB.Form frmViewSessionHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verse Viewing History"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5475
   Icon            =   "frmViewSessionHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   5220
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdSortBbl 
      Caption         =   "Sort in &Bible Order"
      Height          =   375
      Left            =   3420
      TabIndex        =   7
      ToolTipText     =   "Sort the list and view in Bible order"
      Top             =   3300
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortAlpha 
      Caption         =   "Sort in &Alpha Order"
      Height          =   375
      Left            =   3420
      TabIndex        =   6
      ToolTipText     =   "Sort the list and view in alphabetical order"
      Top             =   2820
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortView 
      Caption         =   "Sort in &Viewed Order"
      Height          =   375
      Left            =   3420
      TabIndex        =   5
      ToolTipText     =   "Sort the list and view in viewed order"
      Top             =   2340
      Width           =   1815
   End
   Begin VB.CheckBox chkSaveSession 
      Caption         =   "Check1"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Save session history between application startups"
      Top             =   5640
      Width           =   195
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo changes"
      Height          =   375
      Left            =   3420
      TabIndex        =   11
      ToolTipText     =   "Undo any changes to the history list"
      Top             =   4980
      Width           =   1815
   End
   Begin VB.CommandButton cmdCopySel 
      Caption         =   "&Copy Selected"
      Height          =   375
      Left            =   3420
      TabIndex        =   4
      ToolTipText     =   "Copy select lines to the clipboard"
      Top             =   1740
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "&Select &All"
      Height          =   375
      Left            =   3420
      TabIndex        =   3
      ToolTipText     =   "Select all entries"
      Top             =   1260
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemoveSel 
      Caption         =   "&Remove Selected"
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      ToolTipText     =   "Remove selected lines from the list (cannot remove active verse)"
      Top             =   780
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelDupes 
      Caption         =   "&Delete Older Duplicates"
      Height          =   375
      Left            =   3420
      TabIndex        =   1
      ToolTipText     =   "Delete older entries that are duplicates of more recent entries"
      Top             =   300
      Width           =   1815
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go To Verse"
      Default         =   -1  'True
      Height          =   375
      Left            =   3420
      TabIndex        =   8
      ToolTipText     =   "Go to the selected book, chapter, and verse."
      Top             =   3900
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   3420
      TabIndex        =   12
      ToolTipText     =   "Close this dialog, accept any changes"
      Top             =   5460
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   180
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   300
      Width           =   3015
   End
   Begin VB.Line Line2 
      X1              =   3420
      X2              =   5220
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Line Line1 
      X1              =   3420
      X2              =   5220
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Label lblSaveSession 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save session view history"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Save session view history between application startups"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select multiple lines by holding the CNTRL key, or SHIFT for ranges."
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   3420
      TabIndex        =   14
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of up to the last 1000 verses viewed:"
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   60
      Width           =   2895
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuPopupDelDupes 
         Caption         =   "&Delete older duplicates"
      End
      Begin VB.Menu mnuPopupRemovelSel 
         Caption         =   "&Remove selected entries"
      End
      Begin VB.Menu mnuPopupSelectAll 
         Caption         =   "&Select all"
      End
      Begin VB.Menu mnuPopupCopySel 
         Caption         =   "&Copy selected entries"
      End
      Begin VB.Menu mnuPopupSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupSortView 
         Caption         =   "Sort &view order"
      End
      Begin VB.Menu mnuPopupSortAlpha 
         Caption         =   "Sort &alpha order"
      End
      Begin VB.Menu mnuPopupSortBbl 
         Caption         =   "Sort &Bible order"
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupGoto 
         Caption         =   "&Go to selected verse"
      End
      Begin VB.Menu mnuPopupSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupUndo 
         Caption         =   "&Undo changes"
      End
      Begin VB.Menu mnuPopupClose 
         Caption         =   "Cl&ose window"
      End
   End
End
Attribute VB_Name = "frmViewSessionHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyToolTips As clsToolTip   'MyToolTips can be any name useful to you
Private Const VVV As String = "Verse Viewing History"
Private lcol As Collection

'*******************************************************************************
' Subroutine Name   : chkSaveSession_Click
' Purpose           : Toggle save history option
'*******************************************************************************
Private Sub chkSaveSession_Click()
  SaveSetting App.Title, "Settings", "SaveHistory", CStr(Me.chkSaveSession.Value)
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopySel_Click
' Purpose           : Copy selected lines to the clipboard
'*******************************************************************************
Private Sub cmdCopySel_Click()
  Dim S As String
  Dim Idx As Long
  
  With Me.List1
    For Idx = 0 To .ListCount - 1
      If .Selected(Idx) Then
        S = S & .List(Idx) & vbCrLf
      End If
    Next Idx
  End With
  
  If Len(S) <> 0 Then
    Clipboard.Clear
    Clipboard.SetText S, vbCFText
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDelDupes_Click
' Purpose           : Remove duplicates
'*******************************************************************************
Private Sub cmdDelDupes_Click()
  Dim S As String
  Dim Idx As Long
  Dim ccc As Collection
  
  Set ccc = New Collection
  With Me.List1
    On Error Resume Next                    'skip errors created by duplicate indexes
    For Idx = .ListCount - 1 To 0 Step -1
      S = .List(Idx)                        'get an item
      ccc.Add S, S                          'add to list, skip dupes
    Next Idx
    
    .Clear                                  'clear display list
    For Idx = ccc.Count To 0 Step -1
      .AddItem ccc(Idx)                     'add an item
      ccc.Remove Idx                        'remove from collection to flush it
    Next Idx
  End With
'
' now flush internal list to temp collection, and remove duplicates
'
  With lcol
    Do While .Count
      S = .Item(.Count)
      ccc.Add S, S
      .Remove .Count
    Loop
    For Idx = 1 To ccc.Count                'now refresh internal list
      .Add ccc(1)
      ccc.Remove 1                          'flush temp list
    Next Idx
  End With
    
  Set ccc = Nothing
  Me.cmdUndo.Enabled = True
  Me.mnuPopupUndo.Enabled = True
  Call List1_Click
  Call ShowCount
End Sub

'*******************************************************************************
' Subroutine Name   : cmdRemoveSel_Click
' Purpose           : Remove selected items
'*******************************************************************************
Private Sub cmdRemoveSel_Click()
  Dim S As String, T As String
  Dim Idx As Long, I As Long
  
  LockWindowUpdate Me.List1.hwnd
  S = lcol(lcol.Count)                          'get current entry
  With Me.List1
    For Idx = .ListCount - 1 To 0 Step -1       'scan through display list
      If .Selected(Idx) Then                    'if selected to remove
        If .List(Idx) = S Then                  'matches current?
          S = vbNullString                      'yes, do not remove, but only once
        Else
          T = .List(Idx)                        'grab item to remove
          .RemoveItem Idx                       'remove it from the display list
          For I = lcol.Count - 1 To 1 Step -1   'scan below current entry
            If lcol(I) = T Then                 'if a match was found
              lcol.Remove I                     'also remove it from the internal list
              Exit For                          'done scanning internal for now
            End If
          Next I                                'scan next internal
        End If
      End If
    Next Idx                                    'scan next display item
  End With
  Call List1_Click                              'refresh display data
  Me.cmdUndo.Enabled = True
  Me.mnuPopupUndo.Enabled = True
  Call ShowCount
  LockWindowUpdate 0
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSelAll_Click
' Purpose           : Select all
'*******************************************************************************
Private Sub cmdSelAll_Click()
  Dim Idx As Long
  
  With Me.List1
    For Idx = 0 To .ListCount - 1
      .Selected(Idx) = True
    Next Idx
  End With
  Call List1_Click
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSortAlpha_Click
' Purpose           : Sort the list alphabetically
'*******************************************************************************
Private Sub cmdSortAlpha_Click()
  Dim Idx As Long
  
  LockWindowUpdate Me.List1.hwnd
  Me.List2.Clear                            'clear the sorted list
  With Me.List1
    Do While .ListCount
      Me.List2.AddItem .List(0)             'copy current list to alpha sorted list
      .RemoveItem 0                         'remove the item from the displayed list
    Loop
  End With
'
' copy the sorted list back over to the displayed list
'
  With Me.List2
    Do While .ListCount
      Me.List1.AddItem .List(0)
      .RemoveItem 0                         'remove the item from the sorted list
    Loop
  End With
  Me.cmdSortAlpha.Enabled = False
  Me.cmdSortBbl.Enabled = True
  Me.cmdSortView.Enabled = True
  Call List1_Click
  LockWindowUpdate 0
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSortBbl_Click
' Purpose           : Sort in Bible Order
'*******************************************************************************
Private Sub cmdSortBbl_Click()
  Dim Idx As Long, I As Long, J As Long, K As Long
  Dim S As String, Ary() As String
  
  LockWindowUpdate Me.List1.hwnd
  Me.List2.Clear                            'clear the sorted list
  With Me.List1
    Do While .ListCount
      Me.List2.AddItem .List(0)             'copy current list to alpha sorted list
      .RemoveItem 0                         'remove the item from the displayed list
    Loop
  End With
  
  With Me.List2
    For I = 27 To 1 Step -1                 'scan through book names backward
      Ary = Split(Books(I), ",")
      S = Ary(3) & " "                      'get a book with a space
      J = Len(S)                            'get its length
      For K = .ListCount - 1 To 0 Step -1   'scan through sorted list backward
        If Left$(.List(K), J) = S Then      'if a match found...
          Me.List1.AddItem .List(K), 0      'add to display list at beginning
          .RemoveItem K                     'remove item from the sorted list
        End If
      Next K
    Next I
    Do While .ListCount
      Me.List1.AddItem .List(0)
      .RemoveItem 0
    Loop
    Me.List1.ListIndex = -1
  End With
  
  Me.cmdSortBbl.Enabled = False
  Me.cmdSortAlpha.Enabled = True
  Me.cmdSortView.Enabled = True
  Call List1_Click
  LockWindowUpdate 0
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSortView_Click
' Purpose           : Display list in viewed order
'*******************************************************************************
Private Sub cmdSortView_Click()
  Dim Idx As Long
  
  LockWindowUpdate Me.List1.hwnd
  With Me.List1
    .Clear
    For Idx = 1 To lcol.Count
      .AddItem lcol(Idx)
    Next Idx
  End With
  Me.cmdSortView.Enabled = False
  Me.cmdSortAlpha.Enabled = True
  Me.cmdSortBbl.Enabled = True
  Call List1_Click
  LockWindowUpdate 0
End Sub

'*******************************************************************************
' Subroutine Name   : cmdUndo_Click
' Purpose           : Undo changes
'*******************************************************************************
Private Sub cmdUndo_Click()
  Dim S As String, Ary() As String
  Dim Idx As Long
  
  With Me.List1
    .Clear
    With colHistory
      For Idx = 1 To .Count
        S = .Item(Idx)
        Ary = Split(Books(CLng(Left$(S, 2))), ",")
        S = Ary(3) & " " & Mid$(S, 3, 2) & ":" & Right$(S, 2)
        Me.List1.AddItem S
      Next Idx
    End With
    .ListIndex = .ListCount - 1
    .Selected(.ListCount - 1) = True
  End With
  Call GetList
  Me.cmdSortView.Enabled = False
  Me.cmdSortAlpha.Enabled = True
  Me.cmdSortBbl.Enabled = True
  
  Me.cmdUndo.Enabled = False
  Me.mnuPopupUndo.Enabled = False
  Call ShowCount
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Load form
'*******************************************************************************
Private Sub Form_Load()
  Dim Idx As Long, B As Long, C As Long, V As Long, I As Long
  Dim S As String, Ary() As String, Cur As String
  
  Set lcol = New Collection         'define internal list
  Set MyToolTips = New clsToolTip   'declare object
  With MyToolTips
    .Create Me                      'create object
    .MaxTipWidth = 1440 * 6         'set to 4 inches
    .DelayTime(ttDelayShow) = 20 * 1000 'set to 20 seconds
    .SetFont , 12
    .AddTool Me.List1
    .ToolText(Me.List1) = vbNullString
  End With
  
  With frmGrkXlate
    Me.Top = .Top
    Me.Left = (.Width - Me.Width) / 2 + .Left
    Me.Height = .Height - Screen.TwipsPerPixelY * 4
  End With
  
  Me.List1.BackColor = cLight
  Ary = Split(Books(Bk), ",")
  Cur = Ary(3) & " " & Format(Chp, "00") & ":" & Format(Vrs, "00")
  
  With colHistory
    For Idx = 1 To .Count
      S = .Item(Idx)
      Ary = Split(Books(CLng(Left$(S, 2))), ",")
      S = Ary(3) & " " & Mid$(S, 3, 2) & ":" & Right$(S, 2)
      Me.List1.AddItem S
      lcol.Add S
    Next Idx
  End With
  
  With Me.List1
    .ListIndex = .ListCount - 1
    .Selected(.ListCount - 1) = True
  End With
'
' Get History option
'
  Me.chkSaveSession.Value = CInt(GetSetting(App.Title, "Settings", "SaveHistory", "1"))
'
' hide popup menu
'
  Me.mnuPopup.Visible = False
  Me.cmdUndo.Enabled = False
  Me.mnuPopupUndo.Enabled = False
'
' init sort buttons
'
  Me.cmdSortView.Enabled = False
  Me.cmdSortAlpha.Enabled = True
  Me.cmdSortBbl.Enabled = True
'
'Show form title and entry count
'
  Call ShowCount
End Sub

'*******************************************************************************
' Subroutine Name   : GetList
' Purpose           : Grab the master list to the internal list
'*******************************************************************************
Private Sub GetList()
  Dim Idx As Long
  
  With lcol
    Do While .Count                 'flush internal list
      .Remove 1
    Loop
  End With
  
  With Me.List1
    For Idx = 0 To .ListCount - 1   'rebuild master list
      lcol.Add .List(Idx)
    Next Idx
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : ShowCount
' Purpose           : Show form title and entry count
'*******************************************************************************
Private Sub ShowCount()
  Dim S As String
  
  S = VVV & " (" & CStr(Me.List1.ListCount) & " entr"
  If Me.List1.ListCount = 1 Then
    S = S & "y)"
  Else
    S = S & "ies)"
  End If
  Me.Caption = S
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Update the form background tilingx
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Enable Go key if something active
'*******************************************************************************
Private Sub Form_Resize()
  Dim I As Long
  If Me.Width <> 5565 Then Me.Width = 5565
  If Me.Height < 5400 Then Me.Height = 5400
  Me.chkSaveSession.Top = Me.ScaleHeight - Me.chkSaveSession.Height - 120
  Me.lblSaveSession.Top = Me.chkSaveSession.Top
  Me.List1.Height = Me.chkSaveSession.Top - Me.List1.Top - 120
  Me.cmdCancel.Top = Me.List1.Top + Me.List1.Height - Me.cmdCancel.Height
  Me.cmdUndo.Top = Me.cmdCancel.Top - Me.cmdUndo.Height - 120
  I = Me.cmdGo.Top + Me.cmdGo.Height
  Me.lblInfo.Top = (Me.cmdUndo.Top - I - Me.lblInfo.Height) \ 2 + I
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Close session
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdGo_Click
' Purpose           : Go to selected verse
'*******************************************************************************
Private Sub cmdGo_Click()
  Dim Cur As String, S As String, Ary() As String
  Dim Idx As Long, B As Long, C As Long, V As Long
  
  With Me.List1
    Cur = .List(.ListIndex)
  End With
  With colHistory
    For Idx = 1 To .Count
      S = .Item(Idx)
      B = CLng(Left$(S, 2))
      C = CLng(Mid$(S, 3, 2))
      V = CLng(Right$(S, 2))
      Ary = Split(Books(B), ",")
      If Cur = Ary(3) & " " & Format(C, "00") & ":" & Format(V, "00") Then
        Bk = B                                  'get book, chapter and verse
        Chp = C
        Vrs = V
        ChpCnt = CLng(Ary(4))                   'get the chapter count
        Call frmGrkXlate.GetVerseCount          'get the verse count
        Call frmGrkXlate.UpdateVerse            'display the verse
        Unload Me
        Exit For
      End If
    Next Idx
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Unload form
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Dim S As String, Ary() As String, T As String
  Dim Idx As Long, I As Long
  
  If Me.cmdUndo.Enabled Then          'if something has changed...
    With colHistory                   'clear listory list
      Do While .Count
        .Remove 1
      Loop
    End With
    
    With Me.List1
      For Idx = 0 To .ListCount - 1   'build new list
        S = .List(Idx)                'get an entry
        I = InStr(3, S, " ")          'find book data
        T = Left$(S, I - 1)           'save it
        S = Right$(S, 5)              'maintain chapter and verse
        For I = 1 To 27               'find book index
          Ary = Split(Books(I), ",")
          If Ary(3) = T Then Exit For
        Next I
        colHistory.Add Format(I, "00") & Left$(S, 2) & Right$(S, 2)
      Next Idx
    End With
  End If
  
  Set MyToolTips = Nothing          'destroy object
  Set lcol = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : lblSaveSession_Click
' Purpose           : Toggle save history option
'*******************************************************************************
Private Sub lblSaveSession_Click()
  If Me.chkSaveSession.Value = vbChecked Then
    Me.chkSaveSession.Value = vbUnchecked
  Else
    Me.chkSaveSession.Value = vbChecked
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : List1_Click
' Purpose           : Enable Go key if something active
'*******************************************************************************
Private Sub List1_Click()
  Dim Idx As Long
  Dim S As String
  
  Me.cmdGo.Enabled = Me.List1.SelCount = 1    'if only 1 item selected, allow GO button
  If Me.cmdGo.Enabled Then
    With Me.List1
      S = lcol(lcol.Count)                    'current verse
      For Idx = 0 To .ListCount - 1
        If .Selected(Idx) Then                'see if selected is also current verse
          If .List(Idx) = S Then
            Me.cmdGo.Enabled = False          'if so, disable
          End If
          Exit For                            'exit regardless when selected found
        End If
      Next Idx
    End With
  End If
  Me.mnuPopupGoto.Enabled = Me.cmdGo.Enabled  'reflect GO enablement ino the menu
  Me.cmdSelAll.Enabled = Me.List1.ListCount > 1
  Me.cmdRemoveSel.Enabled = Me.List1.SelCount <> 0 And Me.List1.ListCount > 1
  With Me.List1
    If .SelCount = 1 Then
      S = lcol(lcol.Count)
      For Idx = 0 To .ListCount - 1
        If .Selected(Idx) = True Then
          If .List(Idx) = S Then
            Me.cmdRemoveSel.Enabled = False
          End If
          Exit For
        End If
      Next Idx
    End If
  End With
  Me.cmdDelDupes.Enabled = CheckDupe()
  Me.cmdCopySel.Enabled = Me.List1.SelCount <> 0
  
  Me.mnuPopupSelectAll.Enabled = Me.cmdSelAll.Enabled
  Me.mnuPopupRemovelSel.Enabled = Me.cmdRemoveSel.Enabled
  Me.mnuPopupDelDupes.Enabled = Me.cmdDelDupes.Enabled
  Me.mnuPopupCopySel.Enabled = Me.cmdCopySel.Enabled
  Me.mnuPopupSortView.Enabled = Me.cmdSortView.Enabled
  Me.mnuPopupSortAlpha.Enabled = Me.cmdSortAlpha.Enabled
  Me.mnuPopupSortBbl.Enabled = Me.cmdSortBbl.Enabled
End Sub

'*******************************************************************************
' Subroutine Name   : List1_DblClick
' Purpose           : Go to selected verse
'*******************************************************************************
Private Sub List1_DblClick()
  Me.cmdGo.Value = True
End Sub

'*******************************************************************************
' Subroutine Name   : List1_MouseDown
' Purpose           : Show popup menu
'*******************************************************************************
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton And Shift = 0 Then
    PopupMenu mnuPopup, vbPopupMenuRightButton
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : List1_MouseMove
' Purpose           : update tooltip as required
'*******************************************************************************
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Cur As String, S As String, Ary() As String, T As String
  Dim Idx As Long, B As Long, C As Long, V As Long
  
  Cur = Trim$(GetStringFromMouseMove(Me.List1, X, Y))  'get line data
  If Len(Cur) <> 0 Then
    With colHistory
      For Idx = 1 To .Count
        S = .Item(Idx)                                'get BkChVs format
        B = CLng(Left$(S, 2))                         'grab book
        C = CLng(Mid$(S, 3, 2))                       'grab chapter
        V = CLng(Right$(S, 2))                        'grab verse
        Ary = Split(Books(B), ",")                    'get book
        If Cur = Ary(3) & " " & Format(C, "00") & ":" & Format(V, "00") Then
          S = Format(B, "00") & Format(C, "00") & Format(V, "00")
          Idx = FindExactMatch(frmGrkXlate.lstGrk, S) 'point to Greek text for verse
          If Idx <> -1 Then
            Select Case BblVersion
              Case 1
                S = "YLT"
              Case 2
                S = "RSV"
              Case UserPVer
                S = "MPV"
              Case 4
                S = "MKJV"
              Case 5
                S = "WEB"
              Case 6
                S = "ASV"
              Case 7
                S = "DBY"
              Case 8
                S = "WBS"
              Case Else
                S = "KJV"
            End Select
            Cur = Cur & " (" & S & ")  " & Mid$(Bible(Idx), 8)  'strip off leader
          End If
          Exit For
        End If
      Next Idx
    End With
  End If
  S = MyToolTips.ToolText(Me.List1)
  If S <> Cur Then MyToolTips.ToolText(Me.List1) = Cur
End Sub

Private Sub mnuPopupClose_Click()
  Me.cmdCancel.Value = True
End Sub

Private Sub mnuPopupCopySel_Click()
  Me.cmdCopySel.Value = True
End Sub

Private Sub mnuPopupDelDupes_Click()
  Me.cmdDelDupes.Value = True
End Sub

Private Sub mnuPopupGoto_Click()
  Me.cmdGo.Value = True
End Sub

Private Sub mnuPopupRemovelSel_Click()
  Me.cmdRemoveSel.Value = True
End Sub

Private Sub mnuPopupSelectAll_Click()
  Me.cmdSelAll.Value = True
End Sub

Private Sub mnuPopupSortAlpha_Click()
  Me.cmdSortAlpha.Value = True
End Sub

Private Sub mnuPopupSortBbl_Click()
  Me.cmdSortBbl.Value = True
End Sub

Private Sub mnuPopupSortView_Click()
  Me.cmdSortView.Value = True
End Sub

Private Sub mnuPopupUndo_Click()
  Me.cmdUndo.Value = True
End Sub

'*******************************************************************************
' Function Name     : CheckDupe
' Purpose           : Return true if list contains duplicates
'*******************************************************************************
Private Function CheckDupe() As Boolean
  Dim S As String
  Dim Idx As Long
  Dim ccc As Collection
  
  Set ccc = New Collection
  With Me.List1
    On Error Resume Next
    For Idx = .ListCount - 1 To 0 Step -1
      S = .List(Idx)
      ccc.Add S, S
      If Err.Number <> 0 Then Exit For
    Next Idx
    CheckDupe = ccc.Count <> .ListCount
  End With
  Set ccc = Nothing
End Function
