VERSION 5.00
Begin VB.Form frmFavs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Favorites"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5490
   Icon            =   "frmFavs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear All Entries"
      Height          =   315
      Left            =   3300
      TabIndex        =   4
      ToolTipText     =   "Erase the entire list"
      Top             =   1680
      Width           =   1850
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete Entry"
      Height          =   315
      Left            =   3300
      TabIndex        =   3
      ToolTipText     =   "Delete the selected entry (Del)"
      Top             =   1260
      Width           =   1850
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort &Alpha Order"
      Height          =   315
      Left            =   3300
      TabIndex        =   1
      ToolTipText     =   "Sort the list alphabetically"
      Top             =   300
      Width           =   1850
   End
   Begin VB.CommandButton cmdSortBbl 
      Caption         =   "&Sort &Bible Order"
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      ToolTipText     =   "Sort the list in Bible order"
      Top             =   720
      Width           =   1850
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
      Height          =   5340
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   300
      Width           =   3015
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4260
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      ToolTipText     =   "Cancel changes"
      Top             =   5280
      Width           =   1850
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Accept"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      ToolTipText     =   "Accept changes"
      Top             =   4800
      Width           =   1850
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   5160
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double-click verse to go to it."
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   2070
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuPopupSortAlpha 
         Caption         =   "Sort &alpha order"
      End
      Begin VB.Menu mnuPopupSortBbl 
         Caption         =   "Sort &Bible order"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupDel 
         Caption         =   "&Delete selected entry"
      End
      Begin VB.Menu mnuPopupClear 
         Caption         =   "&Clear all entries"
      End
   End
End
Attribute VB_Name = "frmFavs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyToolTips As clsToolTip   'MyToolTips can be any name useful to you
Private Favs As Long
Private SortBbl As Boolean

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Cancel changes
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClear_Click
' Purpose           : Clewar all entries from the list
'*******************************************************************************
Private Sub cmdClear_Click()
  If MessageBox(Me, "Verify clearing entire favorites list.", vbYesNo Or vbQuestion, "Verify Delete All") = vbNo Then Exit Sub
  Me.List1.Clear
  Me.cmdCancel.Caption = "Cancel"
  Call List1_Click
  Me.cmdOK.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDel_Click
' Purpose           : Delete a selected verse from the list
'*******************************************************************************
Private Sub cmdDel_Click()
  Dim I As Long
  
  With Me.List1
    I = .ListIndex
    If MessageBox(Me, "Verify deletion of: " & UCase$(.List(I)) & ".", vbYesNo Or vbQuestion, "Verify Delete") = vbNo Then Exit Sub
    .RemoveItem I
    I = I - 1
    If .ListCount > 0 Then
      If I = -1 Then I = 0
      .ListIndex = I
    End If
    Call List1_Click
    Me.cmdOK.Enabled = True
  End With
  Me.cmdCancel.Caption = "Cancel"
  Call List1_Click
  Me.List1.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : Accept changes to the list
'*******************************************************************************
Private Sub cmdOK_Click()
  Dim Idx As Long
  Dim S As String
  
  With colFavs
    Do While .Count
      .Remove 1
    Loop
  End With
  
  For Idx = 0 To Favs - 1
    frmGrkXlate.mnuFavList(Idx).Visible = False
  Next Idx
  
  With Me.List1
    SaveSetting App.Title, "Settings", "FavCnt", CStr(.ListCount)
    For Idx = 1 To .ListCount
      S = .List(Idx - 1)
      SaveSetting App.Title, "Settings", "Fav" & CStr(Idx), S
      colFavs.Add S, S
      frmGrkXlate.mnuFavList(Idx - 1).Caption = S
      frmGrkXlate.mnuFavList(Idx - 1).Visible = True
    Next Idx
    frmGrkXlate.mnuFavSep.Visible = CBool(.ListCount)
    frmGrkXlate.mnuFavDel.Enabled = CBool(.ListCount)
  End With
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSort_Click
' Purpose           : Sort the list in alphabetical order
'*******************************************************************************
Private Sub cmdSort_Click()
  With Me.List1                   'copy into sorting list
    Do While .ListCount
      Me.List2.AddItem .List(0)
      .RemoveItem 0
    Loop
  End With
  With Me.List2                   'copy back
    Do While .ListCount
      Me.List1.AddItem .List(0)
      .RemoveItem 0
    Loop
  End With
  Me.cmdSort.Enabled = False
  Me.mnuPopupSortAlpha.Enabled = False
  Me.cmdSortBbl.Enabled = True
  Me.mnuPopupSortBbl.Enabled = True
  SortBbl = False
  Me.cmdOK.Enabled = True
  Me.List1.ListIndex = -1
  Me.cmdCancel.Caption = "Cancel"
  Call List1_Click
  Me.List1.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : CmdSortBbl_Click
' Purpose           : Sort the list in Bible order
'*******************************************************************************
Private Sub cmdSortBbl_Click()
  Dim S As String, Ary() As String, A As String
  Dim Idx As Long, I As Long, J As Long, Mx As Long
  
  With Me.List1                   'copy into sorting list
    Do While .ListCount
      Me.List2.AddItem .List(0)
      .RemoveItem 0
    Loop
  End With
  With Me.List2
    For Idx = 27 To 1 Step -1             'scan through books in reverse order
      Ary = Split(Books(Idx), ",")
      A = Ary(3)                          'get current book
      For J = .ListCount - 1 To 0 Step -1 'scan list in reverse
        S = .List(J)                      'get an entry
        I = InStrRev(S, " ")              'find space before Chp:Vrs
        If Left$(S, I - 1) = A Then       'in current book?
          .RemoveItem J                   'yes, so remove the entry
          Me.List1.AddItem S, 0           'add to start of master list
        End If
      Next J                              'check next entry
    Next Idx                              'check for next book
  End With
  Me.cmdSort.Enabled = True
  Me.mnuPopupSortAlpha.Enabled = True
  Me.cmdSortBbl.Enabled = False
  Me.mnuPopupSortBbl.Enabled = False
  SortBbl = True
  Me.cmdOK.Enabled = True
  Me.List1.ListIndex = -1
  Me.cmdCancel.Caption = "Cancel"
  Call List1_Click
  Me.List1.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Load the form
'*******************************************************************************
Private Sub Form_Load()
  Dim Idx As Long
  
  Set MyToolTips = New clsToolTip   'declare object
  With MyToolTips
    .Create Me                      'create object
    .MaxTipWidth = 1440 * 6         'set to 4 inches
    .DelayTime(ttDelayShow) = 20 * 1000 'set to 20 seconds
    .SetFont , 12
    .AddTool Me.List1
    .ToolText(Me.List1) = vbNullString
  End With
  
  With colFavs
    Favs = .Count
    For Idx = 1 To Favs
      Me.List1.AddItem colFavs(Idx)
    Next Idx
  End With
  
  Me.cmdDel.Enabled = CBool(Me.List1.ListCount)
  Me.mnuPopupDel.Enabled = Me.cmdDel.Enabled
  Me.cmdClear.Enabled = Me.cmdDel.Enabled
  Me.mnuPopupClear.Enabled = Me.cmdClear.Enabled
  Me.cmdSort.Enabled = Me.cmdDel.Enabled
  Me.mnuPopupSortAlpha.Enabled = Me.cmdDel.Enabled
  If Me.cmdDel.Enabled Then Me.List1.ListIndex = 0
  Me.cmdOK.Enabled = False
  SortBbl = CBool(GetSetting(App.Title, "Settings", "SortBbl", "True"))
  Me.cmdSort.Enabled = SortBbl
  Me.cmdSortBbl.Enabled = Not SortBbl
  Me.mnuPopupSortAlpha.Enabled = SortBbl
  Me.mnuPopupSortBbl.Enabled = Not SortBbl
  
  Me.List1.BackColor = cLight
  
  With frmGrkXlate
    Me.Top = .Top
    Me.Left = (.Width - Me.Width) / 2 + .Left
    Me.Height = .Height - Screen.TwipsPerPixelY * 4
  End With
  Call List1_Click
  Me.mnuPopup.Visible = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Repaint the background
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize the form
'*******************************************************************************
Private Sub Form_Resize()
  Me.List1.Height = Me.ScaleHeight - Me.List1.Top - 120
  Me.cmdCancel.Top = Me.List1.Top + Me.List1.Height - Me.cmdCancel.Height
  Me.cmdOK.Top = Me.cmdCancel.Top - Me.cmdOK.Height - 120
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Remove form's resources
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  SaveSetting App.Title, "Settings", "SortBbl", CStr(SortBbl)
  Set MyToolTips = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : List1_Click
' Purpose           : Process clicks in the list
'*******************************************************************************
Private Sub List1_Click()
  Me.cmdDel.Enabled = Me.List1.SelCount <> 0
  Me.mnuPopupDel.Enabled = Me.cmdDel.Enabled
  Me.cmdClear.Enabled = Me.List1.ListCount <> 0
  Me.mnuPopupClear.Enabled = Me.cmdClear.Enabled
  If Not Me.cmdClear.Enabled Then
    Me.cmdSortBbl.Enabled = False
    Me.cmdSortBbl.Enabled = False
    Me.mnuPopupSortAlpha.Enabled = False
    Me.mnuPopupSortBbl.Enabled = False
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : List1_DblClick
' Purpose           : Go to selected verse
'*******************************************************************************
Private Sub List1_DblClick()
  Dim Cur As String, S As String, Ary() As String
  Dim Idx As Long
  
  With Me.List1
    Cur = .List(.ListIndex)
  End With
  Idx = InStrRev(Cur, " ")
  S = Left$(Cur, Idx - 1)
  Cur = Mid$(Cur, Idx + 1)
  Idx = InStr(1, Cur, ":")
  Chp = CLng(Left$(Cur, Idx - 1))
  Vrs = CLng(Mid$(Cur, Idx + 1))
  
  For Bk = 1 To 27
    Ary = Split(Books(Bk), ",")
    If S = Ary(3) Then Exit For
  Next Bk
  
  ChpCnt = CLng(Ary(4))                   'get the chapter count
  Call frmGrkXlate.GetVerseCount          'get the verse count
  Call frmGrkXlate.UpdateVerse            'display the verse
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : List1_KeyDown
' Purpose           : Allow typing Del key to delete an entry
'*******************************************************************************
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    KeyCode = 0
    If Me.List1.SelCount <> 0 Then
      Me.cmdDel.Value = True
    End If
  End If
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
  Dim Cur As String, S As String, Ary() As String
  Dim Idx As Long, B As Long, C As Long, V As Long
  
  Cur = Trim$(GetStringFromMouseMove(Me.List1, X, Y))  'get line data
  If Len(Cur) <> 0 Then
    Idx = InStr(3, Cur, " ") - 1
    For B = 1 To 27
      Ary = Split(Books(B), ",")
      If Left$(Cur, Idx) = Ary(3) Then Exit For
    Next B
    S = Mid$(Cur, Idx + 2)
    Idx = InStr(1, S, ":")
    C = CLng(Left$(S, Idx - 1))
    V = CLng(Mid$(S, Idx + 1))
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
      Cur = Cur & " (" & S & ")  " & Mid$(Bible(Idx), 8)                 'strip off leader
    End If
  End If
  S = MyToolTips.ToolText(Me.List1)
  If S <> Cur Then MyToolTips.ToolText(Me.List1) = Cur
End Sub

Private Sub mnuPopupClear_Click()
  Me.cmdClear.Value = True
End Sub

Private Sub mnuPopupDel_Click()
  Me.cmdDel.Value = True
End Sub

Private Sub mnuPopupSortAlpha_Click()
  Me.cmdSort.Value = True
End Sub

Private Sub mnuPopupSortBbl_Click()
  Me.cmdSortBbl.Value = True
End Sub
