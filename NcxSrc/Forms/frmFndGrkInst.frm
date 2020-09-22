VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmFndGrkInst 
   Caption         =   "Find All Instances of Current Greek Word"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   Icon            =   "frmFndGrkInst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleMode       =   0  'User
   ScaleWidth      =   9061.539
   Begin VB.CheckBox chkMatchExact 
      Height          =   195
      Left            =   6360
      TabIndex        =   9
      ToolTipText     =   "When checked, no word-ending variations are considered"
      Top             =   540
      Width           =   195
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   6780
      TabIndex        =   1
      ToolTipText     =   "Copy list to the clipboard"
      Top             =   6120
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7860
      TabIndex        =   2
      ToolTipText     =   "Close search"
      Top             =   6120
      Width           =   795
   End
   Begin VB.ListBox lstSearch 
      Height          =   2205
      Left            =   180
      MouseIcon       =   "frmFndGrkInst.frx":0C62
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3780
      Width           =   8475
   End
   Begin RichTextLib.RichTextBox rtbNotes 
      Height          =   2475
      Left            =   180
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   780
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   4366
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      OLEDropMode     =   0
      TextRTF         =   $"frmFndGrkInst.frx":0DB4
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
   Begin VB.Label lblSizer 
      BackStyle       =   0  'Transparent
      Caption         =   "o"
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
      Left            =   8580
      TabIndex        =   11
      Top             =   6540
      Width           =   315
   End
   Begin VB.Label lblExactMatch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Perform EXACT Greek match"
      Height          =   195
      Left            =   6600
      TabIndex        =   10
      ToolTipText     =   "When checked, no word-ending variations are considered"
      Top             =   540
      Width           =   2070
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   246.154
      X2              =   8861.539
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click line to activate it in the main application"
      ForeColor       =   &H80000015&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   6000
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Greek Word Found In:"
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
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   3540
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Definition:"
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
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   540
      Width           =   885
   End
   Begin VB.Label lblGreek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Greek Word"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Greek Word:"
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
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmFndGrkInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyToolTips As clsToolTip
Private Cap As String
Private MatchCol As Collection
Private Gwd As String
Private GwdV As String
Private IsLoading As Boolean

Private Sub chkMatchExact_Click()
  If IsLoading Then Exit Sub
  SaveSetting App.Title, "Settings", "GrkExactMatch", CStr(Me.chkMatchExact.Value)
  Screen.MousePointer = vbDefault
  DoEvents
  DoSearch
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopy_Click
' Purpose           : Copy list to the clipboard
'*******************************************************************************
Private Sub cmdCopy_Click()
  Dim Idx As Long
  Dim S As String
  
  With Me.lstSearch
    For Idx = 0 To .ListCount - 1
      S = S & vbCrLf & .List(Idx)
    Next Idx
  End With
  Clipboard.Clear
  Clipboard.SetText Mid$(S, 2)
  Me.lstSearch.SetFocus
End Sub

Private Sub Form_Load()
  Dim Ary() As String
  
  IsLoading = True
  Set MatchCol = New Collection
  Cap = Me.Caption
  Me.chkMatchExact.Value = CLng(GetSetting(App.Title, "Settings", "GrkExactMatch", "0"))
  
  Set MyToolTips = New clsToolTip
  With MyToolTips
    .Create Me              'create object
    .MaxTipWidth = 1440 * 6 'set to 4 inches
    .DelayTime(ttDelayShow) = 20 * 1000 'set to 20 seconds
    .SetFont , 12
    .AddTool Me.lstSearch
    .ToolText(Me.lstSearch) = vbNullString
  End With
  Me.rtbNotes.BackColor = cVLight

  With frmGrkXlate
    Me.lblGreek.Caption = .lstGrkWords.List(.lstGrkWords.ListIndex)
    Me.rtbNotes.TextRTF = .rtbNotes.TextRTF
    GwdV = " " & BBLLine(.lstGrkWords.ListIndex + 1) & " "
    Ary = Split(Grk(VrsIdx), " ")
    Gwd = " " & Ary(.lstGrkWords.ListIndex + 1) & " "
  End With
  With frmGrkXlate
    Me.Top = .Top
    Me.Height = CLng(GetSetting(App.Title, "Settings", "GrkSrchHt", CStr(Me.Height)))
    If Me.Top + Me.Height > Screen.Height Then Me.Height = .Height
    Me.Width = CLng(GetSetting(App.Title, "Settings", "GrkSrchwt", CStr(Me.Width)))
    Me.Left = .Left + .Width - Me.Width
'    If Me.Left + Me.Width > Screen.Width Then Me.Width = .Width
  End With
  DoSearch
  IsLoading = False
End Sub

Private Sub DoSearch()
  Dim Ary() As String, S As String
  Dim Idx As Long, mBk As Long, mChp As Long, mVrs As Long
  
  With MatchCol
    Do While .Count
      .Remove 1
    Loop
  End With
  
  With frmGrkXlate
    Me.lblGreek.Caption = .lstGrkWords.List(.lstGrkWords.ListIndex)
    Me.rtbNotes.TextRTF = .rtbNotes.TextRTF
  End With
  With Me.lstSearch
    .Clear
    If Me.chkMatchExact.Value = 0 Then
      For Idx = 0 To UBound(GrkBBL) - 1
        If InStr(1, GrkBBL(Idx) & " ", GwdV) Then
          S = Bible(Idx)
          MatchCol.Add Left$(S, 6)
          mBk = CLng(Left$(S, 2))
          mChp = CLng(Mid$(S, 3, 2))
          mVrs = CLng(Mid$(S, 5, 2))
          Ary = Split(Books(mBk), ",")
          .AddItem Ary(3) & " " & CStr(mChp) & ":" & CStr(mVrs) & "  " & Mid$(S, 8)
        End If
      Next Idx
      Me.Caption = Cap & " - Verse Matches Found: " & CStr(MatchCol.Count)
    Else
      For Idx = 0 To UBound(Grk) - 1
        If InStr(1, Grk(Idx) & " ", Gwd) Then
          S = Bible(Idx)
          MatchCol.Add Left$(S, 6)
          mBk = CLng(Left$(S, 2))
          mChp = CLng(Mid$(S, 3, 2))
          mVrs = CLng(Mid$(S, 5, 2))
          Ary = Split(Books(mBk), ",")
          .AddItem Ary(3) & " " & CStr(mChp) & ":" & CStr(mVrs) & "  " & Mid$(S, 8)
        End If
      Next Idx
      Me.Caption = Cap & " - Exact Verse Matches Found: " & CStr(MatchCol.Count)
    End If
    If MatchCol.Count = 0 Then
      Me.lstSearch.BackColor = vb3DLight
    Else
      Me.lstSearch.BackColor = vbWhite
    End If
  End With
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)
End Sub

Private Sub Form_Resize()
  Static Resizing As Boolean
  
  If Resizing Then Exit Sub
  Resizing = True
  If Me.Width < 5500 Then Me.Width = 5500
  If Me.Height < 7000 Then Me.Height = 7000
  Me.cmdClose.Top = Me.ScaleHeight - Me.lblSizer.Height / 2 - Me.cmdClose.Height
  Me.cmdCopy.Top = Me.cmdClose.Top
  Me.lstSearch.Height = Me.cmdClose.Top - Me.lstSearch.Top - 150
  Me.lblNote.Top = Me.lstSearch.Top + Me.lstSearch.Height
  With frmGrkXlate
    Me.Top = .Top
    Me.Left = .Left + .Width - Me.Width
    If Me.Height > .Height Then Me.Height = .Height
  End With
  Me.lblSizer.Left = Me.ScaleWidth - Me.lblSizer.Width
  Me.lblSizer.Top = Me.ScaleHeight - Me.lblSizer.Height
  Me.cmdClose.Left = Me.ScaleWidth - Me.cmdClose.Width - Me.lstSearch.Left
  Me.cmdCopy.Left = Me.lstSearch.Left
  Me.lstSearch.Width = Me.ScaleWidth - Me.lstSearch.Left * 2
  Me.rtbNotes.Width = Me.ScaleWidth - Me.rtbNotes.Left - Me.lstSearch.Left
  Me.lblExactMatch.Left = Me.ScaleWidth - Me.lblExactMatch.Width - Me.lstSearch.Left
  Me.chkMatchExact.Left = Me.lblExactMatch.Left - Me.chkMatchExact.Width - 30
  Resizing = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set MatchCol = Nothing
  Set MyToolTips = Nothing
  SaveSetting App.Title, "Settings", "GrkSrchHt", CStr(Me.Height)
  SaveSetting App.Title, "Settings", "GrkSrchWt", CStr(Me.Width)
End Sub

Private Sub lblExactMatch_Click()
  With chkMatchExact
    .Value = Abs(.Value - 1)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lstSearch_MouseMove
' Purpose           : update tooltip as required
'*******************************************************************************
Private Sub lstSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String, T As String
  
  S = Trim$(GetStringFromMouseMove(Me.lstSearch, X, Y))  'get line data
  T = MyToolTips.ToolText(Me.lstSearch)           'get current tooltip (grabs max 80 chars)
  If Len(T) = 0 And Len(S) <> 0 Then
    MyToolTips.ToolText(Me.lstSearch) = S
  ElseIf Len(S) = 0 And Len(T) <> 0 Then
    MyToolTips.ToolText(Me.lstSearch) = S
  ElseIf Left$(S, Len(T)) <> T Then
    MyToolTips.ToolText(Me.lstSearch) = S
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lstSearch_Click
' Purpose           : Update main form with selected verse
'*******************************************************************************
Private Sub lstSearch_Click()
  Dim S As String, Ary() As String
  Dim Idx As Long
  
  With Me.lstSearch
    If Me.chkMatchExact.Value = 0 Then
      Me.Caption = Cap & " - Match " & CStr(.ListIndex + 1) & " of " & CStr(.ListCount)
    Else
      Me.Caption = Cap & " - Exact Match " & CStr(.ListIndex + 1) & " of " & CStr(.ListCount)
    End If
    S = MatchCol(.ListIndex + 1)
  End With
  Bk = CLng(Left$(S, 2))          'grab book
  Ary = Split(Books(Bk), ",")
  ChpCnt = CLng(Ary(4))           'grab chapter count
  Chp = CLng(Mid$(S, 3, 2))       'grab chapter
  Vrs = CLng(Right$(S, 2))        'grab verse
  With frmGrkXlate
    Call .GetVerseCount  'grab verse count
    Call .UpdateVerse
    If Me.chkMatchExact.Value = 0 Then
      Ary = Split(Mid$(GrkBBL(VrsIdx), 8), " ")
      S = Trim$(GwdV)
      For Idx = 0 To UBound(Ary)
        If Ary(Idx) = S Then
          .lstGrkWords.ListIndex = Idx
          Exit For
        End If
      Next Idx
    Else
      Ary = Split(Mid$(Grk(VrsIdx), 8), " ")
      S = Trim$(Gwd)
      For Idx = 0 To UBound(Ary)
        If Ary(Idx) = S Then
          .lstGrkWords.ListIndex = Idx
          Exit For
        End If
      Next Idx
    End If
  End With
End Sub

