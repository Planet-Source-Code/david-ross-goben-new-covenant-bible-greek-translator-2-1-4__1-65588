VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKJVDict 
   Caption         =   "KJV Dictionary"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   Icon            =   "frmKJVDict.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   6885
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1620
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5460
      Width           =   795
   End
   Begin VB.CommandButton cmdWhatIsThis 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "What is this form?"
      Top             =   240
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5220
      Top             =   5460
   End
   Begin VB.CommandButton cmdFind 
      Height          =   315
      Left            =   6180
      Picture         =   "frmKJVDict.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Find a particular word in the lists (in any list) (Ctrl-F)"
      Top             =   240
      Width           =   315
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   9075
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10716
            Text            =   "Double-click a line to save its contents to the clipboard"
            TextSave        =   "Double-click a line to save its contents to the clipboard"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvAlpha 
      Height          =   4755
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "KJV Word"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Modern Definition for KJV Word"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.TabStrip tsAlpha 
      Height          =   5175
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9128
      TabWidthStyle   =   2
      TabFixedWidth   =   422
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   24
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "A"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "B"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "C"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "D"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "E"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "F"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "G"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "H"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "I"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "J"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "K"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "L"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "M"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "N"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "O"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "P"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Q"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab18 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "R"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab19 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "S"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab20 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "T"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab21 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "U"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab22 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "V"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab23 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "W"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab24 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Y"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type a letter or click an alphabet tab.  Up/Down and PageUp/PageDown Scrolls."
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5805
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "LblWidth"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   630
   End
End
Attribute VB_Name = "frmKJVDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const AlphaList As String = "ABCDEFGHIJKLMNOPQRSTUVWY"
Private OldTab As Long
Private MyToolTips(23) As clsToolTip

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Ary() As String, S As String, T As String
  Dim Idx As Long, I As Long, J As Long
  Dim Lf As Long, Wd As Long, Tp As Long, Ht As Long
  Dim LI As MSComctlLib.ListItem
  Dim MinWidth As Long
  
  Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\KJVDict.txt", ForReading, False)
  Ary = Split(ts.ReadAll, vbCrLf)
  ts.Close
  Me.cmdCancel.Left = -2440
'
' load all the listview controls we will neeed
'
  For Idx = 1 To Len(AlphaList)
    S = Mid$(AlphaList, Idx, 1)               'get a character
    If Idx > 1 Then                           'if not index 0
      Load lvAlpha(Idx - 1)                   'load a new one
      Me.lvAlpha(Idx - 1).Top = Me.lvAlpha(0).Top
      Me.lvAlpha(Idx - 1).Left = Me.lvAlpha(0).Left
    End If
    Me.lvAlpha(Idx - 1).BackColor = cVLight
    
    Set MyToolTips(Idx - 1) = New clsToolTip  'add a tooltip for the listview
    With MyToolTips(Idx - 1)
      .Create Me                              'create object
      .MaxTipWidth = 1440 * 2                 'width max = 2 inches
      .DelayTime(ttDelayShow) = 20 * 1000     'set to 20 seconds
      .SetFont , 10
      .AddTool Me.lvAlpha(Idx - 1)
      .ToolText(Me.lvAlpha(Idx - 1)) = vbNullString
    End With
  Next Idx
  
  Me.Caption = Me.Caption & " - " & CStr(UBound(Ary)) & " words"
  
  MinWidth = 1440                             'define minimum column 1 width to 1 inch
  For Idx = 1 To UBound(Ary) - 1              'scan the data (line 0 is a header)
    S = Ary(Idx)                              'grab an entry
    I = InStr(1, AlphaList, Left$(S, 1)) - 1  'get an index to its listview
    J = InStr(1, S, vbTab)                    'find the tab separator
    T = Left$(S, J - 1)                       'grab the KJV word or phrase
    Me.lblwidth.Caption = T
    If MinWidth < Me.lblwidth.Width Then MinWidth = Me.lblwidth.Width 'expand minwiddth
    Set LI = Me.lvAlpha(I).ListItems.Add(, T, T)  'add an intem to the listview
    LI.SubItems(1) = Mid$(S, J + 1)           'add the modern definition
  Next Idx
  
  J = Me.lvAlpha(0).Width - MinWidth - 360    'compute column 2 width
  For Idx = 1 To Len(AlphaList)
    Me.lvAlpha(Idx - 1).ColumnHeaders(1).Width = MinWidth 'set column 1
    Me.lvAlpha(Idx - 1).ColumnHeaders(2).Width = J        'and column to
    Me.tsAlpha.Tabs(Idx).ToolTipText = CStr(Me.lvAlpha(Idx - 1).ListItems.Count) & " words"
  Next Idx
  
  Call GetScreenWorkArea(Lf, Wd, Tp, Ht)  'get screen work area sizing
  Me.Left = Lf
  Me.Top = Tp
  Me.Height = Ht
  OldTab = -1
  With frmGrkXlate.lstWords
    S = vbNullString
    If .ListCount <> 0 Then
      S = .List(.ListIndex)
      If Not FindMatch(S) Then S = vbNullString
    End If
  End With
  If Len(S) = 0 Then
    Set Me.tsAlpha.SelectedItem = Me.tsAlpha.Tabs(1)        'set "A" for top tab
'''    Me.StatusBar1.Panels(1).Text = CStr(Me.lvAlpha(0).ListItems.Count) & " words"
'''    Me.lvAlpha(0).Visible = True                            'show its list
  End If
  ShowKJVDict = True
  Me.Timer1.Enabled = True
  Me.Show
  frmGrkXlate.mnuWinKJV.Enabled = True
  frmGrkXlate.CheckWin
End Sub

Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
End Sub

Private Sub Form_Resize()
  Dim I As Long, J As Long
  Static Resizing As Boolean
  
  If Me.WindowState = vbMinimized Then
    frmGrkXlate.ZOrder 0
    frmGrkXlate.SetFocus
    Exit Sub 'ignore if minimized
  End If
  If Resizing Then Exit Sub                     'ignore if we are now processing
  Resizing = True                               'indicate we are processing
  If Me.Width < 7000 Then Me.Width = 7000       'avoid going smaller than minimum
  If Me.Height < 7000 Then Me.Height = 7000
  I = Me.tsAlpha.Height - Me.lvAlpha(0).Height
  Me.tsAlpha.Width = Me.ScaleWidth - Me.tsAlpha.Left * 2
  Me.tsAlpha.Height = Me.ScaleHeight - Me.StatusBar1.Height - Me.tsAlpha.Top - 60
  Me.lvAlpha(0).Width = Me.ScaleWidth - Me.lvAlpha(0).Left * 2
  Me.lvAlpha(0).Height = Me.tsAlpha.Height - I
  J = Me.lvAlpha(0).Width - Me.lvAlpha(0).ColumnHeaders(1).Width - 360
  For I = 1 To Len(AlphaList)
    If I > 1 Then
      Me.lvAlpha(I - 1).Width = Me.lvAlpha(0).Width
      Me.lvAlpha(I - 1).Height = Me.lvAlpha(0).Height
    End If
    Me.lvAlpha(I - 1).ColumnHeaders(2).Width = J
  Next I
  Me.cmdFind.Left = Me.ScaleWidth - Me.cmdFind.Width - Me.tsAlpha.Left - 30
  Me.cmdWhatIsThis.Left = Me.cmdFind.Left - Me.cmdWhatIsThis.Width
  Resizing = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim Idx As Long
  '
  ' release all tooltip resources
  '
  For Idx = 0 To 23               '(1-24; X and Z not present)
    Set MyToolTips(Idx) = Nothing 'release
  Next Idx
  ShowKJVDict = False
  frmGrkXlate.mnuWinKJV.Enabled = False
  frmGrkXlate.CheckWin
End Sub

'*******************************************************************************
' Subroutine Name   : Form_KeyDown
' Purpose           : Keyboard checking
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim C As String
  Dim I As Long
  
  Select Case Shift
    Case vbCtrlMask                 'check for Ctrl-F (Find)
      If KeyCode = vbKeyF Then
        KeyCode = 0
        Me.cmdFind.Value = True
      End If
    Case vbKeyShift, 0              'check for valid Alpha characters
      C = UCase$(Chr$(KeyCode))
      I = InStr(1, AlphaList, C)    'in list (A-Y, less X and Z)
      If I <> 0 Then                'found it
        KeyCode = 0                 'disable further key processing
        Set Me.tsAlpha.SelectedItem = Me.tsAlpha.Tabs(I)  'select tab
        Me.lvAlpha(I - 1).SetFocus                        'set focust to proper listview
      End If
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : lvAlpha_DblClick
' Purpose           : Save line contents to the clipboard
'*******************************************************************************
Private Sub lvAlpha_DblClick(Index As Integer)
  Dim lItm As MSComctlLib.ListItem

  Set lItm = Me.lvAlpha(Index).SelectedItem
  If lItm Is Nothing Then Exit Sub            'no, so ignore
  Clipboard.Clear
  Clipboard.SetText lItm.Text & ": " & lItm.SubItems(1), vbCFText
End Sub

'*******************************************************************************
' Subroutine Name   : lvAlpha_MouseMove
' Purpose           : Displat text in tooltip if longer than displayed
'*******************************************************************************
Private Sub lvAlpha_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String
  Dim lItm As MSComctlLib.ListItem
  Static LastTip As String
  
  Set lItm = Me.lvAlpha(Index).HitTest(X, Y)  'anything hovered over?
  If lItm Is Nothing Then Exit Sub            'no, so ignore
  S = lItm.SubItems(1)                        'get text
  If S <> LastTip Then
    LastTip = S
    Me.lblwidth.Caption = S                   'test width
    If Me.lblwidth.Width < Me.lvAlpha(Index).ColumnHeaders(2).Width - 240 Then S = vbNullString
    MyToolTips(Index).ToolText(Me.lvAlpha(Index)) = S
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Timer1_Timer
' Purpose           : 1-shot time. This simply ensures that "A" has focus
'*******************************************************************************
Private Sub Timer1_Timer()
  Me.Timer1.Enabled = False
  Me.lvAlpha(OldTab).SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : tsAlpha_Click
' Purpose           : Check for a new tab being selected
'*******************************************************************************
Private Sub tsAlpha_Click()
  Dim NewTab As Long
  
  NewTab = Me.tsAlpha.SelectedItem.Index - 1  'check new index
  If NewTab = OldTab Then Exit Sub            'same as previous? Ignore if so
  On Error Resume Next
  Me.lvAlpha(OldTab).Visible = False          'hide old list
  Me.lvAlpha(NewTab).Visible = True           'show new
  Me.lvAlpha(NewTab).ZOrder 0                 'ensure layered top-most
  Me.StatusBar1.Panels(1).Text = CStr(Me.lvAlpha(NewTab).ListItems.Count) & " words"
  OldTab = NewTab                             'make new current tab
  Me.lvAlpha(NewTab).SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : cmdFind_Click
' Purpose           : Find a word in the KJV list
'*******************************************************************************
Private Sub cmdFind_Click()
  Dim S As String
  
  S = Trim$(InputMsgBox(Me, "Enter word or phrase to find:", "Find KJV word", vbNullString))
  If Len(S) = 0 Then Exit Sub
  
  If Not FindMatch(S) Then
    MessageBox Me, "Cannot find selected word or phrase: " & S, _
               vbOKOnly Or vbExclamation, "text Not Found"
  End If
End Sub

Private Function FindMatch(Text As String) As Boolean
  Dim I As Long, Idx As Long
  Dim FoundMatch As Boolean
  
  I = InStr(1, AlphaList, UCase$(Left$(Text, 1)))
  If I <> 0 Then
    With Me.lvAlpha(I - 1)
      For Idx = 1 To .ListItems.Count
        If StrComp(Text, .ListItems(Idx), vbTextCompare) = 0 Then
          Set Me.tsAlpha.SelectedItem = Me.tsAlpha.Tabs(I)
          OldTab = I - 1
          Set Me.lvAlpha(OldTab).SelectedItem = Me.lvAlpha(OldTab).ListItems(Idx)
          Me.lvAlpha(OldTab).SelectedItem.EnsureVisible
          FoundMatch = True
          On Error Resume Next
          Me.lvAlpha(OldTab).SetFocus
          Exit For
        End If
      Next Idx
    End With
  End If
  
  FindMatch = FoundMatch
End Function

'*******************************************************************************
' Subroutine Name   : cmdWhatIsThis_Click
' Purpose           : Help explaining this form
'*******************************************************************************
Private Sub cmdWhatIsThis_Click()
  MessageBox Me, "This form displays a dictionary of words used in the King James Version (KJV)" & vbCrLf & _
                 "whose meaning are vague or have changed, sometime dramatically.  You can refer" & vbCrLf & _
                 "to this table to verify proper intepretation of verses." & vbCrLf & vbCrLf & _
                 "Type a letter to change tabs, and arrows or page-up/page-down to scroll." & vbCrLf & vbCrLf & _
                 "Use FIND (Ctrl-F) to search for a word in the dictionary (you do not need to be" & vbCrLf & _
                 "on any particular tab to find a word).", vbOKOnly Or vbExclamation, "What Is This Form?"
End Sub


