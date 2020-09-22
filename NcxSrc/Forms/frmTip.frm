VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   5190
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   7335
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSequential 
      Caption         =   "save whether or not this form should be displayed at startup"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   4320
      Width           =   195
   End
   Begin VB.TextBox txtFocus 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "txtFocus"
      Top             =   4140
      Width           =   795
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "save whether or not this form should be displayed at startup"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   4080
      Width           =   195
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      ToolTipText     =   "Display another tip"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   7035
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "Clear viewed tip list"
         Top             =   3480
         Width           =   675
      End
      Begin VB.ListBox lstTips 
         Height          =   2595
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Select a particular tip number"
         Top             =   840
         Width           =   675
      End
      Begin RichTextLib.RichTextBox txtTipText 
         Height          =   3195
         Left            =   1260
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5636
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmTip.frx":0000
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
      Begin VB.Label lblTipList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip List"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   660
         Width           =   510
      End
      Begin VB.Image imgClick 
         Height          =   615
         Left            =   240
         Top             =   60
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   1080
         X2              =   7020
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   300
         Picture         =   "frmTip.frx":008B
         Top             =   120
         Width           =   480
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         FillColor       =   &H80000010&
         FillStyle       =   0  'Solid
         Height          =   3795
         Left            =   0
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Did you know..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A00000&
         Height          =   330
         Left            =   1320
         TabIndex        =   9
         Top             =   120
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      ToolTipText     =   "Close the Tip of the Day window"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblchkSequential 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Display tips sequentially, not randomly"
      Height          =   195
      Left            =   420
      TabIndex        =   7
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000016&
      X1              =   120
      X2              =   7200
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   7200
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label lblAGC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """A Gnostic Cycle: Exploring the Origin of Christianity."""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A00000&
      Height          =   240
      Left            =   780
      MouseIcon       =   "frmTip.frx":0955
      MousePointer    =   99  'Custom
      TabIndex        =   13
      ToolTipText     =   "Available from AuthorHouse.com (www.authorhouse.com/BookStore/ItemDetail~bookid~33204.aspx)"
      Top             =   4860
      Width           =   5490
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Background information for non-program tid-bits can be found in the book, "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00505050&
      Height          =   555
      Left            =   120
      TabIndex        =   12
      Top             =   4620
      Width           =   7095
   End
   Begin VB.Label lblchkLoadTipsAtStartup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Show tips at application startup"
      Height          =   195
      Left            =   420
      TabIndex        =   5
      Top             =   4080
      Width           =   2205
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As Collection
' Name of tips file
Const TIP_FILE = "DB\TipOfDay.txt"
' Index in collection of tip currently being displayed.
Dim CurrentTip As Long

'*******************************************************************************
' Subroutine Name   : chkSequential_Click
' Purpose           : Save seuqntial Flag
'*******************************************************************************
Private Sub chkSequential_Click()
  SaveSetting App.Title, "Settings", "Display Tips Sequentially", Me.chkSequential.Value
End Sub

'*******************************************************************************
' Subroutine Name   : cmdReset_Click
' Purpose           : Reset viewed tip list
'*******************************************************************************
Private Sub cmdReset_Click()
  TipData = String$(Tips.Count, " ")
  Mid$(TipData, CurrentTip, 1) = "X"            'mark current tip as viewed
  Me.cmdReset.ToolTipText = "Clear viewed tip list (1 viewed so far)"
  Me.cmdReset.Enabled = False
  MessageBox Me, "Tip display list reset.", vbOKOnly Or vbInformation, "Reset Display List"
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Start up the form
'*******************************************************************************
Private Sub Form_Load()
  Dim ShowAtStartup As Long
'
' if the FileSystemObject is not loaded, then ignore this form for now
'
  If Fso Is Nothing Then
    Unload Me
    Exit Sub
  End If
'
' See if we should be shown at startup
'
  Set Tips = New Collection
  ShowAtStartup = CLng(GetSetting(App.Title, "Settings", "Show Tips at Startup", "1"))
  If ShowAtStartup = 0 Then
    Unload Me
    Exit Sub
  End If
'
' Set the checkbox, this will force the value to be written back out to the registry
'
  Me.chkLoadTipsAtStartup.Value = vbChecked
'
' get sequential flag
'
  Me.chkSequential.Value = CLng(GetSetting(App.Title, "Settings", "Display Tips Sequentially", "0"))
  CurrentTip = CLng(GetSetting(App.Title, "Settings", "LastTip", "0"))

'
' Seed Rnd
'
  Randomize
'
' Read in the tips file and display a tip at random.
'
  If LoadTips(App.Path & "\" & TIP_FILE) = False Then
    Me.txtTipText.Text = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
       "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
       "Then place it in the same directory as the application. "
  End If
  
  Me.lstTips.BackColor = cDark
  Me.Shape1.BorderColor = cDark
  Me.Shape1.FillColor = cDark
  Me.Picture1.BackColor = cDark
  Me.Picture1.BackColor = cVLight
  Me.txtTipText.BackColor = cVLight
  Me.Line1.BorderColor = cDark
  
  Me.txtFocus.Left = -1440  'hide focus field
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Unload the form
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  SaveSetting App.Title, "Settings", "TipList", TipData 'save the tip data fir nexr visit
  Set Tips = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : Form_MouseMove
' Purpose           : Allow a link to lose its underlining
'*******************************************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Me.lblAGC.FontUnderline Then Me.lblAGC.FontUnderline = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Set up the painted form background
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
End Sub

'*******************************************************************************
' Subroutine Name   : DoNextTip
' Purpose           : Select a new tip
'*******************************************************************************
Private Sub DoNextTip()
  Dim Idx As Long, I As Long
  
  If TipData = String$(Tips.Count, "X") Then    'clear list if all have been viewed
    TipData = String$(Tips.Count, " ")
  End If
  
  If Me.chkSequential.Value Then                  'if processing sequentially
    CurrentTip = CurrentTip + 1                   'simply go to next tip
    If CurrentTip > Len(TipData) Then CurrentTip = 1  'wrap if we went beyond the list
  Else
    CurrentTip = Int((Tips.Count * Rnd) + 1)      'Select a tip at random.
    Do While Mid$(TipData, CurrentTip, 1) <> " "  're-select if already shown
      CurrentTip = Int((Tips.Count * Rnd) + 1)
    Loop
  End If
  Mid$(TipData, CurrentTip, 1) = "X"            'mark current tip as viewed
  Me.lstTips.ListIndex = CurrentTip - 1
  SaveSetting App.Title, "Settings", "LastTip", CStr(CurrentTip)
'
' count number viewed
'
  I = 0
  For Idx = 1 To Len(TipData)
    If Mid$(TipData, Idx, 1) = "X" Then
      I = I + 1
    End If
  Next Idx
  Me.cmdReset.ToolTipText = "Clear viewed tip list (" & CStr(I) & " viewed so far)"
  Me.cmdReset.Enabled = I > 1
End Sub

'*******************************************************************************
' Function Name     : LoadTips
' Purpose           : Load the Tips list from a file
'*******************************************************************************
Function LoadTips(sFile As String) As Boolean
  Dim NextTip As String   ' Each tip read in from file.
  Dim Ary() As String     ' storage for data
  Dim Idx As Long
'
' Make sure a file is specified.
'
  If Len(sFile) = 0 Then
    LoadTips = False
    Exit Function
  End If
'
' Make sure the file exists before trying to open it.
'
  If Not Fso.FileExists(sFile) Then
    LoadTips = False
    Exit Function
  End If
'
' read the file
'
  Set ts = Fso.OpenTextFile(sFile, ForReading, False)
  Ary = Split(ts.ReadAll, vbCrLf)
  ts.Close
'
' parse it
'
  With Me.lstTips
    .Clear
    For Idx = 0 To UBound(Ary)
      NextTip = Trim$(Ary(Idx))
      If Len(NextTip) Then
        Tips.Add NextTip
        .AddItem CStr(Tips.Count)
      End If
    Next Idx
    .ListIndex = -1
  End With
'
' initialize the tip data as needed
'
  TipData = GetSetting(App.Title, "Settings", "TipList", vbNullString)
  If Len(TipData) = 0 Then TipData = String$(Tips.Count, " ") 'if new, then create
'
' Display a tip at random.
'
  DoNextTip
  LoadTips = True
End Function

'*******************************************************************************
' Subroutine Name   : chkLoadTipsAtStartup_Click
' Purpose           : Save whether or not this form should be displayed at startup
'*******************************************************************************
Private Sub chkLoadTipsAtStartup_Click()
  SaveSetting App.Title, "Settings", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub

'*******************************************************************************
' Subroutine Name   : cmdNextTip_Click
' Purpose           : Show the next tip
'*******************************************************************************
Private Sub cmdNextTip_Click()
  DoNextTip
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : User is satisfied and wants to continue with the app
'*******************************************************************************
Private Sub cmdOK_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : DisplayCurrentTip
' Purpose           : Display the currently selected tip
'*******************************************************************************
Public Sub DisplayCurrentTip()
  Dim S As String, T As String
  Dim I As Long, J As Long, SS As Long
  
  If Tips.Count > 0 Then
    Me.Caption = "Tip of the Day.    Tip # " & CStr(CurrentTip) & " of " & CStr(Tips.Count) & "."
    S = Tips.Item(CurrentTip)                   'get text to show
    Me.txtTipText.Text = vbNullString           'ensure back to topp, if a previous scrolled
    With Me.txtTipText
      LockWindowUpdate .hwnd                    'prevent flash
      I = InStr(1, S, "{")                      'any Greek data?
      Do While I <> 0                           'yes
        If I > 1 Then                           'normal text before it?
          SS = Len(.Text)                       'yes, so append normal to end
          .SelStart = SS
          .SelText = Left$(S, I - 1)            'add normal text
          .SelStart = SS
          .SelLength = Len(.Text) - SS          'select new data
          .SelBold = False                      'not bold
          .SelFontName = "Times New Roman"      'and normal text
        End If
        J = InStr(I + 1, S, "}")                'find end of Greek
        If J = 0 Then Debug.Assert False
        T = Mid$(S, I + 1, J - I - 1)           'grab Greek text
        SS = Len(.Text)                         'append to end
        .SelStart = SS
        .SelText = T
        .SelStart = SS
        .SelLength = Len(.Text) - SS            'select new Greek text
        .SelBold = True                         'bold
        .SelFontName = "Symbol"                 'and symbol font
        S = Mid$(S, J + 1)                      'trim off displayed data
        I = InStr(1, S, "{")                    'any more Greek data?
      Loop
      If Len(S) <> 0 Then                       'any normal data left?
        SS = Len(.Text)                         'yes, so append to end
        .SelStart = SS
        .SelText = S
        .SelStart = SS
        .SelLength = Len(.Text) - SS
        .SelBold = False
        .SelFontName = "Times New Roman"        'make sure it is normal
      End If
      .SelStart = 0
      LockWindowUpdate 0
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Label2_MouseMove
' Purpose           : Turn off link underline if mouse not over link
'*******************************************************************************
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Me.lblAGC.FontUnderline Then Me.lblAGC.FontUnderline = False
End Sub

'*******************************************************************************
' Subroutine Name   : lblAGC_Click
' Purpose           : Pass click to main menu option
'*******************************************************************************
Private Sub lblAGC_Click()
  frmGrkXlate.mnuHLPVisitSponsor_Click
End Sub

'*******************************************************************************
' Subroutine Name   : lblAGC_MouseMove
' Purpose           : Turn on link underline if mouse over link
'*******************************************************************************
Private Sub lblAGC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not Me.lblAGC.FontUnderline Then Me.lblAGC.FontUnderline = True
End Sub

'*******************************************************************************
' Subroutine Name   : lblchkLoadTipsAtStartup_Click
' Purpose           : Duplicate Checkbox functionality
'*******************************************************************************
Private Sub lblchkLoadTipsAtStartup_Click()
  With Me.chkLoadTipsAtStartup
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = vbChecked
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lblchkSequential_Click
' Purpose           : Duplicate Checkbox functionality
'*******************************************************************************
Private Sub lblchkSequential_Click()
  With Me.chkSequential
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = vbChecked
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lstTips_Click
' Purpose           : Select a tip
'*******************************************************************************
Private Sub lstTips_Click()
  CurrentTip = Me.lstTips.ListIndex + 1         'set the selected tip as the current
  Mid$(TipData, CurrentTip, 1) = "X"            'mark it as viewed
  Me.DisplayCurrentTip                          'show it
End Sub

'*******************************************************************************
' Subroutine Name   : txtTipText_GotFocus
' Purpose           : Hide focus in textbox
'*******************************************************************************
Private Sub txtTipText_GotFocus()
  Me.txtFocus.SetFocus
End Sub
