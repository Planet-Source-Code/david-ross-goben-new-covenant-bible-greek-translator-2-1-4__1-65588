VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSaveBible 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Write Bible To a File"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   Icon            =   "frmSaveBible.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIncludeTheoNotes 
      Caption         =   "Check1"
      Height          =   195
      Left            =   420
      TabIndex        =   27
      Top             =   3060
      Width           =   195
   End
   Begin VB.ComboBox cboBumpFactor 
      Height          =   315
      ItemData        =   "frmSaveBible.frx":030A
      Left            =   3720
      List            =   "frmSaveBible.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   23
      ToolTipText     =   "Increase the point sizes of the fonts written out by additional points"
      Top             =   3300
      Width           =   615
   End
   Begin VB.CheckBox chkTextColor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Note Text Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Click to change the text color for personal notes"
      Top             =   2520
      Width           =   1995
   End
   Begin VB.CheckBox chkPNotesAbove 
      Height          =   195
      Left            =   660
      TabIndex        =   19
      Top             =   2820
      Width           =   195
   End
   Begin VB.CheckBox chkAddPNotes 
      Height          =   195
      Left            =   420
      TabIndex        =   17
      Top             =   2580
      Width           =   195
   End
   Begin VB.CheckBox chkCenterBkHeading 
      Height          =   195
      Left            =   420
      TabIndex        =   9
      Top             =   1620
      Width           =   195
   End
   Begin VB.CheckBox chkCenterChapHeading 
      Height          =   195
      Left            =   420
      TabIndex        =   11
      Top             =   1860
      Width           =   195
   End
   Begin VB.CheckBox chkAddNoteSpace 
      Height          =   195
      Left            =   660
      TabIndex        =   15
      Top             =   2340
      Width           =   195
   End
   Begin VB.CheckBox chkNewPage 
      Height          =   195
      Left            =   420
      TabIndex        =   7
      Top             =   1380
      Width           =   195
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7860
      TabIndex        =   21
      Top             =   3240
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   6720
      TabIndex        =   0
      Top             =   3240
      Width           =   915
   End
   Begin VB.CheckBox chkVerseLines 
      Height          =   195
      Left            =   420
      TabIndex        =   13
      Top             =   2100
      Width           =   195
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   8340
      TabIndex        =   5
      ToolTipText     =   "Browse for File or storage location"
      Top             =   960
      Width           =   435
   End
   Begin VB.OptionButton OptRtf 
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Save text in plain text format"
      Top             =   660
      Width           =   195
   End
   Begin VB.OptionButton OptRtf 
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Save file in formatted RTF encoding"
      Top             =   660
      Width           =   195
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8340
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblIncludeTheoNotes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Include Theological Notes at the end of each chapter."
      Height          =   195
      Left            =   660
      TabIndex        =   28
      Top             =   3060
      Width           =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   360
      X2              =   8700
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   360
      X2              =   8700
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bible will be based upon the currently active Bible:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   26
      Top             =   180
      Width           =   5235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "points."
      Height          =   195
      Left            =   4440
      TabIndex        =   25
      ToolTipText     =   "Increase the point sizes of the fonts written out by additional points"
      Top             =   3360
      Width           =   465
   End
   Begin VB.Label lblBumpDefault 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bump default font point size () for output by"
      Height          =   195
      Left            =   420
      TabIndex        =   24
      ToolTipText     =   "Increase the point sizes of the fonts written out by additional points"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label lblPNotesAbove 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Place personal notes above verse (otherwise it is placed below it)"
      Height          =   195
      Left            =   900
      TabIndex        =   18
      Top             =   2820
      Width           =   4605
   End
   Begin VB.Label lblAddPNotes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Personal Notes to verses to output (forces placing each verse on its own line)"
      Height          =   195
      Left            =   660
      TabIndex        =   16
      Top             =   2580
      Width           =   5790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Center Book headings on the page."
      Height          =   195
      Index           =   6
      Left            =   660
      TabIndex        =   8
      ToolTipText     =   "Place book titles in the center of the page line"
      Top             =   1620
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Center Chapter Headings on the page."
      Height          =   195
      Index           =   5
      Left            =   660
      TabIndex        =   10
      ToolTipText     =   "Place chapter titled titles in the center of the page line"
      Top             =   1860
      Width           =   2730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add a blank line below each verse to make room for notes."
      Height          =   195
      Index           =   4
      Left            =   900
      TabIndex        =   14
      Top             =   2340
      Width           =   4155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start each Book on a fresh page."
      Height          =   195
      Index           =   3
      Left            =   660
      TabIndex        =   6
      ToolTipText     =   "When a new book is started, begin on a new page, rather than at the bottom of the previous"
      Top             =   1380
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Place each verse on their own line (otherwise merge verses until a paragraph terminator is encountered)."
      Height          =   195
      Index           =   2
      Left            =   660
      TabIndex        =   12
      Top             =   2100
      Width           =   7350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plain Text File (*.txt)"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      ToolTipText     =   "Save the file in plain text, suitable for any text editor"
      Top             =   660
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rich Text File (*.rtf)"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Save file in RTF format with formatting, readable by WordPad, Word, and most all enhanced word processors"
      Top             =   660
      Width           =   1350
   End
   Begin VB.Label lblSavePath 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   420
      TabIndex        =   22
      Top             =   960
      Width           =   7815
   End
End
Attribute VB_Name = "frmSaveBible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const lRtf As Integer = 0
Private Const lTxt As Integer = 1
Private Const lNewP As Integer = 3
Private Const lCtrB As Integer = 6
Private Const lCtrC As Integer = 5
Private Const lNewL As Integer = 2
Private Const lBlnk As Integer = 4

Private Sub chkAddPNotes_Click()
  Me.chkPNotesAbove.Enabled = CBool(Me.chkAddPNotes.Value)
  Me.lblPNotesAbove.Enabled = Me.chkPNotesAbove.Enabled
  Me.chkTextColor.Enabled = Me.chkPNotesAbove.Enabled = True And Me.OptRtf(0).Value = True
End Sub

Private Sub chkTextColor_Click()
  If Me.chkTextColor.Value = vbChecked Then
    Me.chkTextColor.Value = vbUnchecked
    With Me.CommonDialog1
      .Flags = cdlCCFullOpen Or cdlCCRGBInit
      .Color = Me.chkTextColor.ForeColor
      .CancelError = True
      On Error Resume Next
      .ShowColor
      If Err.Number <> 0 Then Exit Sub
      On Error GoTo 0
      Me.chkTextColor.ForeColor = .Color
    End With
  End If
End Sub

Private Sub chkVerseLines_Click()
  Me.chkAddNoteSpace.Enabled = CBool(Me.chkVerseLines.Value)
  Me.Label1(4).Enabled = Me.chkAddNoteSpace.Enabled
End Sub

Private Sub cmdBrowse_Click()
  Dim S As String
  Dim bYes As Boolean
  Dim Idx As Long
  Do
    bYes = False
    With Me.CommonDialog1
      .Flags = cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNPathMustExist
      .DialogTitle = "Save Bible As..."
      .FileName = Me.lblSavePath.Caption
      If IsRTF Then
        .Filter = "Rich Text File (*.rtf)|*.rtf|All Files|*.*"
        .DefaultExt = ".rtf"
      Else
        .Filter = "Text Files (*.txt)|*.txt|All Files|*.*"
        .DefaultExt = ".txt"
      End If
      .CancelError = True
      On Error Resume Next
      .ShowSave
      If Err.Number <> 0 Then Exit Sub
      S = Trim$(.FileName)
      If Len(S) = 0 Then Exit Sub
    End With
    
    Err.Clear
    If Fso.FileExists(S) Then
      Select Case MessageBox(Me, "Select File already exists? Replace it?", vbYesNoCancel Or vbQuestion, "Replace File?")
        Case vbCancel
          Exit Sub
        Case vbYes
          Exit Do
      End Select
    Else
      Exit Do
    End If
  Loop
  
  Me.lblSavePath.Caption = S
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  SaveSetting App.Title, "Settings", "BumpFactor", CStr(Me.cboBumpFactor.ListIndex)
  FileName = Me.lblSavePath.Caption
  SaveSetting App.Title, "Settings", "SaveBible", FileName
  IsRTF = Me.OptRtf(0).Value
  SaveSetting App.Title, "Settings", "RtfSave", CStr(Abs(CLng(IsRTF)))
  VerseLines = CBool(Me.chkVerseLines.Value)
  SaveSetting App.Title, "Settings", "VerseLines", CStr(VerseLines)
  BookNewPage = CBool(Me.chkNewPage.Value)
  SaveSetting App.Title, "Settings", "NewPage", CStr(BookNewPage)
  SaveSetting App.Title, "Settings", "AddNoteSpace", CStr(Me.chkAddNoteSpace.Value)
  AddNoteSpace = False
  If Me.chkAddNoteSpace.Enabled Then AddNoteSpace = CBool(Me.chkAddNoteSpace.Value)
  CenterBkHeading = rtfLeft
  If Me.chkCenterBkHeading.Enabled And CBool(Me.chkCenterBkHeading.Value) Then CenterBkHeading = rtfCenter
  SaveSetting App.Title, "Settings", "CenterBkHeading", CStr(CenterBkHeading)
  CenterChapHeading = rtfLeft
  If Me.chkCenterChapHeading.Enabled And CBool(Me.chkCenterChapHeading.Value) Then CenterChapHeading = rtfCenter
  SaveSetting App.Title, "Settings", "CenterChapHeading", CStr(CenterChapHeading)
  SaveSetting App.Title, "Settings", "AddPNotes", CStr(Me.chkAddPNotes.Value)
  SaveSetting App.Title, "Settings", "PNotesAbove", CStr(Me.chkPNotesAbove.Value)
  IncludeTheoNotes = CBool(Me.chkIncludeTheoNotes.Value)
  SaveSetting App.Title, "Settings", "IncludeTheoNotes", CStr(Me.chkIncludeTheoNotes.Value)
  PNotesColor = Me.chkTextColor.ForeColor
  SaveSetting App.Title, "Settings", "PNotesColor", CStr(PNotesColor)
  AddPNotes = CBool(Me.chkAddPNotes.Value)
  If AddPNotes Then
    PNotesAbove = CBool(Me.chkPNotesAbove.Value)
    VerseLines = True
  Else
    PNotesAbove = False
  End If
  bCancel = False
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Idx As Integer
  Dim S As String
  
  Me.lblTitle.Caption = "Bible will be based upon the currently active Bible: " & VersionText
  
  Idx = 1 - Abs(CLng(CBool(GetSetting(App.Title, "Settings", "RtfSave", "0"))))
  Me.OptRtf(Idx).Value = True
  Me.chkVerseLines.Value = Abs(CLng(CBool(GetSetting(App.Title, "Settings", "VerseLines", "0"))))
  Call chkVerseLines_Click
  Me.chkNewPage.Value = Abs(CLng(CBool(GetSetting(App.Title, "Settings", "NewPage", "0"))))
  Me.chkAddNoteSpace.Value = Abs(CLng(CBool(GetSetting(App.Title, "Settings", "AddNoteSpace", "0"))))
  Me.chkCenterBkHeading.Value = Abs(CLng(CBool(GetSetting(App.Title, "Settings", "CenterBkHeading", "0"))))
  Me.chkCenterChapHeading.Value = Abs(CLng(CBool(GetSetting(App.Title, "Settings", "CenterChapHeading", "0"))))
  If Idx = 1 Then
    S = "MyBible.txt"
  Else
    S = "MyBible.rtf"
  End If
  Me.lblSavePath.Caption = GetSetting(App.Title, "Settings", "SaveBible", AddSlash(App.Path) & S)
  Call OptRtf_Click(Idx)
  Me.chkAddPNotes.Value = Abs(CLng(CBool(GetSetting(App.Title, "Settings", "AddPNotes", "0"))))
  Me.chkPNotesAbove.Value = Abs(CLng(CBool(GetSetting(App.Title, "Settings", "PNotesAbove", "0"))))
  Call chkAddPNotes_Click
  Me.chkIncludeTheoNotes.Value = Abs(CBool(GetSetting(App.Title, "Settings", "IncludeTheoNotes", "0")))
  Me.chkTextColor.ForeColor = CLng(GetSetting(App.Title, "Settings", "PNotesColor", CStr(&HFF0000)))
  
  Me.OptRtf(0).BackColor = cMedium
  Me.OptRtf(1).BackColor = cMedium
  Me.chkVerseLines.BackColor = cMedium
  Me.chkNewPage.BackColor = cMedium
  Me.chkCenterBkHeading.BackColor = cMedium
  Me.chkCenterChapHeading.BackColor = cMedium
  Me.chkAddPNotes.BackColor = cMedium
  Me.chkPNotesAbove.BackColor = cMedium
  Me.chkIncludeTheoNotes.BackColor = cMedium
  
  Me.lblSavePath.BackColor = cVLight
  IsRTF = Me.OptRtf(0).Value
  bCancel = True
  
  S = Me.lblBumpDefault.Caption
  Idx = InStr(1, S, "(")
  S = Left$(S, Idx) & CStr(FntSize) & Mid$(S, Idx + 1)
  Me.lblBumpDefault.Caption = S
  With Me.cboBumpFactor
    .AddItem "0"
    .AddItem "2"
    .AddItem "4"
    .AddItem "6"
    .AddItem "8"
    .ListIndex = CLng(GetSetting(App.Title, "Settings", "BumpFactor", "1"))
  End With
End Sub

Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
End Sub

Private Sub Label1_Click(Index As Integer)
  Select Case Index
    Case lRtf
      Me.OptRtf(0).Value = True
    Case lTxt
      Me.OptRtf(1).Value = True
    Case lNewP
      If Me.chkNewPage.Value = vbChecked Then
        Me.chkNewPage.Value = vbUnchecked
      Else
        Me.chkNewPage.Value = vbChecked
      End If
    Case lCtrB
      If Me.chkCenterBkHeading.Value = vbChecked Then
        Me.chkCenterBkHeading.Value = vbUnchecked
      Else
        Me.chkCenterBkHeading.Value = vbChecked
      End If
    Case lCtrC
      If Me.chkCenterChapHeading.Value = vbChecked Then
        Me.chkCenterChapHeading.Value = vbUnchecked
      Else
        Me.chkCenterChapHeading.Value = vbChecked
      End If
    Case lNewL
      If Me.chkVerseLines.Value = vbChecked Then
        Me.chkVerseLines.Value = vbUnchecked
      Else
        Me.chkVerseLines.Value = vbChecked
      End If
    Case lBlnk
      If Me.chkAddNoteSpace.Value = vbChecked Then
        Me.chkAddNoteSpace.Value = vbUnchecked
      Else
        Me.chkAddNoteSpace.Value = vbChecked
      End If
  End Select
End Sub

Private Sub lblAddPNotes_Click()
  With Me.chkAddPNotes
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = Checked
    End If
  End With
End Sub

Private Sub lblIncludeTheoNotes_Click()
  With Me.chkIncludeTheoNotes
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = vbChecked
    End If
  End With
End Sub

Private Sub lblPNotesAbove_Click()
  With Me.chkPNotesAbove
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = Checked
    End If
  End With

End Sub

Private Sub lblSavePath_Click()
  Me.cmdBrowse.Value = True
End Sub

Private Sub OptRtf_Click(Index As Integer)
  Dim S As String
  
  Me.chkCenterBkHeading.Enabled = Me.OptRtf(0).Value
  Me.chkCenterChapHeading.Enabled = Me.chkCenterBkHeading.Enabled
  Me.Label1(5).Enabled = Me.chkCenterBkHeading.Enabled
  Me.Label1(6).Enabled = Me.chkCenterBkHeading.Enabled
  IsRTF = Index = 0
  Me.chkTextColor.Enabled = IsRTF = True And Me.chkAddPNotes.Value = vbChecked
  Me.cboBumpFactor.Enabled = IsRTF
  
  S = Me.lblSavePath.Caption
  If Len(S) <> 0 Then
    If Mid$(S, Len(S) - 3, 1) = "." Then
      If Index = 1 Then
        If StrComp(Right$(S, 3), "txt") <> 0 Then Me.lblSavePath.Caption = Left$(S, Len(S) - 3) & "txt"
      Else
        If StrComp(Right$(S, 3), "rtf") <> 0 Then Me.lblSavePath.Caption = Left$(S, Len(S) - 3) & "rtf"
      End If
    End If
  End If
End Sub
