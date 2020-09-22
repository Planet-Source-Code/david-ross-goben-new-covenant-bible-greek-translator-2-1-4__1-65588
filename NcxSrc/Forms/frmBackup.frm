VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Modified Files to Secondary Location"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Restore backup files to main database location..."
      Height          =   375
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "This command is useful to recover corrupted files, and to restore to the last backup image"
      Top             =   1620
      Width           =   3735
   End
   Begin VB.ComboBox cboAutoSave 
      Height          =   315
      ItemData        =   "frmBackup.frx":27A2
      Left            =   2340
      List            =   "frmBackup.frx":27A4
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Updates only occur if there are changes present"
      Top             =   1200
      Width           =   795
   End
   Begin VB.CheckBox chkAutoMin 
      Height          =   195
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Updates only occur if there are changes present"
      Top             =   1260
      Width           =   195
   End
   Begin VB.CheckBox chkAutosave 
      Height          =   195
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Save an images of updatable files prior to beginning your work (handy as a session UNDO)"
      Top             =   960
      Width           =   195
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   1620
      Width           =   855
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
      Left            =   4560
      TabIndex        =   7
      ToolTipText     =   "Save backup files"
      Top             =   1620
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      ToolTipText     =   "Browse for a storage path"
      Top             =   480
      Width           =   435
   End
   Begin VB.Label lblActiveDB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Active database storage location:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   195
      Left            =   2760
      TabIndex        =   14
      Top             =   2100
      Width           =   2880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Active database storage location:"
      Height          =   195
      Left            =   300
      TabIndex        =   13
      Top             =   2100
      Width           =   2370
   End
   Begin VB.Label lblTimerBackup1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "minutes to the main, activate database."
      Height          =   195
      Index           =   1
      Left            =   3180
      TabIndex        =   11
      ToolTipText     =   "Updates only occur if there are changes present"
      Top             =   1260
      Width           =   2775
   End
   Begin VB.Label lblTimerBackup1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto-save changes every"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   10
      ToolTipText     =   "Updates only occur if there are changes present"
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label lblAuthosave 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto-save a backup to the above backup path on program startup."
      Height          =   195
      Left            =   480
      TabIndex        =   9
      ToolTipText     =   "Save an image of the updatable files prior to beginning your work (handy as a session UNDO)"
      Top             =   960
      Width           =   4740
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "lblWidth"
      Height          =   195
      Left            =   5940
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblSavePath 
      BackColor       =   &H80000016&
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
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Double-click to explore..."
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Backup Location:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1650
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Loading As Boolean        'flag indicating form is loading

'*******************************************************************************
' Subroutine Name   : cboAutoSave_Click
' Purpose           : Choose minute interval for auto-save
'*******************************************************************************
Private Sub cboAutoSave_Click()
  If Loading Then Exit Sub
  Me.chkAutoMin.Enabled = CBool(Me.cboAutoSave.ListIndex)
  Me.lblTimerBackup1(0).Enabled = Me.chkAutoMin.Enabled
  Me.lblTimerBackup1(1).Enabled = Me.chkAutoMin.Enabled
  SaveSetting App.Title, "Settings", "AutoTimer", CStr(Me.chkAutoMin.Value)
  SaveSetting App.Title, "Settings", "TimeSet", Me.cboAutoSave.Text
  If Me.chkAutoMin.Enabled Then
    AutoTime = CLng(Me.cboAutoSave.Text)  'get minutes interval count
    AutoTimeUpd = AutoTime
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : chkAutoMin_Click
' Purpose           : Auto-save option
'*******************************************************************************
Private Sub chkAutoMin_Click()
  If Loading Then Exit Sub
  SaveSetting App.Title, "Settings", "AutoTimer", CStr(Me.chkAutoMin.Value)
  SaveSetting App.Title, "Settings", "TimeSet", Me.cboAutoSave.Text
  If Me.chkAutoMin.Enabled Then
    AutoTime = CLng(Me.cboAutoSave.Text)
    AutoTimeUpd = AutoTime
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : chkAutosave_Click
' Purpose           : Auto-save to backup directory on startup
'*******************************************************************************
Private Sub chkAutosave_Click()
  Me.cmdCopy.Enabled = Not CBool(Me.chkAutosave.Value)
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBrowse_Click
' Purpose           : Browse for backup folder
'*******************************************************************************
Private Sub cmdBrowse_Click()
  Dim S As String

  S = DirBrowser(Me.hwnd, ViewDirsOnly, "Select Data Backup Path", BackupPath)
  If Len(S) Then
    If StrComp(S, AddSlash(AppPath) & "DB", vbTextCompare) = 0 Then
      MessageBox Me, "Cannot save backups to the database storage location.", vbOKOnly Or vbExclamation, "Path Invalid"
      Exit Sub
    End If
    If Right$(AddSlash(S), 2) = ":\" Then
      MessageBox Me, "Cannot save backups to the Drive Root Folder.", vbOKOnly Or vbExclamation, "Path Invalid"
    End If
    BackupPath = S
    
    Me.lblSavePath.Caption = BackupPath
    Me.lblWidth.Caption = BackupPath
    If Me.lblWidth.Width > Me.lblSavePath.Width Then
      Me.lblSavePath.ToolTipText = BackupPath
    End If
    SaveSetting App.Title, "Settings", "BackupPath", BackupPath
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Abort
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopy_Click
' Purpose           : Copy the backed up images to the main application
'*******************************************************************************
Private Sub cmdCopy_Click()
  Dim Ary() As String
  Dim tso As TextStream

  If MessageBox(Me, "You are about to over-write you main database with" & vbCrLf & _
                    "backup files. Are you sure you want to do this?", _
                    vbYesNo Or vbQuestion Or vbDefaultButton2, "Verify Over-write of Main Database") = vbNo Then Exit Sub
  AutoDirty = False                           'reset automatic backup
  Me.cmdCopy.Enabled = False                  'turn botton off (would be redundant)
  Me.Enabled = False
  Screen.MousePointer = vbHourglass           'show that we are busy
  DoEvents
'
' restore GreekWordRef.txt
'
  If Fso.FileExists(AddSlash(BackupPath) & "GreekWordRef.txt") Then
    Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "GreekWordRef.txt", ForReading, False)
    WordRef = Split(ts.ReadAll, vbCrLf)
    ts.Close
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\GreekWordRef.txt", ForWriting, False)
    ts.Write Join(WordRef, vbCrLf)
    ts.Close
  End If
'
' restore WordMap.txt
'
  If Fso.FileExists(AddSlash(BackupPath) & "WordMap.txt") Then
    Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "WordMap.txt", ForReading, False)
    WordMap = Split(ts.ReadAll, vbCrLf)
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\WordMap.txt", ForWriting, True)
    ts.Write Join(WordMap, vbCrLf)
    ts.Close
  End If
'
' restore GreekBBL.txt
'
  If Fso.FileExists(AddSlash(BackupPath) & "GreekBBL.txt") Then
    Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "GreekBBL.txt", ForReading, False)
    GrkBBL = Split(ts.ReadAll, vbCrLf)
    ts.Close
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\GreekBBL.txt", ForWriting, False)
    ts.Write Join(GrkBBL, vbCrLf)
    ts.Close
    BBLDirty = False 'no longer dirty (if it was)
  End If
'
' Restore MPV.txt, if it exists
'
  If Fso.FileExists(AddSlash(BackupPath) & "MPV.txt") Then
    Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "MPV.txt", ForReading, False)
    Set tso = Fso.OpenTextFile(AddSlash(AppPath) & "DB\MPV.txt", ForWriting, True)
    If BblVersion = UserPVer Then
      Bible = Split(ts.ReadAll, vbCrLf)
      tso.Write Join(Bible, vbCrLf)
    Else
      Ary = Split(ts.ReadAll, vbCrLf)
      tso.Write Join(Ary, vbCrLf)
    End If
    tso.Close
    PVDirty = False 'no longer dirty (if it was)
  End If
  
  Call frmGrkXlate.UpdateVerse    'reset verse, in case this changes it
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  
  MessageBox Me, "Backup image loaded and set to main image.", vbOKOnly Or vbInformation, "Backup Restore Complete"
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSave_Click
' Purpose           : Save backup now
'*******************************************************************************
Private Sub cmdSave_Click()
  Dim Ary() As String
  Dim Idx As Long, I As Long
  Dim Errors As Boolean
'
' update word references
'
  If Not Fso.FolderExists(BackupPath) Then
    On Error Resume Next
    Fso.CreateFolder BackupPath
    If Err.Number <> 0 Then
      MessageBox Me, "Cannot create backup folder path: " & BackupPath, vbOKOnly Or vbExclamation, "Cannot Backup to Location"
      Exit Sub
    End If
    On Error GoTo 0
  End If
  Me.Enabled = False
  Screen.MousePointer = vbHourglass
  DoEvents
  For Idx = 1 To UBound(BBLWIdx)
    If Len(WordRef(Idx)) <> 0 Then
      Ary = Split(WordRef(Idx), vbTab)
      Ary(1) = CStr(BBLWIdx(Idx))
      WordRef(Idx) = Join(Ary, vbTab)
    End If
  Next Idx
  On Error Resume Next
'
' ALWAYS save word reference table (constantly updated)
'
  Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "GreekWordRef.txt", ForWriting, True)
  Errors = CBool(Err.Number)
  ts.Write Join(WordRef, vbCrLf)
  ts.Close
'
' and word map (user-selections for syninyms of words)
'
  Err.Clear
  Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "WordMap.txt", ForWriting, True)
  Errors = Errors Or CBool(Err.Number)
  ts.Write Join(WordMap, vbCrLf)
  ts.Close
'
' and Greek word reference
'
  Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "GreekBBL.txt", ForWriting, True)
  Errors = CBool(Err.Number)
  ts.Write Join(GrkBBL, vbCrLf)
  ts.Close
'
' save user personal bible
'
  If PersonalVersion Then
    If BblVersion = UserPVer Then
      Err.Clear
      Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "MPV.txt", ForWriting, True)
      Errors = Errors Or CBool(Err.Number)
      ts.Write Join(Bible, vbCrLf)
      ts.Close
    Else
      Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\MPV.txt", ForReading, False)
      Ary = Split(ts.ReadAll, vbCrLf)
      ts.Close
      Err.Clear
      Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "MPV.txt", ForWriting, True)
      Errors = Errors Or CBool(Err.Number)
      ts.Write Join(Ary, vbCrLf)
      ts.Close
    End If
  End If
'
' save personal verse notes
'
  Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "DB\MyNotes.txt", ForWriting, True)
  ts.Write Join(MyNotes, vbCrLf)
  ts.Close
'
' save viewing history
'
  Set ts = Fso.OpenTextFile(AddSlash(BackupPath) & "DB\History.txt", ForWriting, True)
  With colHistory
    I = .Count - 1000
    If I < 1 Then I = 1
    For Idx = I To .Count
      ts.WriteLine .Item(Idx)
    Next Idx
    ts.Close
  End With
'
' now close up
'
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  
  If Errors Then
    MessageBox Me, "Errors encountered while writing backups. Investigate the backup path:" & vbCrLf & _
    BackupPath, vbOKOnly Or vbExclamation, "Errors Encountered"
  Else
    MessageBox Me, "Files backed up successfully", vbOKOnly Or vbInformation, "Backup Succeeded"
  End If
  
  SaveSetting App.Title, "Settings", "Autosave", CStr(Me.chkAutosave.Value)
  SaveSetting App.Title, "Settings", "AutoTimer", CStr(Me.chkAutoMin.Value)
  SaveSetting App.Title, "Settings", "TimeSet", Me.cboAutoSave.Text
  If Me.chkAutoMin.Enabled Then
    AutoTime = CLng(Me.cboAutoSave.Text)
    AutoTimeUpd = AutoTime
  End If
  AutoDirty = False   'we have saved the data, so
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Form startup
'*******************************************************************************
Private Sub Form_Load()
  Dim S As String
  
  Loading = True
  S = GetSetting(App.Title, "Settings", "BackupPath", vbNullString)
  Me.chkAutosave.Value = CLng(GetSetting(App.Title, "Settings", "AutoSave", "0"))
  Me.chkAutoMin.Value = CLng(GetSetting(App.Title, "Settings", "AutoTimer", "0"))
  With Me.cboAutoSave
    .AddItem 0
    .AddItem 1
    .AddItem 5
    .AddItem 15
    .AddItem 30
    Select Case CLng(GetSetting(App.Title, "Settings", "TimeSet", "0"))
      Case 1
        .ListIndex = 1
      Case 5
        .ListIndex = 2
      Case 15
        .ListIndex = 3
      Case 30
        .ListIndex = 4
      Case Else
        .ListIndex = 0
    End Select
    Me.chkAutoMin.Enabled = CBool(.ListIndex)
    Me.lblTimerBackup1(0).Enabled = Me.chkAutoMin.Enabled
    Me.lblTimerBackup1(1).Enabled = Me.chkAutoMin.Enabled
  End With
  BackupPath = AddSlash(AppPath) & "DB\Backup"
  S = GetSetting(App.Title, "Settings", "BackupPath", BackupPath)
  If Fso.FolderExists(S) Then
    BackupPath = S
  Else
    Me.cmdCopy.Enabled = False
  End If
  Me.lblSavePath.Caption = BackupPath
  Me.lblWidth.Caption = BackupPath
  If Me.lblWidth.Width > Me.lblSavePath.Width Then
    Me.lblSavePath.ToolTipText = BackupPath
  End If
  Me.lblSavePath.BackColor = cVLight
  Me.chkAutosave.BackColor = cMedium
  Me.chkAutoMin.BackColor = cMedium
  Loading = False
  Me.lblActiveDB.Caption = AddSlash(AppPath) & "DB"
  Me.lblActiveDB.ToolTipText = Me.lblActiveDB.Caption & ". Double-click to explore..."
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Paint texture on background
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmGrkXlate.tmrAutoBackup.Enabled = (Me.chkAutoMin.Enabled And Me.chkAutoMin.Value = vbChecked)
End Sub

Private Sub lblActiveDB_DblClick()
  BrowsePath Me.hwnd, Me.lblActiveDB.Caption
End Sub

Private Sub lblAuthosave_Click()
  If Me.chkAutosave.Value = vbChecked Then
    Me.chkAutosave.Value = vbUnchecked
  Else
    Me.chkAutosave.Value = vbChecked
  End If
End Sub

Private Sub lblSavePath_DblClick()
  BrowsePath Me.hwnd, Me.lblSavePath.Caption
End Sub

Private Sub lblTimerBackup1_Click(Index As Integer)
  If Me.chkAutoMin.Value = vbChecked Then
    Me.chkAutoMin.Value = vbUnchecked
  Else
    Me.chkAutoMin.Value = vbChecked
  End If
End Sub
