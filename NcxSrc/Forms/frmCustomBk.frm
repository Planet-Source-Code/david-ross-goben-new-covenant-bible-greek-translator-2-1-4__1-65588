VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCustomBk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Custom Background"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
   Icon            =   "frmCustomBk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGuess 
      Caption         =   "Try to &GUESS Colors"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "This command will try to determine 4 good theme colors for the selected image"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      ToolTipText     =   "Cancel Changes"
      Top             =   3000
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Accept changes and apply them"
      Top             =   3000
      Width           =   915
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "With this checked, selecting a color button allows you to pick a color from the image"
      Top             =   2760
      Width           =   195
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse for picture..."
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Brouse for a background picture image"
      Top             =   180
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Theme Color Selection"
      Height          =   2415
      Left            =   180
      TabIndex        =   12
      Top             =   240
      Width           =   2115
      Begin VB.CommandButton cmdColor 
         Height          =   435
         Index           =   3
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1860
         Width           =   435
      End
      Begin VB.CommandButton cmdColor 
         Height          =   435
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   435
      End
      Begin VB.CommandButton cmdColor 
         Height          =   435
         Index           =   1
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   780
         Width           =   435
      End
      Begin VB.CommandButton cmdColor 
         Height          =   435
         Index           =   0
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblColors 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Very Light Color:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Top             =   1980
         Width           =   1155
      End
      Begin VB.Label lblColors 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Light Color:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label lblColors 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medium Color:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label lblColors 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Color:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2235
      Left            =   2460
      ScaleHeight     =   2175
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   480
      Width           =   2415
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1755
         Left            =   0
         ScaleHeight     =   1755
         ScaleWidth      =   1995
         TabIndex        =   17
         Top             =   0
         Width           =   1995
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTip 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Now Select a color button, above, and click on the  displayed image."
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblPick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pick colors from picture"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      ToolTipText     =   "With this checked, selecting a color button allows you to pick a color from the image"
      Top             =   2760
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image:"
      Height          =   195
      Left            =   2460
      TabIndex        =   10
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmCustomBk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' using this API is much faster than using the Picturebox.Point() function
'
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'
' local storage of temporary defaults
'
Private FName As String
Private clrDark As Long
Private clrMedium As Long
Private clrLight As Long
Private ClrVLight As Long
Private ColorPick As Boolean
Private BtnIdx As Long

Private Sub Form_Load()
  
  clrDark = custDark                  'initialize temp color storage to defaults
  clrMedium = custMedium
  clrLight = custLight
  ClrVLight = CustVLight
  Me.Frame1.BackColor = cLight
'
' get default custom image pathname
'
  FName = GetSetting(App.Title, "Settings", "CustPicture", vbNullString)
  If Len(FName) <> 0 Then
    If Not Fso.FileExists(FName) Then FName = vbNullString
  End If
  Call SetColors
  bCancel = True
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Form repaint. If a custom image set, use it
'*******************************************************************************
Private Sub Form_Paint()
  If Len(FName) <> 0 Then
    PaintTileFormBackground Me, Me.Picture2   'use custom image
  Else
    PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Check1_Click
' Purpose           : Toggle user being able to pic colors by clicking image
'*******************************************************************************
Private Sub Check1_Click()
  If Me.Check1.Value = vbChecked Then
    Me.lblTip.Visible = True
    Me.cmdGuess.Visible = False
  Else
    Me.Picture2.MousePointer = vbDefault
    Me.lblTip.Visible = False
    Me.cmdGuess.Visible = True
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBrowse_Click
' Purpose           : Browse for an image file
'*******************************************************************************
Private Sub cmdBrowse_Click()
  With Me.CommonDialog1
    .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNLongNames
    .DialogTitle = "Browse For Image"
    .Filter = "Image Files|*.bmp;*.cur;*.gif;*.ico;*.jpg;*.jpeg;*.wmf;*.emf|All Files (*.*)|*.*"
    .FileName = FName
    .CancelError = True
    On Error Resume Next
    .ShowOpen
    FName = Trim$(.FileName)
    If Err.Number <> 0 Then FName = vbNullString
    If Len(FName) = 0 Then Exit Sub
    Err.Clear
    Me.Picture2.Picture = LoadPicture(FName)
    Me.Picture2.Picture = Me.Picture2.Image
    If Err.Number <> 0 Then
      MessageBox Me, "Invalid image format: " & FName, vbOKOnly Or vbExclamation, "Imvalid Format"
      FName = vbNullString
      Exit Sub
    End If
    PaintTileFormBackground Me, Me.Picture2
    PaintTilePicBackground Me.Picture1, Me.Picture2
    Me.lblPick.Refresh
    On Error GoTo 0
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Cancel without saving anything
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdColor_Click
' Purpose           : User clicked a button. Either bring up the color dialog
'                   : if the checkbox is unchecked, or else allow the user
'                   : to pick a color directly from the image
'*******************************************************************************
Private Sub cmdColor_Click(Index As Integer)
  Dim Clr As Long
  Dim S As String
  
  BtnIdx = Index
  If Me.Check1.Value = vbChecked Then
    Me.Picture2.MousePointer = 2
    Me.Picture1.MousePointer = 2
    ColorPick = True
  Else
    Me.Picture2.MousePointer = vbDefault
    Me.Picture1.MousePointer = vbDefault
    With Me.CommonDialog1
      .Flags = cdlCCFullOpen Or cdlCCRGBInit
      .Color = Me.cmdColor(Index).BackColor
      .CancelError = True
      On Error Resume Next
      .ShowColor
      If Err.Number <> 0 Then Exit Sub
      On Error GoTo 0
      Clr = .Color
    End With
    Select Case BtnIdx
      Case 0
        clrDark = Clr
      Case 1
        clrMedium = Clr
      Case 2
        clrLight = Clr
      Case 3
        ClrVLight = Clr
    End Select
  End If
  SetColors
End Sub

'*******************************************************************************
' Subroutine Name   : lblPick_Click
' Purpose           : Reflect the checkbox with the label associated with it
'*******************************************************************************
Private Sub lblPick_Click()
  With Me.Check1
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = vbChecked
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdGuess_Click
' Purpose           : Try to GUESS at 4 good theme colors. Quick and dirty
'*******************************************************************************
Private Sub cmdGuess_Click()
  Dim Clrz(255) As Long
  Dim X As Long, Y As Long, Clr As Long, Idx As Long
  Dim Qtr As Long, Base As Long, BClr As Long, DC As Long
  Dim Wd As Long, Ht As Long
  Dim MyRgb As RGB_Type
  
  Me.Enabled = False                                'we are busy
  Screen.MousePointer = vbHourglass
  DoEvents
  
  With Me.Picture2
    DC = .hdc                                       'get oft-used DC
    Wd = .Width                                     'initial width and height limits
    Ht = .Height
  End With
  With Me.Picture1
    If Wd > .ScaleWidth Then Wd = .ScaleWidth       'adjust width and height as needed
    If Ht > .ScaleHeight Then Ht = .ScaleHeight
    
    For X = Me.Picture1.ScaleLeft \ 15 To Wd \ 15 - 1  'scan by pixels
      For Y = Me.ScaleTop \ 15 To Ht \ 15 - 1
        Clr = GetPixel(DC, X, Y)                    'get a color
        MyRgb = ToRGB(Clr)                          'convert to RGB
        Clrz((MyRgb.rgbRed + MyRgb.rgbGreen + MyRgb.rgbBlue) \ 3) = Clr 'average
      Next Y
    Next X
  End With
  
  BClr = &H808080                                   'hacks. Do not get too dark
  Base = 128 + 64
  Qtr = 256 / 16
  X = Base + Qtr * 3 + Qtr / 2 - 1
  Y = Base + Qtr * 4 - 1
  
  For Clr = X To Y
    If Clrz(Clr) <> 0 Then Exit For
  Next Clr
  If Clr = Y + 1 Then
    Y = X
    X = Base + Qtr * 3 - 1
    For Clr = X To Y
      If Clrz(Clr) <> 0 Then Exit For
    Next Clr
    If Clr = Y + 1 Then
      ClrVLight = &HF0F0F0
    Else
      ClrVLight = Clrz(Clr)
    End If
  Else
    ClrVLight = Clrz(Clr)
  End If
  
  X = Base + Qtr * 2 + Qtr / 2 - 1
  Y = Base + Qtr * 3 - 1
  For Clr = X To Y
    If Clrz(Clr) <> 0 Then Exit For
  Next Clr
  If Clr = Y + 1 Then
    Y = X
    X = Base + Qtr * 2 - 1
    For Clr = X To Y
      If Clrz(Clr) <> 0 Then Exit For
    Next Clr
    If Clr = Y + 1 Then
      clrLight = &HE0E0E0
    Else
      clrLight = Clrz(Clr)
    End If
  Else
    clrLight = Clrz(Clr)
  End If
  
  X = Base + Qtr + Qtr / 2 - 1
  Y = Base + Qtr * 2 - 1
  For Clr = X To Y
    If Clrz(Clr) <> 0 Then Exit For
  Next Clr
  If Clr = Y + 1 Then
    Y = X
    X = Base + Qtr - 1
    For Clr = X To Y
      If Clrz(Clr) <> 0 Then Exit For
    Next Clr
    If Clr = Y + 1 Then
      clrMedium = &HD0D0D0
    Else
      clrMedium = Clrz(Clr)
    End If
  Else
    clrMedium = Clrz(Clr)
  End If
  
  X = Base + Qtr / 2 - 1
  Y = Base + Qtr - 1
  For Clr = X To Y
    If Clrz(Clr) <> 0 Then Exit For
  Next Clr
  If Clr = Y + 1 Then
    Y = X
    X = Base
    For Clr = X To Y
      If Clrz(Clr) <> 0 Then Exit For
    Next Clr
    If Clr = Y + 1 Then
      clrDark = &HC0C0C0
    Else
      clrDark = Clrz(Clr)
    End If
  Else
    clrDark = Clrz(Clr)
  End If
  SetColors
  Me.Enabled = True
  Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : Accept selections. Set thme to the system
'*******************************************************************************
Private Sub cmdOK_Click()
  custDark = clrDark
  custMedium = clrMedium
  custLight = clrLight
  CustVLight = ClrVLight
  SaveSetting App.Title, "Settings", "custDark", CStr(custDark)
  SaveSetting App.Title, "Settings", "custMedium", CStr(custMedium)
  SaveSetting App.Title, "Settings", "custlight", CStr(custLight)
  SaveSetting App.Title, "Settings", "CustVLight", CStr(CustVLight)
  SaveSetting App.Title, "Settings", "CustPicture", FName
  frmGrkXlate.picTile(bkCustom).Picture = Me.Picture2.Picture
  bCancel = False
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : SetColors
' Purpose           : Update dialog box with current changed  settings
'*******************************************************************************
Private Sub SetColors()
  Me.cmdColor(0).BackColor = clrDark
  Me.cmdColor(0).ToolTipText = ClrTip(clrDark)
  Me.cmdColor(1).BackColor = clrMedium
  Me.cmdColor(1).ToolTipText = ClrTip(clrMedium)
  Me.cmdColor(2).BackColor = clrLight
  Me.cmdColor(2).ToolTipText = ClrTip(clrLight)
  Me.cmdColor(3).BackColor = ClrVLight
  Me.cmdColor(3).ToolTipText = ClrTip(ClrVLight)
  
  Me.Frame1.BackColor = clrLight
  Me.Picture2.Cls
  If Len(FName) <> 0 Then
    On Error Resume Next
    Me.Picture2.Picture = LoadPicture(FName)
    If Err.Number = 0 Then Exit Sub
  End If
  On Error Resume Next
  Me.Picture2.Picture = frmGrkXlate.picTile(bkCustom).Picture
End Sub

'*******************************************************************************
' Function Name     : ClrTip
' Purpose           : Set the tooltip for each button with its background color value
'*******************************************************************************
Private Function ClrTip(Clr As Long) As String
  Dim S As String
  
  S = Hex$(Clr)
  S = Left$("000000", 6 - Len(S)) & S
  ClrTip = "&h" & S & "; Red= &h" & Right$(S, 2) & ", Green= &h" & Mid$(S, 3, 2) & ", Blue= &h" & Left$(S, 2)
End Function

'*******************************************************************************
' Subroutine Name   : Picture1_Paint
' Purpose           : Refresh container picture background as needed
'*******************************************************************************
Private Sub Picture1_Paint()
  If Len(FName) <> 0 Then
    PaintTilePicBackground Me.Picture1, Me.Picture2
  Else
    PaintTilePicBackground Me.Picture1, frmGrkXlate.picTile(Background)   'repaint background
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Picture2_MouseDown
' Purpose           : When the user can select colors from the image, check the
'                   : pixel being clicked
'*******************************************************************************
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Clr As Long
  If ColorPick Then
    Clr = Me.Picture2.Point(X, Y)
    Select Case BtnIdx
      Case 0
        clrDark = Clr
      Case 1
        clrMedium = Clr
      Case 2
        clrLight = Clr
      Case 3
        ClrVLight = Clr
    End Select
    Me.Picture1.MousePointer = vbDefault
    Me.Picture2.MousePointer = vbDefault
    ColorPick = False
  End If
  Call SetColors
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Clr As Long
  If ColorPick Then
    Clr = Me.Picture1.Point(X, Y)
    Select Case BtnIdx
      Case 0
        clrDark = Clr
      Case 1
        clrMedium = Clr
      Case 2
        clrLight = Clr
      Case 3
        ClrVLight = Clr
    End Select
    Me.Picture1.MousePointer = vbDefault
    Me.Picture2.MousePointer = vbDefault
    ColorPick = False
  End If
  Call SetColors
End Sub
