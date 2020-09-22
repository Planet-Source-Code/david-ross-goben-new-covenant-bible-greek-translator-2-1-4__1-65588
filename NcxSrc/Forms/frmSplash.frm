VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9930
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFader 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9480
      Tag             =   "Fade in/out timer"
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   9480
      Top             =   480
   End
   Begin VB.Label lblInfo2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSplash.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   660
      TabIndex        =   8
      Top             =   3900
      Width           =   8655
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Covenant Bible Greek Translator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A0A0A0&
      Height          =   510
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   420
      Width           =   7800
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   300
      Picture         =   "frmSplash.frx":0113
      Top             =   780
      Width           =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   435
      X2              =   9375
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label lblJohn11Greek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "en arch hn o logoV kai o logoV hn proV ton qeon kai qeoV hn o logoV"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   660
      TabIndex        =   6
      Top             =   5700
      Width           =   8520
   End
   Begin VB.Label lblJohn11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "John 1:1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   5280
      Width           =   885
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   1560
      X2              =   9315
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8430
      TabIndex        =   4
      Top             =   120
      Width           =   885
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSplash.frx":1CFD
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   660
      TabIndex        =   3
      Top             =   2220
      Width           =   8655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004-2006 by David Ross Goben. All rights reserved."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1545
      TabIndex        =   2
      Top             =   1740
      Width           =   7815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NA26/27 Greek Text is used, which is the most original text available."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1575
      TabIndex        =   1
      Top             =   1320
      Width           =   7755
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Covenant Bible Greek Translator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   540
      Width           =   7800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   435
      X2              =   9375
      Y1              =   5640
      Y2              =   5640
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'*******************************************************************************
' Variables used for Fading
'*******************************************************************************
Private m_Fader As Byte                     'fader value on systems that support it
Private m_FaderDown As Boolean              'true when fading out
Private m_FaderInc As Byte                  'amount to change fade increments
Private m_OsType As Integer                 'keep track of operating system
Public AllowFormFading As Boolean           'Set to TRUE to allow fading at all
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Set up visage, and then fade in
'*******************************************************************************
Private Sub Form_Load()
  Me.lblInfo.ForeColor = vbBlue
  Me.lblInfo2.ForeColor = vbBlue
  Me.lblCaption(0).ForeColor = vbBlue
  Me.lblCaption(1).Left = Me.lblCaption(0).Left + 45
  Me.lblCaption(1).Top = Me.lblCaption(0).Top + 45
  Me.lblCaption(0).ZOrder 0
  Me.lblVersion.Caption = "Version " & GetAppVersion()
  
  If SplashIsOn Then
    Me.lblInfo.Visible = False
    Me.lblInfo2.Visible = False
    Me.lblJohn11.Top = Me.lblInfo.Top
    Me.Line1.Y1 = Me.lblJohn11.Top + Me.lblJohn11.Height + 60
    Me.Line1.Y2 = Me.Line1.Y1
    Me.Line2.Y1 = Me.Line1.Y1
    Me.Line2.Y2 = Me.Line1.Y1
    Me.lblJohn11Greek.Top = Me.Line1.Y1 + 60
  End If
  Me.Height = Me.lblJohn11Greek.Top + Me.lblJohn11Greek.Height + Me.lblVersion.Top + 60
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  Call FaderInit(15)                        'fade the form in
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Set background tiling
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Timer1.Enabled = False                 'ensure splash timer disabled
  SplashIsOn = False                        'turn off splash flag, if enabled
  On Error Resume Next
  frmGrkXlate.mnuHlpAbout.Enabled = True    'esnure menu item enabled
End Sub

Private Sub Image1_Click()
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Private Sub Label2_Click(Index As Integer)
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Private Sub lblCaption_Click(Index As Integer)
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Private Sub lblInfo_Click()
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Private Sub lblInfo2_Click()
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Private Sub lblVersion_Click()
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Private Sub lblVote_Click()
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

'*******************************************************************************
' Subroutine Name   : Timer1_Timer
' Purpose           : Unload the form when the timer times out
'*******************************************************************************
Private Sub Timer1_Timer()
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Private Sub Form_Click()
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Call FaderUnload  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'*******************************************************************************
' Subroutine Name   : FaderInit
' Purpose           : Invoked from within Form Load after all setups are done
'*******************************************************************************
Private Sub FaderInit(Optional FaderIncrement As Byte = 26)
  'the following coditional compilation variable is defined in the Project Properies
  'Make tab under Conditional Compilation Objects as: AllowFading = 1
  #If Allowfading Then                          'if conditional constant defined...
    AllowFormFading = True
  #End If
  If AllowFormFading Then                       'If fading is allowed for this form...
    m_OsType = GetOSType                        'get operating system flag
    Select Case m_OsType
      Case 6, Is > 7                            'w2k or XP?
        m_FaderInc = FaderIncrement             'set fading icrement
        m_Fader = m_FaderInc                    'init barely visible
        m_FaderDown = False                     'we are fading in
        SetWindowTranslucency Me.hwnd, m_Fader  'set translucency
        Me.Enabled = False                      'disable hanky-panky on form
        Me.tmrFader.Enabled = True              'turn on fader
        Exit Sub                                'done with w2k/xp...
    End Select
  End If
'
' the following executes on non-w2k/xp or AllowFormFading=False
'
  Me.Show                                       'ensure fully visible
  DoEvents
  AllowFormFading = False                       'force flag off if not w2k/xp
'
' if splash is on, auto-remove splash screen after timeout
'
  If SplashIsOn Then
    Me.Timer1.Enabled = True
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : FaderUnload
' Purpose           : Invoked where you would normally do an Unload Me
'
' Add the following to any Form_QueryUnload event:
'  If UnloadMode = vbFormControlMenu Then
'    Call FaderUnload
'    If AllowFormFading Then Cancel = 1 'the fader will actually unload the form
'  End If
'*******************************************************************************
Private Sub FaderUnload()
  If AllowFormFading Then           'If fading is allowed for this form
    Select Case m_OsType
      Case 6, Is > 7                'W2K and XP+
        m_Fader = 255               'set up for fading down
        m_FaderDown = True          'initiate fading down
        Me.Enabled = False          'disable form for now
        Me.tmrFader.Enabled = True  'start timer
        Exit Sub
    End Select
  End If
'
' the following executes on non-w2k/xp or AllowFormFading=False
'
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : tmrFader_Timer
' Purpose           : Allow fade-in and fade-outs
'                     Initialize tmrFader with Enabled=False, and Interval=10
'*******************************************************************************
Private Sub tmrFader_Timer()
  On Error Resume Next                    'in case inc moves < 0 or > 255...
  If m_FaderDown Then
    m_Fader = m_Fader - m_FaderInc        'fading out
    If Err.Number Then m_Fader = 0        'set minimum on error
  Else
    m_Fader = m_Fader + m_FaderInc        'fading in
    If Err.Number Then m_Fader = 255      'set maximum on error
  End If
  If m_Fader = 255 Or m_Fader = 0 Then    'at either extent?
    Me.tmrFader.Enabled = False           'yes, so disable the timer
    Me.Enabled = True                     'enable the form
  End If
  SetWindowTranslucency Me.hwnd, m_Fader  'set final fade
'
' if fading down...
'
  If m_Fader = 0 Then
    AllowFormFading = False               'force flag off to properly unload
    Unload Me                             'unload if faded out
  End If
'
' if fading up...
'
  If m_Fader = 255 Then
    If SplashIsOn Then                    'if we are being used as a splash screen
      Me.Timer1.Enabled = True            'turn on the normal timer
    End If
  End If
End Sub
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

