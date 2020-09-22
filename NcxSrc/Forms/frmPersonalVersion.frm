VERSION 5.00
Begin VB.Form frmPersonalVersion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Your Own Personal Bible Version"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Select Base Model for your Personal Verison:"
      Height          =   2655
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4035
      Begin VB.PictureBox Picture1 
         Height          =   2355
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   3795
         TabIndex        =   3
         Top             =   240
         Width           =   3855
         Begin VB.OptionButton Option1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   0
            TabIndex        =   18
            Top             =   1500
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   0
            TabIndex        =   16
            Top             =   300
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   600
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   0
            TabIndex        =   8
            Top             =   1200
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   2040
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   0
            TabIndex        =   6
            Top             =   900
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   0
            TabIndex        =   5
            Top             =   1770
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   195
         End
         Begin VB.Label lblOpt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Web&ster's Translation (1611)"
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
            Index           =   8
            Left            =   240
            TabIndex        =   19
            Top             =   1500
            Width           =   2985
         End
         Begin VB.Label lblOpt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Darby's Translation"
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
            Index           =   7
            Left            =   240
            TabIndex        =   17
            Top             =   300
            Width           =   2055
         End
         Begin VB.Label lblOpt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&King James Version (1611)"
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
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   2760
         End
         Begin VB.Label lblOpt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Young's Literal Translation"
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
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   2040
            Width           =   2790
         End
         Begin VB.Label lblOpt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Revised Standard Version"
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
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   1200
            Width           =   2745
         End
         Begin VB.Label lblOpt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Modern King James Version"
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
            Index           =   4
            Left            =   240
            TabIndex        =   11
            Top             =   900
            Width           =   2910
         End
         Begin VB.Label lblOpt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&World English Bible"
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
            Index           =   5
            Left            =   240
            TabIndex        =   10
            Top             =   1770
            Width           =   2055
         End
         Begin VB.Label lblOpt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&American Standard Version (1917)"
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
            Index           =   6
            Left            =   240
            TabIndex        =   9
            Top             =   0
            Width           =   3540
         End
      End
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
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   3060
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2820
      TabIndex        =   1
      Top             =   3060
      Width           =   1395
   End
End
Attribute VB_Name = "frmPersonalVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Bible As Long         'local storage of tentative bible version selection

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Init no choice yet made, set the current version as the default
'*******************************************************************************
Private Sub Form_Load()
  Me.Icon = frmGrkXlate.Icon
  PersonalVersionBase = -1
  Select Case BblVersion
    Case 3                                'if current version is custom...
      Me.Option1(0).Value = True          'select KJV
    Case Else
      Me.Option1(BblVersion).Value = True 'else active bible
  End Select
  Me.Option1(0).BackColor = cMedium
  Me.Option1(1).BackColor = cMedium
  Me.Option1(2).BackColor = cMedium
  Me.Option1(4).BackColor = cMedium
  Me.Option1(5).BackColor = cMedium
  Me.Frame1.BackColor = cMedium
  Me.Picture1.BackColor = cMedium
  Me.Picture1.BorderStyle = 0
End Sub

Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Ignore any selections
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : Accept user choice for a bible version
'*******************************************************************************
Private Sub cmdOK_Click()
  PersonalVersionBase = Bible
  Unload Me
End Sub

Private Sub lblOpt_Click(Index As Integer)
  Me.Option1(Index) = True
End Sub

'*******************************************************************************
' Subroutine Name   : Option1_Click
' Purpose           : User made a selection
'*******************************************************************************
Private Sub Option1_Click(Index As Integer)
  Bible = Index
End Sub

