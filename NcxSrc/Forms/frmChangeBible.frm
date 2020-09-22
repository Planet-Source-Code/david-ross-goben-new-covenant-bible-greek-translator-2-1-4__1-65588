VERSION 5.00
Begin VB.Form frmChangeBible 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Bible"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   360
      ScaleHeight     =   1815
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   300
      Width           =   3795
      Begin VB.OptionButton Option1 
         Caption         =   "King James Bible (1611)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   3675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Young's Literal Translation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   375
         Width           =   3675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Revised Standard Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   750
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "My Personal Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1260
         Width           =   3495
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   3780
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2760
      TabIndex        =   1
      Top             =   2340
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   420
      TabIndex        =   0
      Top             =   2340
      Width           =   1395
   End
End
Attribute VB_Name = "frmChangeBible"
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
  Me.Option1(BblVersion).Value = True
  Me.Option1(3).Enabled = PersonalVersion
  Me.cmdOk.Enabled = False
  Me.Picture1.BackColor = cMedium
  Me.Option1(0).BackColor = cMedium
  Me.Option1(1).BackColor = cMedium
  Me.Option1(2).BackColor = cMedium
  Me.Option1(3).BackColor = cMedium
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
Private Sub cmdOk_Click()
  BblVersion = Bible
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : Option1_Click
' Purpose           : Enable OK only when the choice is not the current
'*******************************************************************************
Private Sub Option1_Click(Index As Integer)
  Bible = Index
  Me.cmdOk.Enabled = Index <> BblVersion
End Sub
