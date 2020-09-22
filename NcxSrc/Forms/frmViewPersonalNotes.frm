VERSION 5.00
Begin VB.Form frmViewPersonalNotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View All Personal Notes"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   Icon            =   "frmViewPersonalNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
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
      TabIndex        =   13
      ToolTipText     =   "View with these options in the Definition Panel"
      Top             =   1260
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4500
      TabIndex        =   12
      ToolTipText     =   "Abort this operation, do not save any changes"
      Top             =   1260
      Width           =   915
   End
   Begin VB.CheckBox chkIncChpH 
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   540
      Width           =   195
   End
   Begin VB.CheckBox chkIncTheo 
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1140
      Width           =   195
   End
   Begin VB.CheckBox chkIncVerse 
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   195
   End
   Begin VB.CheckBox chkIncBkH 
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   195
   End
   Begin VB.CheckBox chkCtrChpH 
      Height          =   195
      Left            =   2460
      TabIndex        =   7
      Top             =   540
      Width           =   195
   End
   Begin VB.CheckBox chkCtrBkH 
      Height          =   195
      Left            =   2460
      TabIndex        =   3
      Top             =   240
      Width           =   195
   End
   Begin VB.Label lblIncChpH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Include &Chapter Headings"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "When a new book is started, bein on a new page, rather than at the bottom of the previous"
      Top             =   540
      Width           =   1845
   End
   Begin VB.Label lblIncTheo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Include &Theological Notes"
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   1140
      Width           =   1860
   End
   Begin VB.Label lblIncVerse 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Include &verse from currently selected Bible"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   2985
   End
   Begin VB.Label lblIncBkH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Include &Book Headings"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "When a new book is started, bein on a new page, rather than at the bottom of the previous"
      Top             =   240
      Width           =   1665
   End
   Begin VB.Label lblCtrChpH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Center Chapte&r Headings on the page."
      Height          =   195
      Left            =   2700
      TabIndex        =   6
      ToolTipText     =   "Place chapter titled titles in the center of the page line"
      Top             =   540
      Width           =   2730
   End
   Begin VB.Label lblCtrBkH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Center Boo&k headings on the page."
      Height          =   195
      Left            =   2700
      TabIndex        =   2
      ToolTipText     =   "Place book titles in the center of the page line"
      Top             =   240
      Width           =   2520
   End
End
Attribute VB_Name = "frmViewPersonalNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'
' get default settings
'
  Me.chkIncBkH.Value = CLng(GetSetting(App.Title, "Settings", "PNIncBkH", "0"))
  Me.chkIncChpH.Value = CLng(GetSetting(App.Title, "Settings", "PNIncChpH", "0"))
  Me.chkCtrBkH.Value = CLng(GetSetting(App.Title, "Settings", "PNCtrBkH", "0"))
  Me.chkCtrChpH.Value = CLng(GetSetting(App.Title, "Settings", "PNCtrChpH", "0"))
  Me.chkIncVerse.Value = CLng(GetSetting(App.Title, "Settings", "PNIncVerse", "0"))
  Me.chkIncTheo.Value = CLng(GetSetting(App.Title, "Settings", "PNIncTheo", "0"))
  Call chkIncBkH_Click    'ensure addtional options processed
  Call chkIncChpH_Click
  bCancel = True
End Sub

Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdView_Click
' Purpose           : View the stuff
'*******************************************************************************
Private Sub cmdView_Click()
  Call SaveSetting(App.Title, "Settings", "PNIncBkH", CStr(Me.chkIncBkH.Value))
  Call SaveSetting(App.Title, "Settings", "PNIncChpH", CStr(Me.chkIncChpH.Value))
  Call SaveSetting(App.Title, "Settings", "PNCtrBkH", CStr(Me.chkCtrBkH.Value))
  Call SaveSetting(App.Title, "Settings", "PNCtrChpH", CStr(Me.chkCtrChpH.Value))
  Call SaveSetting(App.Title, "Settings", "PNIncVerse", CStr(Me.chkIncVerse.Value))
  Call SaveSetting(App.Title, "Settings", "PNIncTheo", CStr(Me.chkIncTheo.Value))
  bCancel = False
  Unload Me
End Sub

Private Sub chkIncBkH_Click()
  Me.chkCtrBkH.Enabled = CBool(Me.chkIncBkH.Value)
  Me.lblCtrBkH.Enabled = Me.chkCtrBkH.Enabled
End Sub

Private Sub lblIncBkH_Click()
  With Me.chkIncBkH
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = Checked
    End If
  End With
End Sub

Private Sub chkIncChpH_Click()
  Me.chkCtrChpH.Enabled = CBool(Me.chkIncChpH.Value)
  Me.lblCtrChpH.Enabled = Me.chkCtrChpH.Enabled
End Sub
Private Sub lblIncChpH_Click()
  With Me.chkIncChpH
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = Checked
    End If
  End With
End Sub

Private Sub lblCtrBkH_Click()
  With Me.chkCtrBkH
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = Checked
    End If
  End With
End Sub

Private Sub lblCtrChpH_Click()
  With Me.chkCtrChpH
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = Checked
    End If
  End With
End Sub

Private Sub lblIncTheo_Click()
  With Me.chkIncTheo
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = Checked
    End If
  End With
End Sub

Private Sub lblIncVerse_Click()
  With Me.chkIncVerse
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = Checked
    End If
  End With
End Sub

