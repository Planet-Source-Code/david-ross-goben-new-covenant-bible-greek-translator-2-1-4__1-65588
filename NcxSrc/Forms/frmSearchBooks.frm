VERSION 5.00
Begin VB.Form frmSearchBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Bible Books to Search Through"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGospels 
      Caption         =   "Gospels"
      Height          =   375
      Left            =   2070
      TabIndex        =   6
      ToolTipText     =   "Select only the Gospels"
      Top             =   900
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Covenant Books to Search Through"
      Height          =   555
      Left            =   240
      TabIndex        =   4
      Top             =   180
      Width           =   4335
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmSearchBooks.frx":0000
         Left            =   120
         List            =   "frmSearchBooks.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   180
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set All"
      Height          =   375
      Left            =   1185
      TabIndex        =   3
      ToolTipText     =   "Set all to selected"
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   300
      TabIndex        =   2
      ToolTipText     =   "Clear all selections"
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2955
      TabIndex        =   0
      ToolTipText     =   "Accept selected books"
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Ignore any changes"
      Top             =   900
      Width           =   735
   End
End
Attribute VB_Name = "frmSearchBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Const LB_GETITEMHEIGHT = &H1A1

Private Sub cmdClear_Click()
  Dim Idx As Long
  
  For Idx = 0 To 26
    Me.List1.Selected(Idx) = False
  Next Idx
  Me.cmdOK.Enabled = False
  Me.List1.ListIndex = -1
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdGospels_Click()
  Dim Idx As Long
  
  Me.cmdClear.Value = True
  For Idx = 0 To 3
    Me.List1.Selected(Idx) = True
  Next Idx
  Me.List1.ListIndex = -1
End Sub

Private Sub cmdOK_Click()
  Dim Idx As Long
  
  For Idx = 0 To 26
    SearchBooks(Idx + 1) = Me.List1.Selected(Idx)
  Next Idx
  Unload Me
End Sub

Private Sub cmdSet_Click()
  Dim Idx As Long
  
  For Idx = 0 To 26
    Me.List1.Selected(Idx) = True
  Next Idx
  Me.cmdOK.Enabled = True
  Me.List1.ListIndex = -1
End Sub

Private Sub Form_Load()
  Dim Idx As Long, I As Long
  Dim S As String, Ary() As String
  
  Me.Icon = frmSearch.Icon
  Me.Frame1.BackColor = cMedium
  With Me.List1
    .BackColor = cLight
    For Idx = 1 To 27
      Ary = Split(Books(Idx), ",")
      .AddItem Ary(3)
      .Selected(Idx - 1) = SearchBooks(Idx)
    Next Idx
    Idx = SendMessageByNum(.hwnd, LB_GETITEMHEIGHT, 0&, 0&) * Screen.TwipsPerPixelY
    .Height = Idx * 28
    Me.Frame1.Height = Me.Frame1.Height + Idx * 26
    Me.Height = Me.Height + Idx * 26
    Me.cmdClose.Top = Me.ScaleHeight - Me.cmdClose.Height - 120
    Me.cmdOK.Top = Me.cmdClose.Top
    Me.cmdClear.Top = Me.cmdClose.Top
    Me.cmdSet.Top = Me.cmdClose.Top
    Me.cmdGospels.Top = Me.cmdClose.Top
    .ListIndex = -1
  End With
End Sub

Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)
End Sub

Private Sub List1_ItemCheck(Item As Integer)
  Me.cmdOK.Enabled = CBool(Me.List1.SelCount)
End Sub
