VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "InputBox"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   ControlBox      =   0   'False
   Icon            =   "frmInputBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboInputBox 
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Text            =   "cboInputBox"
      Top             =   1740
      Width           =   5655
   End
   Begin VB.CheckBox chkCheckVineAll 
      Caption         =   "Check1"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   195
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Text            =   "txtInput"
      Top             =   1260
      Width           =   5655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblChkVineAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check the full Vine database text if the word is not found in the word list"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1020
      Width           =   5040
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      Height          =   795
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private InputResponse As String            'response from inputbox

Private Sub cboInputBox_GotFocus()
  With Me.cboInputBox
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub cmdCancel_Click()
  InputResponse = vbNullString
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  Dim S As String
  Dim lcol As Collection
  Dim I As Long
  Dim FndIt As Boolean
  
  SaveSetting App.Title, "Settings", "CheckVineAll", CStr(Me.chkCheckVineAll.Value)
  If Not Me.txtInput.Visible Then
    Set lcol = New Collection
    With Me.cboInputBox
      For I = 0 To .ListCount - 1
        S = LCase$(.List(I))
        lcol.Add S, S
      Next I
      S = LCase$(Me.cboInputBox.Text)
      On Error Resume Next
      lcol.Add S, S
      If Err.Number <> 0 Then
        FndIt = True
        For I = 1 To lcol.Count
          If S = lcol(I) Then Exit For
        Next I
      End If
      On Error GoTo 0
    End With
    Set lcol = Nothing
    If Not FndIt Then
      Me.cboInputBox.AddItem Me.cboInputBox.Text
      LastInputBox = Me.cboInputBox.NewIndex
    Else
      LastInputBox = I - 1
    End If
    Me.txtInput.Text = Me.cboInputBox.Text
    If Len(Me.cboInputBox.Text) Then
      SaveList
    End If
  End If
  S = Me.txtInput.Text
  If Len(S) <> 0 Then InputResponse = S
  Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
    KeyCode = 0
    Me.cmdCancel.Value = True
  End If
End Sub

Public Function dwInputBox(iForm As Form, _
                           Prompt As String, _
                           Optional PromptCaption As String, _
                           Optional Default As String, _
                           Optional ShowOption As String = vbNullString, _
                           Optional IsVine As Boolean = False) As String
  Dim S As String
  Dim ShwOpt As Boolean
  Dim I As Long
  
  Me.lblChkVineAll.Caption = ShowOption
  ShwOpt = CBool(Len(ShowOption))
  Me.chkCheckVineAll.Visible = ShwOpt
  Me.lblChkVineAll.Visible = ShwOpt
  
  S = Trim$(PromptCaption)
  If Len(S) = 0 Then S = App.Title
  Me.Caption = S
  S = Trim$(Prompt)
  If Len(S) = 0 Then S = "Enter Response Text:"
  Me.lblMsg.Caption = S
  Me.txtInput.Text = Default
  InputResponse = vbNullString
  
  Me.txtInput.Visible = False
  Me.cboInputBox.Visible = False
  
  If IsVine Then
    With Me.cboInputBox
      Me.cboInputBox.Visible = True
      .Top = Me.txtInput.Top
      .Clear
      For I = 0 To frmGrkXlate.lstImputBox.ListCount - 1
        .AddItem frmGrkXlate.lstImputBox.List(I)
      Next I
      If .ListCount <> 0 Then
        .ListIndex = LastInputBox
        .Text = .List(.ListIndex)
      Else
        .Text = Me.txtInput.Text
      End If
    End With
  Else
    Me.txtInput.Visible = True
  End If
  
  Me.Height = (Me.Height - Me.ScaleHeight) + Me.txtInput.Top + Me.txtInput.Height + 120
  
  frmInputBox.Show vbModal, iForm
  dwInputBox = InputResponse
End Function

Private Sub Form_Load()
  Me.chkCheckVineAll.Value = CLng(GetSetting(App.Title, "Settings", "CheckVineAll", CStr(vbChecked)))
End Sub

Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
  On Error Resume Next
  Me.txtInput.SetFocus
  If Err.Number <> 0 Then
    Me.cboInputBox.SetFocus
  End If
End Sub

Private Sub SaveList()
  Dim S As String
  Dim I As Long
  
  S = Me.cboInputBox.Text
  With frmGrkXlate.lstImputBox
    .Clear
    Do While Me.cboInputBox.ListCount
      .AddItem Me.cboInputBox.List(0)
      Me.cboInputBox.RemoveItem 0
    Loop
    For I = 0 To .ListCount - 1
      If S = .List(I) Then
        LastInputBox = I
        Exit For
      End If
    Next I
  End With
End Sub

Private Sub lblChkVineAll_Click()
  With Me.chkCheckVineAll
    If .Value = vbChecked Then
      .Value = vbUnchecked
    Else
      .Value = vbChecked
    End If
  End With
End Sub

Private Sub txtInput_GotFocus()
  With Me.txtInput
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub
