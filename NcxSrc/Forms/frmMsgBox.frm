VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   180
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   1500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   1500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2700
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   1500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox MyPicture 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   780
      Width           =   555
   End
   Begin VB.Label lbText 
      BackStyle       =   0  'Transparent
      Caption         =   "Message Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   180
      Width           =   2955
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   0
      Left            =   180
      Stretch         =   -1  'True
      Top             =   180
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   1
      Left            =   840
      Stretch         =   -1  'True
      Top             =   840
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   2
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   840
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   3
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   840
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   4
      Left            =   2700
      Stretch         =   -1  'True
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblTest 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Message Area"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1980
      Visible         =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'~frmMessageBox.frm;modMessageBox.bas;modGetStockObject.bas;
'Display a custom messagebox
'******************************************************************************
' This form, using modMessageBox.mod, displays a message box that is customizable.
' By setting IconImage(4).Picture to an image of your choice, and providing the
' vbSystemModal flag, your custom image will be displayed as the message icon.
'
'EXAMPLE
'  With frmMessageBox
'    .IconImage(4).Picture = MyImage.Picture 'custom image
'    If .dwMessageBox(Me, "Are you nuts?", vbYesNo Or vbSystemModal, "Sanity Check") = vbYes Then
'      .dwMessagebox Me, "Obviously"
'    Else
'      .dwMessageBox Me, "I doubt that"
'    End If
'  End With
'
' NOTE: This form uses frmMessageBox.bas.
' NOTE: This form uses modGetStockObject.bas.
'******************************************************************************

Private ButtonResult As VbMsgBoxResult

Public Function dwMessageBox(iForm As Form, Prompt As String, Optional Flags As VbMsgBoxStyle, Optional PromptCaption As String) As VbMsgBoxResult
  Dim i As Integer, LCnt As Integer, BtnCnt As Integer
  Dim Idx As Long, Jdx As Long, Wdth As Long
  Dim S As String
  Dim Flgs As VbMsgBoxStyle
'
'Initialize Buttons
'
  For i = 0 To 2
    cmdButton(i).Visible = False
  Next i
  cmdButton(0).Cancel = False
  cmdButton(1).Cancel = False
  cmdButton(2).Cancel = False
  
  Flgs = Flags And &HF
  If Flgs = vbOKOnly Then
    SetButton 0, "OK"
    cmdButton(0).Cancel = True
    BtnCnt = 1
  ElseIf Flgs = vbOKCancel Then
    SetButton 0, "OK"
    SetButton 1, "Cancel"
    cmdButton(1).Cancel = True
    BtnCnt = 2
  ElseIf Flgs = vbAbortRetryIgnore Then
    SetButton 0, "Abort"
    SetButton 1, "Retry"
    SetButton 2, "Ignore"
    cmdButton(2).Cancel = True
    BtnCnt = 3
  ElseIf Flgs = vbRetryCancel Then
    SetButton 0, "Retry"
    SetButton 1, "Cancel"
    cmdButton(1).Cancel = True
    BtnCnt = 2
  ElseIf Flags And vbYesNo Then
    SetButton 0, "Yes"
    SetButton 1, "No"
    cmdButton(1).Cancel = True
    BtnCnt = 2
  ElseIf Flgs = vbYesNoCancel Then
    SetButton 0, "Yes"
    SetButton 1, "No"
    SetButton 2, "Cancel"
    cmdButton(2).Cancel = True
    BtnCnt = 3
  End If
'
'Initialize default button
'
  If Flags And vbDefaultButton2 Then
    cmdButton(1).Default = True
  ElseIf Flags And vbDefaultButton3 Then
    cmdButton(2).Default = True
  Else
    cmdButton(0).Default = True
  End If
'
'Initialize icon
'
  For i = 0 To 4
    Me.IconImage(i).Visible = False
    If i Then
      Me.IconImage(i).Left = Me.IconImage(0).Left
      Me.IconImage(i).Top = Me.IconImage(0).Top
    End If
  Next i
  
  Flgs = Flags And &HF0
  If Flgs = vbCritical Then
    IconImage(0).Visible = True
  ElseIf Flgs = vbExclamation Then
    IconImage(2).Visible = True
  ElseIf Flgs = vbInformation Then
    IconImage(3).Visible = True
  ElseIf Flgs = vbQuestion Then
    IconImage(1).Visible = True
  ElseIf (Flags And &HF000) = vbSystemModal Then
    IconImage(4).Visible = True
  End If
'
'Displaytext
'
  S = Trim$(Prompt)
  lbText.Caption = S
  If Len(Trim$(PromptCaption)) Then
    Me.Caption = Trim$(PromptCaption)
  Else
    Me.Caption = "Message"
  End If
'
' compute maximum text width
'
  Jdx = 1                                       'init to start of text
  Wdth = 0                                      'minimal width in twips
  LCnt = 0                                      'line count to 0
  With Me.lblTest
    Do While Jdx < Len(S)                       'while we are not done
      Idx = InStr(Jdx, S, vbCrLf)               'find a line
      If Idx = 0 Then Idx = Len(S) + 1          'was last/only
      .Caption = Mid$(S, Jdx, Idx - Jdx)        'stuff to text label
      If Wdth < .Width Then Wdth = .Width       'keep max widtrh
      LCnt = LCnt + 1                           'count a line
      Jdx = Idx + 2                             'point to next line
    Loop
    Wdth = Wdth + 240                           'allow gapping on right margin
  End With
'
' set display label width and height
'
  Me.lbText.Width = Wdth                        'set label width
  Me.lbText.Height = LCnt * Me.lbText.Height + 120
'
' compute minimal width (button-wise)
'
  Idx = (Me.cmdButton(0).Width + 120) * BtnCnt + 480 + (Me.Width - Me.ScaleWidth)
  Jdx = (Me.Width - Me.ScaleWidth) + Wdth + Me.lbText.Left + 60
  If Jdx < Idx Then Jdx = Idx
  Me.Width = Jdx
  Me.Height = Me.lbText.Height + (Me.Height - Me.ScaleHeight) + Me.lbText.Top + Me.cmdButton(0).Height + 480
  
  Me.cmdButton(0).Top = Me.ScaleHeight - Me.cmdButton(0).Height - 120
  Me.cmdButton(1).Top = Me.cmdButton(0).Top
  Me.cmdButton(2).Top = Me.cmdButton(0).Top
  Select Case BtnCnt
    Case 1
      Me.cmdButton(0).Left = (Me.ScaleWidth - Me.cmdButton(0).Width) / 2
    Case 2
      Me.cmdButton(0).Left = (Me.ScaleWidth - (Me.cmdButton(0).Width + 120) * 2) / 2
      Me.cmdButton(1).Left = Me.cmdButton(0).Left + Me.cmdButton(0).Width + 120
    Case 3
      Me.cmdButton(0).Left = (Me.ScaleWidth - (Me.cmdButton(0).Width + 120) * 3) / 2
      Me.cmdButton(1).Left = Me.cmdButton(0).Left + Me.cmdButton(0).Width + 120
      Me.cmdButton(2).Left = Me.cmdButton(1).Left + Me.cmdButton(1).Width + 120
  End Select
  frmMsgBox.Show 1, iForm
  dwMessageBox = ButtonResult
  Unload Me
End Function

Private Sub SetButton(ButtonIndex As Integer, PromptText As String)
  cmdButton(ButtonIndex).Caption = PromptText
  cmdButton(ButtonIndex).Visible = True
End Sub

Private Sub cmdButton_Click(Index As Integer)
  Select Case cmdButton(Index).Caption
    Case "Ok"
      ButtonResult = vbOK
    Case "Cancel"
      ButtonResult = vbCancel
    Case "Ignore"
      ButtonResult = vbIgnore
    Case "Yes"
      ButtonResult = vbYes
    Case "No"
      ButtonResult = vbNo
    Case "Retry"
      ButtonResult = vbRetry
    Case "Abort"
      ButtonResult = vbAbort
  End Select
  Me.Hide
End Sub

Private Sub Form_Load()
  With Me.MyPicture
'
' load stock objects into the image files via the picturebox
'
    GetStockObject IDI_HAND, Me.MyPicture
    Me.IconImage(0).Picture = .Image
    GetStockObject IDI_QUESTION, Me.MyPicture
    Me.IconImage(1).Picture = .Image
    GetStockObject IDI_EXCLAMATION, Me.MyPicture
    Me.IconImage(2).Picture = .Image
    GetStockObject IDI_ASTERISK, Me.MyPicture
    Me.IconImage(3).Picture = .Image
    .Picture = LoadPicture(vbNullString)
    .Visible = False
  End With
  Set Me.lblTest.Font = Me.lbText.Font
End Sub

Private Sub Form_Paint()
'
' set the focus on the desired button
'
  If Me.cmdButton(0).Default Then
    Me.cmdButton(0).SetFocus
  ElseIf Me.cmdButton(1).Default Then
    Me.cmdButton(1).SetFocus
  Else
    Me.cmdButton(2).SetFocus
  End If
End Sub


Public Function MessageBox(iForm As Form, Prompt As String, Optional Flags As VbMsgBoxStyle, Optional PromptCaption As String = vbNullString) As VbMsgBoxResult
  MessageBox = frmMsgBox.dwMessageBox(iForm, Prompt, Flags, PromptCaption)
End Function

