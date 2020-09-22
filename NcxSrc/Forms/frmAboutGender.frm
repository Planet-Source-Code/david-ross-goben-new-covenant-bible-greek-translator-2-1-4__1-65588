VERSION 5.00
Begin VB.Form frmAboutGender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Gender Checking"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   Icon            =   "frmAboutGender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3578
      TabIndex        =   2
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmAboutGender.frx":000C
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-h, -aiV, -ioV, -ia, -ka, -ra, -qa, -hV, -hn, -ai, -aV, -eV, -ew, -iV, -ar, -wV"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   2040
      Width           =   7935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   345
      X2              =   7965
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   345
      X2              =   7965
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAboutGender.frx":08D6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   345
      TabIndex        =   0
      Top             =   660
      Width           =   7635
   End
End
Attribute VB_Name = "frmAboutGender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : OK button
'*******************************************************************************
Private Sub cmdOK_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Paint texture onto background
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
End Sub
