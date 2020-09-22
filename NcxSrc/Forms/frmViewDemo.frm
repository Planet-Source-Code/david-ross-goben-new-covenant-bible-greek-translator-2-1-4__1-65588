VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmViewDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Application Demo for New Covenant Bible Greek Translator"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   Icon            =   "frmViewDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4155
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      ExtentX         =   11456
      ExtentY         =   6271
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmViewDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  On Error Resume Next
  Me.WebBrowser1.Navigate2 AddSlash(App.Path) & "DB\NCXlateDemo_Viewlet.html"
  If Err.Number <> 0 Then
    MessageBox frmGrkXlate, "Error Loading Demo. Aborting Load...", vbOKOnly Or vbExclamation, "Error Encountered"
    Unload Me
  Else
    Me.Width = 14600
    Me.Height = 9000
  End If
  frmGrkXlate.mnuHLPViewDemo.Enabled = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
  With Me.WebBrowser1
    .Left = 0
    .Top = -1240
    .Width = 20000
    .Height = 20000
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmGrkXlate.mnuHLPViewDemo.Enabled = True
End Sub
