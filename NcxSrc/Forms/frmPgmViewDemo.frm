VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPgmViewDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Application Demo for New Covenant Bible Greek Translator"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "frmPgmViewDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   4048
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmPgmViewDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  If Len(Dir$(AddSlash(App.Path) & "ViewDemo\NCXlateDemo_Viewlet.html")) = 0 Then
    MsgBox "Error Locating Demo. Aborting Load...", vbOKOnly Or vbExclamation, "Error Encountered"
    Unload Me
  Else
    On Error Resume Next
    Me.WebBrowser1.Navigate2 AddSlash(App.Path) & "ViewDemo\NCXlateDemo_Viewlet.html"
    If Err.Number <> 0 Then
      MsgBox "Error Loading Demo. Aborting Load...", vbOKOnly Or vbExclamation, "Error Encountered"
      Unload Me
    Else
      Me.Width = 14820
      Me.Height = 9000
    End If
  End If
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
