VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPersonalNotes 
   Caption         =   "Personal Verse Note"
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10500
   Icon            =   "frmPersonalNotes.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   315
      Left            =   8520
      TabIndex        =   114
      ToolTipText     =   "Apply changes and close this dialog"
      Top             =   7260
      Width           =   855
   End
   Begin VB.CommandButton cmdHowTo 
      Height          =   555
      Left            =   9240
      Picture         =   "frmPersonalNotes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   113
      ToolTipText     =   "How to use this table interactively"
      Top             =   3480
      Width           =   555
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Move to Insert Point"
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      ToolTipText     =   "Insert the contents of this field into the text at the cursor position (braces { } will be added if they are not present)"
      Top             =   7380
      Width           =   1695
   End
   Begin VB.TextBox txtGreek 
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      HideSelection   =   0   'False
      Left            =   2220
      TabIndex        =   3
      ToolTipText     =   "Use this field to test Greek spelling"
      Top             =   7380
      Width           =   3075
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   3255
      Left            =   -60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmPersonalNotes.frx":11D4
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   7815
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   17568
            Text            =   "Enbrace Greek words with curly braces ""{word or words}"" and they will be displayed in Greek text on the main form."
            TextSave        =   "Enbrace Greek words with curly braces ""{word or words}"" and they will be displayed in Greek text on the main form."
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9540
      TabIndex        =   5
      ToolTipText     =   "Disregard any changes made to this text"
      Top             =   7260
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   4020
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   5636
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmPersonalNotes.frx":125F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picAlpha 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9315
      TabIndex        =   8
      Top             =   3480
      Width           =   9315
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4320
         TabIndex        =   112
         Top             =   0
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4575
         TabIndex        =   111
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   4755
         TabIndex        =   110
         Top             =   0
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   645
         TabIndex        =   109
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   855
         TabIndex        =   108
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   1050
         TabIndex        =   107
         Top             =   0
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   1230
         TabIndex        =   106
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   1440
         TabIndex        =   105
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   1650
         TabIndex        =   104
         Top             =   0
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   1755
         TabIndex        =   103
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   1920
         TabIndex        =   102
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   255
         TabIndex        =   101
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   2265
         TabIndex        =   100
         Top             =   0
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   60
         TabIndex        =   99
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   2700
         TabIndex        =   98
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   2910
         TabIndex        =   97
         Top             =   0
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   3105
         TabIndex        =   96
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   3315
         TabIndex        =   95
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   3525
         TabIndex        =   94
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   3720
         TabIndex        =   93
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   3915
         TabIndex        =   92
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   4125
         TabIndex        =   91
         Top             =   0
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   2100
         TabIndex        =   90
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   23
         Left            =   2490
         TabIndex        =   89
         Top             =   0
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   450
         TabIndex        =   88
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   4950
         TabIndex        =   87
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "w"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   52
         Left            =   8595
         TabIndex        =   86
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   53
         Left            =   8790
         TabIndex        =   85
         Top             =   0
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   54
         Left            =   8940
         TabIndex        =   84
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "d"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   55
         Left            =   5655
         TabIndex        =   83
         Top             =   0
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   56
         Left            =   5835
         TabIndex        =   82
         Top             =   0
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   57
         Left            =   6015
         TabIndex        =   81
         Top             =   0
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "g"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   58
         Left            =   6120
         TabIndex        =   80
         Top             =   0
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "h"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   59
         Left            =   6300
         TabIndex        =   79
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   60
         Left            =   6465
         TabIndex        =   78
         Top             =   0
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "j"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   61
         Left            =   6570
         TabIndex        =   77
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "k"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   62
         Left            =   6675
         TabIndex        =   76
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   63
         Left            =   5310
         TabIndex        =   75
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   64
         Left            =   6945
         TabIndex        =   74
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   65
         Left            =   5130
         TabIndex        =   73
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   66
         Left            =   7335
         TabIndex        =   72
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   67
         Left            =   7515
         TabIndex        =   71
         Top             =   0
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   68
         Left            =   7695
         TabIndex        =   70
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   69
         Left            =   7875
         TabIndex        =   69
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   70
         Left            =   7995
         TabIndex        =   68
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   71
         Left            =   8160
         TabIndex        =   67
         Top             =   0
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   72
         Left            =   8265
         TabIndex        =   66
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   73
         Left            =   8430
         TabIndex        =   65
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   74
         Left            =   6840
         TabIndex        =   64
         Top             =   0
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "n"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   75
         Left            =   7170
         TabIndex        =   63
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   76
         Left            =   5490
         TabIndex        =   62
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "z"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   77
         Left            =   9105
         TabIndex        =   61
         Top             =   0
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   26
         Left            =   4320
         TabIndex        =   60
         Top             =   300
         Width           =   195
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   27
         Left            =   4575
         TabIndex        =   59
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   28
         Left            =   4755
         TabIndex        =   58
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   29
         Left            =   645
         TabIndex        =   57
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   30
         Left            =   855
         TabIndex        =   56
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   31
         Left            =   1050
         TabIndex        =   55
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   32
         Left            =   1230
         TabIndex        =   54
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   33
         Left            =   1440
         TabIndex        =   53
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   34
         Left            =   1650
         TabIndex        =   52
         Top             =   300
         Width           =   45
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   35
         Left            =   1755
         TabIndex        =   51
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   36
         Left            =   1920
         TabIndex        =   50
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   37
         Left            =   255
         TabIndex        =   49
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   38
         Left            =   2265
         TabIndex        =   48
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   39
         Left            =   60
         TabIndex        =   47
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   40
         Left            =   2700
         TabIndex        =   46
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   41
         Left            =   2910
         TabIndex        =   45
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   42
         Left            =   3105
         TabIndex        =   44
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   43
         Left            =   3315
         TabIndex        =   43
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   44
         Left            =   3525
         TabIndex        =   42
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   45
         Left            =   3720
         TabIndex        =   41
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   46
         Left            =   3915
         TabIndex        =   40
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   47
         Left            =   4125
         TabIndex        =   39
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   48
         Left            =   2100
         TabIndex        =   38
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   49
         Left            =   2490
         TabIndex        =   37
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   50
         Left            =   450
         TabIndex        =   36
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   51
         Left            =   4950
         TabIndex        =   35
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "w"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   78
         Left            =   8595
         TabIndex        =   34
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   79
         Left            =   8790
         TabIndex        =   33
         Top             =   300
         Width           =   90
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   80
         Left            =   8940
         TabIndex        =   32
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "d"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   81
         Left            =   5655
         TabIndex        =   31
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   82
         Left            =   5835
         TabIndex        =   30
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   83
         Left            =   6015
         TabIndex        =   29
         Top             =   300
         Width           =   45
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "g"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   84
         Left            =   6120
         TabIndex        =   28
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "h"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   85
         Left            =   6300
         TabIndex        =   27
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   86
         Left            =   6465
         TabIndex        =   26
         Top             =   300
         Width           =   45
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "j"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   87
         Left            =   6570
         TabIndex        =   25
         Top             =   300
         Width           =   45
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "k"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   88
         Left            =   6675
         TabIndex        =   24
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "b"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   89
         Left            =   5310
         TabIndex        =   23
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   90
         Left            =   6945
         TabIndex        =   22
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   91
         Left            =   5130
         TabIndex        =   21
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   92
         Left            =   7335
         TabIndex        =   20
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   93
         Left            =   7515
         TabIndex        =   19
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   94
         Left            =   7695
         TabIndex        =   18
         Top             =   300
         Width           =   120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   95
         Left            =   7875
         TabIndex        =   17
         Top             =   300
         Width           =   60
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   96
         Left            =   7995
         TabIndex        =   16
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   97
         Left            =   8160
         TabIndex        =   15
         Top             =   300
         Width           =   45
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   98
         Left            =   8265
         TabIndex        =   14
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   99
         Left            =   8430
         TabIndex        =   13
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   100
         Left            =   6840
         TabIndex        =   12
         Top             =   300
         Width           =   45
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   101
         Left            =   7170
         TabIndex        =   11
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "c"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   102
         Left            =   5490
         TabIndex        =   10
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   103
         Left            =   9105
         TabIndex        =   9
         Top             =   300
         Width           =   90
      End
   End
   Begin VB.Label lbltestGreek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Greek Text Constructor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Form to save changes, select CANCEL to ignore changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5415
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuPopupCut 
         Caption         =   "&Cut selection"
      End
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "C&opy selection"
      End
      Begin VB.Menu mnuPopupPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuPopupSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupSelectAll 
         Caption         =   "&Select All"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "mnuPopup2"
      Begin VB.Menu mnuPopup2Copy 
         Caption         =   "Copy selection"
      End
   End
End
Attribute VB_Name = "frmPersonalNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsDirty As Boolean    'true if the data has changed
Private Verse As String       'verse data (kept for heading info)

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  bCancel = True
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopy_Click
' Purpose           : Copy test text to the main text
'*******************************************************************************
Private Sub cmdCopy_Click()
  Dim T As String
  Dim SS As Long
  Dim HaveBraces As Boolean
  
  With Me.RichTextBox1
    SS = .SelStart                                        'save the insertion point
    .SelText = vbNullString                               'clear out what will be placed
    If SS > 0 Then
      T = Mid$(.Text, SS, 2)                              'if if some welection was braced
      HaveBraces = T = "{}"                               'flag if so
    End If
    T = Trim$(Me.txtGreek.Text)                           'grab the text to insert
    Me.txtGreek.Text = vbNullString
    If HaveBraces Then
      If Left$(T, 1) = "{" Then T = Mid$(T, 2)            'trim as needed
      If Right$(T, 1) = "}" Then T = Left$(T, Len(T) - 1)
    Else
      If Left$(T, 1) <> "{" Then T = "{" & T              'or append as needed
      If Right$(T, 1) <> "}" Then T = T & "}"
    End If
  
    .SelStart = SS                                        'ensure the insertion point set
    .SelText = T                                          'stuff data
    .SetFocus
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdHowTo_Click
' Purpose           : Show a little HOWTO dialog
'*******************************************************************************
Private Sub cmdHowTo_Click()
  MessageBox Me, "Clicking a letter (in either row) will append it to the test field.", _
                 vbOKOnly Or vbInformation, "How to Use the Greek Alphabet Table"
End Sub

Private Sub Form_Load()
  Dim S As String
  Dim I As Long
  
  Me.mnuPopup.Visible = False
  Me.mnuPopup2.Visible = False
  
  With frmGrkXlate
    S = .Caption
    I = InStr(1, S, "-")
    Me.Caption = "Personal Verse Notes " & Mid$(S, I)
    Me.RichTextBox2.Text = vbNullString
    Me.RichTextBox2.SelRTF = .rtbGreek.TextRTF
    Me.RichTextBox2.SelText = vbCrLf & vbCrLf
    Me.RichTextBox2.SelRTF = .rtbVerse.TextRTF
  End With
  With Me.RichTextBox2
    .BackColor = cVLight ' clBlue
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelColor = vbBlack
    .SelLength = 0
  End With
  Me.cmdHowTo.BackColor = cMedium
  
'  With Me.txtGreek
'    .Text = "Logon Tou TheoV"
'    .SelStart = 0
'    .SelLength = Len(.Text)
'  End With
  
  With Me.RichTextBox1
    .Font.Size = FntSize
    Verse = MyNotes(VrsIdx)
    S = Mid$(Verse, 8)
    I = InStr(1, S, "\")
    Do While I <> 0
      S = Left$(S, I - 1) & vbCrLf & Mid$(S, I + 1)
      I = InStr(I + 2, S, "\")
    Loop
    .Text = S
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelHangingIndent = HIndent
    .SelStart = Len(.Text)
    .SelLength = 0
  End With
  Me.cmdCopy.Enabled = False
  IsDirty = False
  bCancel = False
  Me.cmdOK.Enabled = False
End Sub

Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)   'repaint background
End Sub

'*******************************************************************************
' Subroutine Name   : Form_QueryUnload
' Purpose           : Update the user notes if anything has changed
'*******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim S As String
  Dim I As Long
  
  If UnloadMode = 0 Or bCancel = False Then
    If Not IsDirty Then
      bCancel = True
      Exit Sub
    End If
    If bCancel Then Exit Sub
    
    S = Me.RichTextBox1.Text
    I = InStr(1, S, vbCrLf)
    Do While I <> 0
      S = Left$(S, I - 1) & "\" & Mid$(S, I + 2)
      I = InStr(I + 1, S, vbCrLf)
    Loop
    MyNotes(VrsIdx) = Left$(Verse, 6) & " " & S
    If Len(S) <> 0 Then HavePersonalNotes = True    'indicate we have personal notes
    MyNotesDirty = True
    AutoDirty = True
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Adjust form contents as needed
'*******************************************************************************
Private Sub Form_Resize()
  Dim I As Long
  Static Resizing As Boolean
  
  If Me.WindowState = vbMinimized Then Exit Sub
  If Resizing Then Exit Sub
  Resizing = True
  I = Me.Width - Me.ScaleWidth
  If Me.Width - I < Me.picAlpha.Width Then Me.Width = Me.picAlpha.Width + I
  If Me.Height < 7000 Then Me.Height = 7000
  Me.cmdCancel.Left = Me.ScaleWidth - Me.cmdCancel.Width - 60
  Me.cmdCancel.Top = Me.ScaleHeight - Me.cmdCancel.Height - 30 - Me.StatusBar1.Height
  Me.cmdOK.Left = Me.cmdCancel.Left - Me.cmdOK.Width - 120
  Me.cmdOK.Top = Me.cmdCancel.Top
  Me.RichTextBox1.Width = Me.ScaleWidth
  Me.RichTextBox1.Height = Me.cmdCancel.Top - Me.RichTextBox1.Top - 30
  Me.RichTextBox2.Width = Me.RichTextBox1.Width
  Me.txtGreek.Top = Me.cmdCancel.Top - 30
  Me.lbltestGreek.Top = Me.txtGreek.Top + 90
  Me.cmdCopy.Top = Me.cmdCancel.Top
  Me.cmdHowTo.Left = Me.ScaleWidth - Me.cmdHowTo.Width
  With Me.picAlpha
    I = (Me.ScaleWidth - .Width) \ 2
    If I < 0 Then I = 0
    .Left = I
  End With
  Resizing = False
End Sub

'*******************************************************************************
' Subroutine Name   : Label2_Click
' Purpose           : A letter was clicked, ass it to the test frame
'*******************************************************************************
Private Sub Label2_Click(Index As Integer)
  Dim T As String
  
  With Me.txtGreek
    T = .Text & Me.Label2(Index).Caption
    .Text = T
    .SelStart = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpCopy2_Click
' Purpose           : Save selection to the clipboard
'*******************************************************************************
Private Sub mnuPopUpCopy2_Click()
  Clipboard.Clear
  Clipboard.SetText Me.RichTextBox2.SelText, vbCFText
  Clipboard.SetText Me.RichTextBox2.SelRTF, vbCFRTF
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopUpCopy_Click
' Purpose           : Save selection to the clipboard
'*******************************************************************************
Private Sub mnuPopUpCopy_Click()
  Clipboard.Clear
  Clipboard.SetText Me.RichTextBox1.SelText, vbCFText
  Clipboard.SetText Me.RichTextBox1.SelRTF, vbCFRTF
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopupCut_Click
' Purpose           : Save selection to the clipboard
'*******************************************************************************
Private Sub mnuPopupCut_Click()
  Clipboard.Clear
  Clipboard.SetText Me.RichTextBox1.SelText, vbCFText
  Clipboard.SetText Me.RichTextBox1.SelRTF, vbCFRTF
  Me.RichTextBox1.SelText = vbNullString
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopupPaste_Click
' Purpose           : Get selection from the clipboard
'*******************************************************************************
Private Sub mnuPopupPaste_Click()
  Me.RichTextBox1.SelText = Clipboard.GetText(vbCFRTF)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuPopupSelectAll_Click
' Purpose           : Select all text
'*******************************************************************************
Private Sub mnuPopupSelectAll_Click()
  With Me.RichTextBox1
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub picAlpha_Paint()
  PaintTilePicBackground Me.picAlpha, frmGrkXlate.picTile(Background)   'repaint background
End Sub

'*******************************************************************************
' Subroutine Name   : RichTextBox1_Change
' Purpose           : Keep track of changed text
'*******************************************************************************
Private Sub RichTextBox1_Change()
  IsDirty = True
  Me.cmdOK.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : RichTextBox1_MouseDown
' Purpose           : Display options if user does a right-click on the textbox when
'                   : data is selected within it
'*******************************************************************************
Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton And Shift = 0 Then
    Me.mnuPopupCopy.Enabled = Me.RichTextBox1.SelLength <> 0
    Me.mnuPopupCut.Enabled = Me.mnuPopupCopy.Enabled
    Me.mnuPopupPaste.Enabled = CBool(Len(Clipboard.GetText))
    Me.mnuPopupSelectAll.Enabled = CBool(Len(Me.RichTextBox1.Text))
    PopupMenu Me.mnuPopup, vbPopupMenuRightButton
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : RichTextBox2_MouseDown
' Purpose           : Display options if user does a right-click on the textbox when
'                   : data is selected within it
'*******************************************************************************
Private Sub RichTextBox2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton And Shift = 0 And Me.RichTextBox2.SelLength <> 0 Then
    PopupMenu Me.mnuPopup2, vbPopupMenuRightButton
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : txtGreek_Change
' Purpose           : Update the tooltip for the test text as needed
'*******************************************************************************
Private Sub txtGreek_Change()
  Dim T As String
  
  T = Me.txtGreek.Text
  If Len(T) = 0 Then
    T = "Use this field to test Greek spelling"
    Me.cmdCopy.Enabled = False
  Else
    T = "Use this field to test Greek spelling. Latizined: " & T
    Me.cmdCopy.Enabled = True
  End If
  Me.txtGreek.ToolTipText = T
End Sub

'*******************************************************************************
' Subroutine Name   : txtGreek_GotFocus
' Purpose           : Highlight the entire text when it gets focus
'*******************************************************************************
Private Sub txtGreek_GotFocus()
  With Me.txtGreek
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub
