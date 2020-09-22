VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Search Bibles for Text Matches"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHowTo 
      Height          =   555
      Left            =   9540
      Picture         =   "frmSearch.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   118
      ToolTipText     =   "How to use this table interactively"
      Top             =   1980
      Width           =   555
   End
   Begin VB.ListBox lstRef 
      Height          =   255
      Left            =   7740
      TabIndex        =   117
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      ToolTipText     =   "Copy list to the clipboard"
      Top             =   5220
      Width           =   915
   End
   Begin VB.Frame frameResults 
      Caption         =   "Results - Click to Navigate to them"
      Height          =   2595
      Left            =   180
      TabIndex        =   8
      Top             =   2520
      Width           =   9855
      Begin VB.ListBox lstSearch 
         Height          =   2205
         Left            =   120
         MouseIcon       =   "frmSearch.frx":1014
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame frameCriteria 
      Caption         =   "Search Criteria"
      Height          =   1755
      Left            =   180
      TabIndex        =   7
      Top             =   240
      Width           =   9855
      Begin VB.PictureBox picCriteria 
         Height          =   1515
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   9495
         TabIndex        =   9
         Top             =   180
         Width           =   9555
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "WBS"
            Height          =   315
            Index           =   9
            Left            =   5445
            TabIndex        =   132
            ToolTipText     =   "Search through the Webster Translation"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "DBY"
            Height          =   315
            Index           =   8
            Left            =   1650
            TabIndex        =   131
            ToolTipText     =   "Search through the Darby Translation"
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton cmdCustom 
            Caption         =   "&Search All Books"
            Height          =   315
            Left            =   0
            TabIndex        =   129
            ToolTipText     =   "Click to select books in New Covenant to Search"
            Top             =   1200
            Width           =   9495
         End
         Begin VB.ComboBox cboSearch 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   127
            ToolTipText     =   "Previous search list"
            Top             =   840
            Width           =   5835
         End
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "ASV"
            Height          =   315
            Index           =   7
            Left            =   900
            TabIndex        =   126
            ToolTipText     =   "Search through the American Standard Version"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "WEB"
            Height          =   315
            Index           =   6
            Left            =   6180
            TabIndex        =   125
            ToolTipText     =   "Search through the World English Bible"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "MKJV"
            Height          =   315
            Index           =   5
            Left            =   3120
            TabIndex        =   124
            ToolTipText     =   "Search through the Modern King James Version"
            Top             =   0
            Width           =   675
         End
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "Greek"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   123
            ToolTipText     =   "Search through the Greek version"
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "KJV"
            Height          =   315
            Index           =   1
            Left            =   2385
            TabIndex        =   122
            ToolTipText     =   "Search through the King James Version"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "YLT"
            Height          =   315
            Index           =   2
            Left            =   4710
            TabIndex        =   121
            ToolTipText     =   "Search through the Young's Literal Translation"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "RSV"
            Height          =   315
            Index           =   3
            Left            =   3915
            TabIndex        =   120
            ToolTipText     =   "Search through the Revised Standard Version"
            Top             =   0
            Width           =   675
         End
         Begin VB.OptionButton optSearchGrk 
            Caption         =   "MPV"
            Height          =   315
            Index           =   4
            Left            =   6915
            TabIndex        =   119
            ToolTipText     =   "Search through your own personal version"
            Top             =   0
            Width           =   675
         End
         Begin VB.CheckBox chkIngnoreCase 
            Caption         =   "&Ignore character case"
            Height          =   195
            Left            =   7620
            TabIndex        =   2
            ToolTipText     =   "Ignore character case during search"
            Top             =   900
            Width           =   1875
         End
         Begin VB.CommandButton cmdClose 
            Cancel          =   -1  'True
            Caption         =   "Close"
            Height          =   375
            Left            =   8640
            TabIndex        =   4
            ToolTipText     =   "Cancel search"
            Top             =   480
            Width           =   795
         End
         Begin VB.CommandButton cmdGo 
            Caption         =   "Search"
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
            Left            =   7620
            TabIndex        =   3
            ToolTipText     =   "Begin search"
            Top             =   480
            Width           =   795
         End
         Begin VB.ComboBox cboOptions 
            Height          =   315
            ItemData        =   "frmSearch.frx":1166
            Left            =   7620
            List            =   "frmSearch.frx":1176
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Select search options"
            Top             =   60
            Width           =   1815
         End
         Begin VB.TextBox txtSearch 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   0
            ToolTipText     =   "Enter text to search for"
            Top             =   540
            Width           =   7335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Previous Searches:"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   128
            Top             =   900
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Entering a number will be assumed to be a Strong's Reference Number)"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   116
            Top             =   360
            Width           =   5085
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text to find:"
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
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   360
            Width           =   1050
         End
      End
   End
   Begin VB.PictureBox picAlpha 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   60
      ScaleHeight     =   615
      ScaleWidth      =   10095
      TabIndex        =   11
      Top             =   1980
      Width           =   10095
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
         Left            =   9225
         TabIndex        =   115
         Top             =   300
         Width           =   90
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
         Left            =   5610
         TabIndex        =   114
         Top             =   300
         Width           =   105
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
         Left            =   7290
         TabIndex        =   113
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
         Left            =   6960
         TabIndex        =   112
         Top             =   300
         Width           =   45
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
         Left            =   8550
         TabIndex        =   111
         Top             =   300
         Width           =   105
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
         Left            =   8385
         TabIndex        =   110
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
         Left            =   8280
         TabIndex        =   109
         Top             =   300
         Width           =   45
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
         Left            =   8115
         TabIndex        =   108
         Top             =   300
         Width           =   105
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
         Left            =   7995
         TabIndex        =   107
         Top             =   300
         Width           =   60
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
         Left            =   7815
         TabIndex        =   106
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
         Left            =   7635
         TabIndex        =   105
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
         Left            =   7455
         TabIndex        =   104
         Top             =   300
         Width           =   120
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
         Left            =   5250
         TabIndex        =   103
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
         Left            =   7065
         TabIndex        =   102
         Top             =   300
         Width           =   165
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
         Left            =   5430
         TabIndex        =   101
         Top             =   300
         Width           =   120
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
         Left            =   6795
         TabIndex        =   100
         Top             =   300
         Width           =   105
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
         Left            =   6690
         TabIndex        =   99
         Top             =   300
         Width           =   45
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
         Left            =   6585
         TabIndex        =   98
         Top             =   300
         Width           =   45
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
         Left            =   6420
         TabIndex        =   97
         Top             =   300
         Width           =   105
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
         Left            =   6240
         TabIndex        =   96
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
         Left            =   6135
         TabIndex        =   95
         Top             =   300
         Width           =   45
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
         Left            =   5955
         TabIndex        =   94
         Top             =   300
         Width           =   120
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
         Left            =   5775
         TabIndex        =   93
         Top             =   300
         Width           =   120
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
         Left            =   9060
         TabIndex        =   92
         Top             =   300
         Width           =   105
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
         Left            =   8910
         TabIndex        =   91
         Top             =   300
         Width           =   90
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
         Left            =   8715
         TabIndex        =   90
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
         Left            =   5070
         TabIndex        =   89
         Top             =   300
         Width           =   120
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
         Left            =   570
         TabIndex        =   88
         Top             =   300
         Width           =   135
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
         Left            =   2610
         TabIndex        =   87
         Top             =   300
         Width           =   150
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
         Left            =   2220
         TabIndex        =   86
         Top             =   300
         Width           =   105
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
         Left            =   4245
         TabIndex        =   85
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
         Left            =   4035
         TabIndex        =   84
         Top             =   300
         Width           =   150
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
         Left            =   3840
         TabIndex        =   83
         Top             =   300
         Width           =   135
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
         Left            =   3645
         TabIndex        =   82
         Top             =   300
         Width           =   135
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
         Left            =   3435
         TabIndex        =   81
         Top             =   300
         Width           =   150
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
         Left            =   3225
         TabIndex        =   80
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
         Left            =   3030
         TabIndex        =   79
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
         Left            =   2820
         TabIndex        =   78
         Top             =   300
         Width           =   150
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
         Left            =   180
         TabIndex        =   77
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
         Left            =   2385
         TabIndex        =   76
         Top             =   300
         Width           =   165
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
         Left            =   375
         TabIndex        =   75
         Top             =   300
         Width           =   135
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
         Left            =   2040
         TabIndex        =   74
         Top             =   300
         Width           =   120
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
         Left            =   1875
         TabIndex        =   73
         Top             =   300
         Width           =   105
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
         Left            =   1770
         TabIndex        =   72
         Top             =   300
         Width           =   45
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
         Left            =   1560
         TabIndex        =   71
         Top             =   300
         Width           =   150
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
         Left            =   1350
         TabIndex        =   70
         Top             =   300
         Width           =   150
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
         Left            =   1170
         TabIndex        =   69
         Top             =   300
         Width           =   120
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
         Left            =   975
         TabIndex        =   68
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
         Left            =   765
         TabIndex        =   67
         Top             =   300
         Width           =   150
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
         Left            =   4875
         TabIndex        =   66
         Top             =   300
         Width           =   135
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
         Left            =   4695
         TabIndex        =   65
         Top             =   300
         Width           =   120
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
         Left            =   4440
         TabIndex        =   64
         Top             =   300
         Width           =   195
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
         Left            =   9225
         TabIndex        =   63
         Top             =   0
         Width           =   75
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
         Left            =   5610
         TabIndex        =   62
         Top             =   0
         Width           =   105
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
         Left            =   7290
         TabIndex        =   61
         Top             =   0
         Width           =   105
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
         Left            =   6960
         TabIndex        =   60
         Top             =   0
         Width           =   90
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
         Left            =   8550
         TabIndex        =   59
         Top             =   0
         Width           =   135
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
         Left            =   8385
         TabIndex        =   58
         Top             =   0
         Width           =   120
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
         Left            =   8280
         TabIndex        =   57
         Top             =   0
         Width           =   75
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
         Left            =   8115
         TabIndex        =   56
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
         Left            =   7995
         TabIndex        =   55
         Top             =   0
         Width           =   120
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
         Left            =   7815
         TabIndex        =   54
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
         Left            =   7635
         TabIndex        =   53
         Top             =   0
         Width           =   90
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
         Left            =   7455
         TabIndex        =   52
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
         Left            =   5250
         TabIndex        =   51
         Top             =   0
         Width           =   120
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
         Left            =   7065
         TabIndex        =   50
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
         Left            =   5430
         TabIndex        =   49
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
         Left            =   6795
         TabIndex        =   48
         Top             =   0
         Width           =   105
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
         Left            =   6690
         TabIndex        =   47
         Top             =   0
         Width           =   105
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
         Left            =   6585
         TabIndex        =   46
         Top             =   0
         Width           =   60
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
         Left            =   6420
         TabIndex        =   45
         Top             =   0
         Width           =   120
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
         Left            =   6240
         TabIndex        =   44
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
         Left            =   6135
         TabIndex        =   43
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
         Left            =   5955
         TabIndex        =   42
         Top             =   0
         Width           =   75
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
         Left            =   5775
         TabIndex        =   41
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
         Left            =   9060
         TabIndex        =   40
         Top             =   0
         Width           =   135
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
         Left            =   8910
         TabIndex        =   39
         Top             =   0
         Width           =   75
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
         Left            =   8715
         TabIndex        =   38
         Top             =   0
         Width           =   120
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
         Left            =   5070
         TabIndex        =   37
         Top             =   0
         Width           =   120
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
         Left            =   570
         TabIndex        =   36
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
         Left            =   2610
         TabIndex        =   35
         Top             =   0
         Width           =   150
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
         Left            =   2220
         TabIndex        =   34
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
         Left            =   4245
         TabIndex        =   33
         Top             =   0
         Width           =   75
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
         Left            =   4035
         TabIndex        =   32
         Top             =   0
         Width           =   135
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
         Left            =   3840
         TabIndex        =   31
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
         Left            =   3645
         TabIndex        =   30
         Top             =   0
         Width           =   120
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
         Left            =   3435
         TabIndex        =   29
         Top             =   0
         Width           =   105
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
         Left            =   3225
         TabIndex        =   28
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
         Left            =   3030
         TabIndex        =   27
         Top             =   0
         Width           =   150
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
         Left            =   2820
         TabIndex        =   26
         Top             =   0
         Width           =   135
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
         Left            =   180
         TabIndex        =   25
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
         Left            =   2385
         TabIndex        =   24
         Top             =   0
         Width           =   180
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
         Left            =   375
         TabIndex        =   23
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
         Left            =   2040
         TabIndex        =   22
         Top             =   0
         Width           =   135
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
         Left            =   1875
         TabIndex        =   21
         Top             =   0
         Width           =   135
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
         Left            =   1770
         TabIndex        =   20
         Top             =   0
         Width           =   45
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
         Left            =   1560
         TabIndex        =   19
         Top             =   0
         Width           =   120
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
         Left            =   1350
         TabIndex        =   18
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
         Left            =   1170
         TabIndex        =   17
         Top             =   0
         Width           =   150
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
         Left            =   975
         TabIndex        =   16
         Top             =   0
         Width           =   120
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
         Left            =   765
         TabIndex        =   15
         Top             =   0
         Width           =   120
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
         Left            =   4875
         TabIndex        =   14
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
         Left            =   4695
         TabIndex        =   13
         Top             =   0
         Width           =   135
      End
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
         Left            =   4440
         TabIndex        =   12
         Top             =   0
         Width           =   165
      End
   End
   Begin VB.Label lblSizer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9960
      TabIndex        =   130
      Top             =   5580
      Width           =   240
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Cap As String = "Search Bibles"

Private OldListHandler As clscboFullDrop   'handle fulldrop on book list

Private MeWidth As Long
Private Meheight As Long
Private MatchCol As Collection
Private MyToolTips As clsToolTip   'MyToolTips can be any name useful to you

Private ListOffset As Long         'adjustment in case of a Strong # reference

Private IsLoading As Boolean

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Initialize the form
'*******************************************************************************
Private Sub Form_Load()
  Dim Idx As Integer, Pbk As Long, Pchp As Long, Pvrs As Long, I As Long
  Dim S As String, Ary() As String, tAry() As String, bAry() As String
  Dim HasData As Boolean
  
  IsLoading = True
  Screen.MousePointer = vbHourglass
  DoEvents
  Set OldListHandler = New clscboFullDrop
'''*** Comment out following 1 line if you are debugging this form code,
'''*** as otherwise the WndProc handler will hange the VB IDE if you try to STOP
'''*** (not step through, this is OK) this code with this code active
  OldListHandler.hwnd = Me.cboSearch.hwnd
  MeWidth = Me.Width
  Meheight = Me.Height
  Me.picCriteria.BorderStyle = 0
  Me.picCriteria.BackColor = cMedium
  For Idx = 0 To 9
    Me.optSearchGrk(Idx).BackColor = cMedium
  Next Idx
  Me.frameCriteria.BackColor = cMedium
  Me.frameResults.BackColor = cMedium
  Me.chkIngnoreCase.BackColor = cMedium
  Me.cmdHowTo.BackColor = cMedium
  
  Me.chkIngnoreCase.Value = CLng(GetSetting(App.Title, "Settings", "IgnoreSearchCase", "0"))
  
  Me.optSearchGrk(BblVersion + 1).Value = True
  Me.cboOptions.ListIndex = CLng(GetSetting(App.Title, "Settings", "SearchOption", "0"))
'
' get previous searches
'
  With colSrch
    Do While .Count
      .Remove 1
    Loop
  End With
  
  With Me.cboSearch
    .Clear
    S = GetSetting(App.Title, "Settings", "PreviousSearches", vbNullString)
    Me.cboSearch.Enabled = CBool(Len(S))
    If Len(S) <> 0 Then
      Ary = Split(S, ",")
      For Idx = 0 To UBound(Ary)
        .AddItem Ary(Idx)
        colSrch.Add Ary(Idx), Ary(Idx)
      Next Idx
      .ListIndex = Idx - 1
      LastSearch = .Text
    End If
  End With
  
  Me.Top = frmGrkXlate.Top
  Me.Height = (Me.Height - Me.ScaleHeight) + Me.frameResults.Top
  Me.Left = Screen.Width - MeWidth
  Me.cmdGo.Enabled = False
  Me.optSearchGrk(4).Enabled = PersonalVersion
  Me.cboOptions.ListIndex = SearchOption
  Set MatchCol = New Collection
  Set MyToolTips = New clsToolTip     'declare object
  With MyToolTips
    .Create Me              'create object
    .MaxTipWidth = 1440 * 6 'set to 4 inches
    .DelayTime(ttDelayShow) = 20 * 1000 'set to 20 seconds
    .SetFont , 12
    .AddTool Me.lstSearch
    .ToolText(Me.lstSearch) = vbNullString
  End With
  
  For I = 0 To 9
    If Me.optSearchGrk(I).Value Then
      S = UCase$(Me.optSearchGrk(I).Caption)
      Exit For
    End If
  Next I
  If S = "MPV" Then
    Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\" & S & ".txt", ForReading, False)
  Else
    Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\" & S & ".txt", ForReading, False)
  End If
  bAry = Split(ts.ReadAll, vbCrLf)
  ts.Close
  
  Me.Show
  DoEvents
  Me.cmdGo.Enabled = CBool(Len(Trim$(Me.txtSearch.Text)))
  frmGrkXlate.mnuWinSearch.Enabled = True
  frmGrkXlate.CheckWin
  S = GetSetting(App.Title, "Settings", "SearchBooks", String$(27, "1"))
  For Idx = 1 To 27
    SearchBooks(Idx) = CBool(Mid$(S, Idx, 1))
  Next Idx
  Call CheckSearch
  SearchOpen = True
  
  If Len(Me.txtSearch.Text) <> 0 Then
    If Fso.FileExists(AddSlash(AppPath) & "DB\LastSearch.txt") Then
      Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\LastSearch.txt", ForReading, False)
      On Error Resume Next
      Ary = Split(ts.ReadAll, vbCrLf)
      HasData = Err.Number = 0
      On Error GoTo 0
      ts.Close
    End If
    
    If HasData Then
      With Me.lstSearch
        For Idx = 0 To UBound(Ary)
          S = Ary(Idx)
          If Len(S) <> 0 Then
            Pbk = CLng(Left$(S, 2))         'grab book
            Pchp = CLng(Mid$(S, 3, 2))      'grab chapter
            Pvrs = CLng(Mid$(S, 5, 2))      'grab verse
            If SearchBooks(Pbk) Then
              I = FindExactMatch(frmGrkXlate.lstGrk, S) 'find the greek text
              If I <> -1 Then
                MatchCol.Add S
                tAry = Split(Books(Pbk), ",")   'get book Dbase entry
                .AddItem tAry(3) & " " & CStr(Pchp) & ":" & CStr(Pvrs) & " " & Mid$(bAry(I), 8)
              End If
            End If
          End If
        Next Idx
        .ListIndex = CLng(GetSetting(App.Title, "Settings", "SearchIndex", "-1"))
      End With
'
' report number of found items
'
      If CBool(Me.lstSearch.ListCount) Then
        Me.Caption = Cap & " - Matches found: " & CStr(Me.lstSearch.ListCount - ListOffset)
'
' set up list display in appropriate font, and adjust size of form to accomodate results
'
        If Me.optSearchGrk(0).Value Then                'if GREEK
          Me.lstSearch.FontName = "Symbol"
          MyToolTips.SetFont "Symbol", 12, True
        Else
          Me.lstSearch.FontName = "MS Sans Serif"       'else English
          MyToolTips.SetFont "MS Sans Serif", 12, True
        End If
        If Me.Height <= Me.Height - Me.ScaleHeight + Me.frameResults.Top Then
          Me.Height = frmGrkXlate.Height
        End If
      End If
    End If
  End If
'
' If a search is forced, then select the list
'
  If Len(ForceSearch) <> 0 Then
    Me.txtSearch.Text = ForceSearch
    ForceSearch = vbNullString
    Me.cboOptions.ListIndex = 2
    Me.cmdGo.Value = True
  End If
  
  Screen.MousePointer = vbDefault
  IsLoading = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Update the form background tilingx
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, frmGrkXlate.picTile(Background)
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize the form. Keep the width consistent
'*******************************************************************************
Private Sub Form_Resize()
  Dim Sz As Long
  If Me.WindowState = vbMinimized Then
    frmGrkXlate.ZOrder 0
    frmGrkXlate.SetFocus
    Exit Sub         'ignore this if minimizing
  End If
  Sz = Me.Height - Me.ScaleHeight + Me.frameResults.Top
  Me.Width = MeWidth
  If Me.WindowState = vbNormal Then
    If Me.Height < Sz Then Me.Height = Meheight
    If CBool(Me.lstSearch.ListCount) = True And Me.Height <= Sz Then Me.Height = Meheight
  End If
  If CBool(Me.lstSearch.ListCount) = True Then
    Me.frameResults.Height = Me.ScaleHeight - Me.frameResults.Top - (Me.cmdCopy.Height + 120)
    Me.lstSearch.Height = Me.frameResults.Height - Me.lstSearch.Top
  End If
  Me.cmdCopy.Left = Me.frameResults.Left + Me.frameResults.Width - Me.cmdCopy.Width - 90
  Me.cmdCopy.Top = Me.ScaleHeight - Me.cmdCopy.Height - 60
  Me.cmdCopy.Visible = CBool(Me.lstSearch.ListCount)
  Me.cmdHowTo.Left = Me.ScaleWidth - Me.cmdHowTo.Width
  Me.lblSizer.Visible = CBool(Me.lstSearch.ListCount)
  Me.lblSizer.Top = Me.ScaleHeight - Me.lblSizer.Height
  Me.lblSizer.Left = Me.ScaleWidth - Me.lblSizer.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim Idx As Long, I As Long, J As Long
  Dim S As String
  
  frmGrkXlate.mnuWinSearch.Enabled = False
  frmGrkXlate.CheckWin
'
' save off the books-to-search selections
'
  S = vbNullString
  For Idx = 1 To 27
    If SearchBooks(Idx) Then
      S = S & "1"               'book is searched
    Else
      S = S & "0"               'book is not searched
    End If
  Next Idx
  Call SaveSetting(App.Title, "Settings", "SearchBooks", S)
'
' save off old search options
'
  With Me.cboSearch
    I = .ListCount - 15
    If I < 1 Then I = 1
    S = vbNullString
    For Idx = I - 1 To .ListCount - 1
      S = S & "," & .List(Idx)
    Next Idx
    SaveSetting App.Title, "Settings", "PreviousSearches", Mid$(S, 2)
  End With
  
  Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\LastSearch.txt", ForWriting, True)
  With MatchCol
    Do While .Count
      ts.WriteLine .Item(1)
      .Remove 1
    Loop
  End With
  Call SaveSetting(App.Title, "Settings", "SearchIndex", CStr(Me.lstSearch.ListIndex))
  ts.Close
  SearchOpen = False
  
  Set MyToolTips = Nothing                  'destroy created objects
  Set MatchCol = Nothing
  Set OldListHandler = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : cboSearch_Click
' Purpose           : Select an old Search
'*******************************************************************************
Private Sub cboSearch_Click()
  Dim S As String
  Static FiddleList As Boolean
  
  If FiddleList Then Exit Sub
  FiddleList = True
  With Me.cboSearch
    If .ListIndex + 1 <> .ListCount Then
      S = .List(.ListIndex)
      .RemoveItem .ListIndex
      .AddItem S
      .ListIndex = .ListCount - 1
    End If
    Me.txtSearch.Text = .List(.ListIndex)
  End With
  FiddleList = False
End Sub

'*******************************************************************************
' Subroutine Name   : chkIngnoreCase_Click
' Purpose           : Flag to indicate ignoring character case
'*******************************************************************************
Private Sub chkIngnoreCase_Click()
  SaveSetting App.Title, "Settings", "IgnoreSearchCase", CStr(Me.chkIngnoreCase.Value)
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCustom_Click
' Purpose           : Custom search through selected books
'*******************************************************************************
Private Sub cmdCustom_Click()
  frmSearchBooks.Show vbModal, Me
  Call CheckSearch
End Sub

'*******************************************************************************
' Subroutine Name   : cmdHowTo_Click
' Purpose           : Show how to use the reference table
'*******************************************************************************
Private Sub cmdHowTo_Click()
  MessageBox Me, "Clicking a letter (in either row) will append it to the text field.", _
                 vbOKOnly Or vbInformation, "How to Use the Greek Alphabet Table"
End Sub

'*******************************************************************************
' Subroutine Name   : CheckSearch
' Purpose           : Save off the books to search in the option button text
'*******************************************************************************
Public Sub CheckSearch()
  Dim Idx As Integer, Cnt As Long
  Dim S As String, Ary() As String, T As String
  
  Cnt = 0
  T = vbNullString
  For Idx = 1 To 27
    If SearchBooks(Idx) Then
      Cnt = Cnt + 1
      Ary = Split(Books(Idx), ",")
      S = Ary(1)
      Select Case Left$(S, 1)
        Case "1" To "3"
          S = Left$(S, 2) & LCase$(Right$(S, 1))
        Case Else
          S = Left$(S, 1) & LCase$(Right$(S, 2))
      End Select
      T = T & "," & S
    End If
  Next Idx
  If Cnt = 27 Then
    T = "&Search All Books"
  ElseIf Cnt = 4 Then
      If SearchBooks(1) = True And SearchBooks(2) = True And SearchBooks(3) = True And SearchBooks(4) = True Then
        T = "&Search: Gospels"
      Else
        T = "&Search: " & Mid$(T, 2)
      End If
  Else
    T = "&Search: " & Mid$(T, 2)
  End If
  Me.cmdCustom.Caption = T
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopy_Click
' Purpose           : Copy list to the clipboard
'*******************************************************************************
Private Sub cmdCopy_Click()
  Dim Idx As Long
  Dim S As String
  
  With Me.lstSearch
    For Idx = 0 To .ListCount - 1
      S = S & vbCrLf & .List(Idx)
    Next Idx
  End With
  Clipboard.Clear
  Clipboard.SetText Mid$(S, 2)
  Me.txtSearch.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : Label2_Click
' Purpose           : User clicked a reference table letter
'*******************************************************************************
Private Sub Label2_Click(Index As Integer)
  Dim T As String
  
  With Me.txtSearch
    T = .Text & Me.Label2(Index).Caption
    .Text = T
    .SelStart = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lstSearch_MouseMove
' Purpose           : update tooltip as required
'*******************************************************************************
Private Sub lstSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim S As String, T As String
  
  S = Trim$(GetStringFromMouseMove(Me.lstSearch, X, Y))  'get line data
  T = MyToolTips.ToolText(Me.lstSearch)           'get current tooltip (grabs max 80 chars)
  If Len(T) = 0 And Len(S) <> 0 Then
    MyToolTips.ToolText(Me.lstSearch) = S
  ElseIf Len(S) = 0 And Len(T) <> 0 Then
    MyToolTips.ToolText(Me.lstSearch) = S
  ElseIf Left$(S, Len(T)) <> T Then
    MyToolTips.ToolText(Me.lstSearch) = S
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : picAlpha_Paint
' Purpose           : Update the background for the options buttons
'                   : For some stupid reason, XP with XP buttons will
'                   : not display them property when they are set on a frame
'*******************************************************************************
Private Sub picAlpha_Paint()
  PaintTilePicBackground Me.picAlpha, frmGrkXlate.picTile(Background)
End Sub

'*******************************************************************************
' Subroutine Name   : cboOptions_Click
' Purpose           : Update selection, but keep focus on the textbox
'*******************************************************************************
Private Sub cboOptions_Click()
  SearchOption = Me.cboOptions.ListIndex
  SaveSetting App.Title, "Settings", "SearchOption", CStr(Me.cboOptions.ListIndex)
  On Error Resume Next
  Me.txtSearch.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : lstSearch_Click
' Purpose           : Update main form with selected verse
'*******************************************************************************
Private Sub lstSearch_Click()
  Dim S As String, Ary() As String
  
  If IsLoading Then Exit Sub
  With Me.lstSearch
    If .ListIndex = 0 And ListOffset <> 0 Then Exit Sub
    Me.Caption = Cap & " - Match " & CStr(.ListIndex + 1 - ListOffset) & " of " & CStr(.ListCount - ListOffset)
    S = MatchCol(.ListIndex + 1 - ListOffset)
  End With
  Bk = CLng(Left$(S, 2))          'grab book
  Ary = Split(Books(Bk), ",")
  ChpCnt = CLng(Ary(4))           'grab chapter count
  Chp = CLng(Mid$(S, 3, 2))       'grab chapter
  Vrs = CLng(Right$(S, 2))        'grab verse
  Call frmGrkXlate.GetVerseCount  'grab verse count
  Call frmGrkXlate.UpdateVerse
End Sub

'*******************************************************************************
' Subroutine Name   : optSearchGrk_Click
' Purpose           : Use Symbol font for editing Greek text
'*******************************************************************************
Private Sub optSearchGrk_Click(Index As Integer)
  If Index = 0 Then
    Me.txtSearch.FontName = "Symbol"
  Else
    Me.txtSearch.FontName = "MS Sans Serif"
  End If
  Me.cboSearch.FontName = Me.txtSearch.FontName
  On Error Resume Next
  Me.txtSearch.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : txtSearch_Change
' Purpose           : Enable accept button only if there is data to search for
'*******************************************************************************
Private Sub txtSearch_Change()
  Me.cmdGo.Enabled = CBool(Len(Trim$(Me.txtSearch.Text)))
End Sub

'*******************************************************************************
' Subroutine Name   : txtSearch_GotFocus
' Purpose           : Select all text in the textbox on focus
'*******************************************************************************
Private Sub txtSearch_GotFocus()
  With Me.txtSearch
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClose_Click
' Purpose           : Nothing to do
'*******************************************************************************
Private Sub cmdClose_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdGo_Click
' Purpose           : Perform a search
'*******************************************************************************
Private Sub cmdGo_Click()
  Dim bAry() As String, S As String, T As String, Phrase As String, pAry() As String
  Dim tAry() As String, TT As String
  Dim Idx As Long, I As Long, Pbk As Long, Pchp As Long, Pvrs As Long, StrNum As Long
  Dim K As Long, Vrsn As Long, II As Long, JJ As Long, CompareCase As Long
  Dim Match As Boolean, Bol As Boolean
  
  If Me.chkIngnoreCase.Value = vbChecked Then
    CompareCase = vbTextCompare
  Else
    CompareCase = vbBinaryCompare
  End If
  LastSearch = Trim$(Me.txtSearch.Text)           'save as last search text
  
  Me.lstSearch.Clear
  DoEvents
  On Error Resume Next
  colSrch.Add LastSearch, LastSearch
  If Err.Number = 0 Then
    Me.cboSearch.Enabled = True
    Me.cboSearch.AddItem LastSearch
    Me.cboSearch.ListIndex = Me.cboSearch.ListCount - 1
  Else
    With colSrch
      For Idx = 1 To .Count
        If StrComp(.Item(Idx), LastSearch, vbTextCompare) = 0 Then
          .Remove Idx
          .Add LastSearch, LastSearch
          Exit For
        End If
      Next Idx
    End With
  End If
  On Error GoTo 0
  Me.cboSearch.Enabled = True
  
  If IsNumeric(LastSearch) Then                   'a Strong #?
    On Error Resume Next
    StrNum = CLng(LastSearch)                     'grab it
    If Err.Number <> 0 Then StrNum = 0            'bogus (may contain special chars)
    On Error GoTo 0
  End If
  Me.Enabled = False
  frmGrkXlate.Enabled = False
  Screen.MousePointer = vbHourglass               'show that we are busy
  Me.Caption = Cap                                'reset title in case it has been changes
  DoEvents
  If Me.optSearchGrk(BblVersion + 1).Value Then
    bAry = Bible                                  'current bible
  ElseIf Me.optSearchGrk(0).Value Then
    bAry = Grk                                    'Greek bible
  Else
    For Vrsn = 1 To 9
      If Me.optSearchGrk(Vrsn).Value Then Exit For
    Next Vrsn
    Select Case Vrsn                              'other bible
      Case 1
        S = "KJV"
      Case 2
        S = "YLT"
      Case 3
        S = "RSV"
      Case 4
        S = "MPV"
      Case 5
        S = "MKJV"
      Case 6
        S = "WEB"
      Case 7
        S = "ASV"
      Case 8
        S = "DBY"
      Case 9
        S = "WBS"
    End Select
    
    If S = "MPV" Then
      Set ts = Fso.OpenTextFile(AddSlash(AppPath) & "DB\" & S & ".txt", ForReading, False)
    Else
      Set ts = Fso.OpenTextFile(AddSlash(App.Path) & "DB\" & S & ".txt", ForReading, False)
    End If
    bAry = Split(ts.ReadAll, vbCrLf)
    ts.Close
  End If
'
' build a reference list
'
  Me.lstRef.Clear
  For Idx = 0 To UBound(bAry) - 1
    Me.lstRef.AddItem Left$(bAry(Idx), 6)
  Next Idx
  
  Phrase = LastSearch                           'text to scan for
'
' if a Strong #, find it in the Definition Reference list to get the DefRef index for it
'
  ListOffset = 0
  If StrNum <> 0 Then
    For Idx = 1 To UBound(DefRef)
      S = DefRef(Idx)
      If Len(S) Then
        tAry = Split(S, vbTab)
        If Phrase = tAry(5) Then                'Strong # match?
          Phrase = CStr(Idx)                    'yes, so now we have the correct reference
          tAry = Split(WordRef(StrNum), vbTab)
          Me.lstSearch.AddItem String$(16, 32) & "<Strong # " & str(StrNum) & " General Definition: " & tAry(2) & ">"
          ListOffset = 1
          Exit For
        End If
      End If
    Next Idx
    
    If Idx <= UBound(DefRef) Then
      For Idx = 0 To UBound(GrkBBL)             'scan through the Greek index table
        S = GrkBBL(Idx)                         'grab a list of reference from the map
        If Len(S) <> 0 Then
          tAry = Split(S, " ")
            For I = 1 To UBound(tAry)
              If tAry(I) = Phrase Then          'check for a matched index number
                K = FindExactMatch(Me.lstRef, tAry(0))  'find the BKCHVS entry
                If K <> -1 Then                         'if we found it
                  S = bAry(K)                           'grab it
                  If Len(S) > 7 Then                    'contains data?
                    Pbk = CLng(Left$(S, 2))         'grab book
                    Pchp = CLng(Mid$(S, 3, 2))      'grab chapter
                    Pvrs = CLng(Mid$(S, 5, 2))      'grab verse
                    If SearchBooks(Pbk) Then
                      MatchCol.Add Left$(S, 6)        'add header to collection if so
                      tAry = Split(Books(Pbk), ",")   'get book Dbase entry
                      Me.lstSearch.AddItem tAry(3) & " " & CStr(Pchp) & ":" & CStr(Pvrs) & " " & Mid$(S, 8)
                    End If
                    Exit For                        'done with verse entry
                  End If
                End If
              End If
            Next I
        End If
      Next Idx
    End If
  Else
'
' Searhc for words, not numbers here
' First, strip non-alphanumerics from text
'
    For Idx = 1 To Len(Phrase)
      Select Case Mid$(Phrase, Idx, 1)
        Case "A" To "Z", "a" To "z", "0" To "9"
        Case Else
          Mid$(Phrase, Idx, 1) = " "
      End Select
    Next Idx
'
' now strip double-spaces
'
    Idx = InStr(1, Phrase, "  ")
    Do While Idx <> 0
      Phrase = Left$(Phrase, Idx) & LTrim$(Mid$(Phrase, Idx + 2))
      Idx = InStr(Idx + 1, Phrase, "  ")
    Loop
'
' if not a phrase search, then break up the prhase intoa list of individual words
'
    Select Case SearchOption
      Case 2
        Phrase = " " & Phrase & " "               'else pad phrase with spaces
      Case 3
      Case Else
        pAry = Split(Phrase, " ")
        For Idx = 0 To UBound(pAry)
          pAry(Idx) = " " & pAry(Idx) & " "       'pad each word with spaces
        Next Idx
    End Select
'
' clear match lists
'
    Me.lstSearch.Clear
    With MatchCol
      Do While .Count
        .Remove 1
      Loop
    End With
'
' scan for matches
'
    For Idx = 0 To UBound(bAry)
      T = Mid$(bAry(Idx), 8)                      'get text of verse
      TT = " " & T                                'grab a copy with a prepended space
      If Len(T) <> 0 Then                         'any data?
        For I = 1 To Len(T)
          Select Case Mid$(T, I, 1)
            Case "A" To "Z", "a" To "z", "0" To "9"
            Case Else
              Mid$(T, I, 1) = " "
          End Select
        Next I
'
' now pad with a space
'
        T = " " & T & " "
'
' now check for matches
'
        Select Case SearchOption
          Case 0  'match ANY
            Match = False               'initiallly assume failure
            For I = 0 To UBound(pAry)
              II = InStr(1, T, pAry(I), CompareCase)
              If II <> 0 Then
                Match = True            'match found
                Do While II <> 0        'and make TT match UPPERCASE
                  JJ = InStr(II + 1, T, " ") 'find trailing space
                  Mid$(TT, II + 1, JJ - II - 1) = UCase$(Mid$(TT, II + 1, JJ - II - 1))
                  II = InStr(JJ, T, pAry(I)) 'march ALL matches in line
                Loop
              End If
            Next I
          Case 1  'match ALL
            Match = True                'assume success
            For I = 0 To UBound(pAry)
              II = InStr(1, T, pAry(I), CompareCase)
              If II = 0 Then            'failure
                Match = False           'mark so
              Else
                Do While II <> 0              'but make TT match UPPERCASE
                  JJ = InStr(II + 1, T, " ")  'find trailing space
                  Mid$(TT, II + 1, JJ - II - 1) = UCase$(Mid$(TT, II + 1, JJ - II - 1))
                  II = InStr(JJ, T, pAry(I))  'mark ALL matches in line
                Loop
              End If
            Next I
          Case 2, 3 'match PHRASE
            Match = False
            II = InStr(1, T, Phrase, CompareCase) 'accept only if the phrase is found
            If II <> 0 Then
              Match = True
              Do While II            'and make TT match UPPERCASE
                Mid$(TT, II, Len(Phrase)) = UCase$(Mid$(TT, II, Len(Phrase)))
                II = InStr(II + Len(Phrase) - 1, T, Phrase)
              Loop
            End If
        End Select
      End If
'
' if a match is found...
'
      If Match Then
        S = Left$(bAry(Idx), 6)         'grab line header
        Pbk = CLng(Left$(S, 2))         'grab book
        Pchp = CLng(Mid$(S, 3, 2))      'grab chapter
        Pvrs = CLng(Right$(S, 2))       'grab verse
        If SearchBooks(Pbk) Then
          MatchCol.Add S                  'add header to collection
          tAry = Split(Books(Pbk), ",")   'get book titled
          Me.lstSearch.AddItem tAry(3) & " " & CStr(Pchp) & ":" & CStr(Pvrs) & " " & TT
        End If
        Match = False
      End If
    Next Idx
  End If
'
' all done with search, so show no longer busy
'
  Me.Enabled = True
  frmGrkXlate.Enabled = True
  Screen.MousePointer = vbDefault
'
' report number of found items
'
  Me.Caption = Cap & " - Matches found: " & CStr(Me.lstSearch.ListCount - ListOffset)
'
' set up list display in appropriate font, and adjust size of form to accomodate results
'
  If CBool(Me.lstSearch.ListCount) Then
    If Me.optSearchGrk(0).Value Then                'if GREEK
      Me.lstSearch.FontName = "Symbol"
      MyToolTips.SetFont "Symbol", 12, True
    Else
      Me.lstSearch.FontName = "MS Sans Serif"       'else English
      MyToolTips.SetFont "MS Sans Serif", 12, True
    End If
    If Me.Height <= Me.Height - Me.ScaleHeight + Me.frameResults.Top Then
'      Me.Height = Meheight
      Me.Height = frmGrkXlate.Height
    End If
  Else          'if no matches found
    MessageBox Me, "No Matches found", vbOKOnly Or vbExclamation, "No Matches"
  End If
  Me.txtSearch.SetFocus
End Sub
