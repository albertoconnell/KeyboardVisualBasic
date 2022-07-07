VERSION 5.00
Begin VB.Form KeyboardCheck 
   BorderStyle     =   0  'None
   Caption         =   "Your Keyboard"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   17775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox T 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Text            =   "test"
      Top             =   5400
      Width           =   17295
   End
   Begin VB.CommandButton Exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Frame KeyboardFrame 
      Caption         =   "YOUR KEYBOARD"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   17295
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "WIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "CTRL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "FN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "ALT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "SHIFT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3480
         Width           =   2415
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   191
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "ENTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   """"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   222
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   ";"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   186
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "\"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   220
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   221
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "["
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   219
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "BACKSPACE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   187
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   189
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   48
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   57
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   56
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   55
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   54
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   53
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   52
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   51
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   50
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   49
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "ESC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   27
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   192
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "TAB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "CAPS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   20
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "CTRL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "SHIFT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "ALT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   18
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   80
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   79
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   73
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   85
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   89
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   84
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   82
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   69
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   87
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   81
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   76
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   75
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   74
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   72
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   190
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   ","
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   188
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   77
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   78
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   66
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   71
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   70
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   68
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   83
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   65
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   90
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   88
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   86
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Height          =   735
         Index           =   32
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4200
         Width           =   5175
      End
      Begin VB.CommandButton K 
         BackColor       =   &H8000000A&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   67
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3480
         Width           =   855
      End
   End
End
Attribute VB_Name = "KeyboardCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim caps, numlock As Boolean
Private Sub Exit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    KeyboardCheck.Top = Screen.Height / 6
    KeyboardCheck.Left = Screen.Width / 6
    
End Sub

Private Sub T_KeyDown(KeyCode As Integer, Shift As Integer)
    'MsgBox KeyCode
    'Space Key
    If KeyCode = 32 Then
        K(32).BackColor = vbGreen
    End If
    'A Key
    If KeyCode = 65 Then
        K(65).BackColor = vbGreen
    End If
    'B Key
    If KeyCode = 66 Then
        K(66).BackColor = vbGreen
    End If
    'C Key
    If KeyCode = 67 Then
        K(67).BackColor = vbGreen
    End If
    'D Key
    If KeyCode = 68 Then
        K(68).BackColor = vbGreen
    End If
    'E Key
    If KeyCode = 69 Then
        K(69).BackColor = vbGreen
    End If
    'F Key
    If KeyCode = 70 Then
        K(70).BackColor = vbGreen
    End If
    'G Key
    If KeyCode = 71 Then
        K(71).BackColor = vbGreen
    End If
    'H Key
    If KeyCode = 72 Then
        K(72).BackColor = vbGreen
    End If
    'I Key
    If KeyCode = 73 Then
        K(73).BackColor = vbGreen
    End If
    'J Key
    If KeyCode = 74 Then
        K(74).BackColor = vbGreen
    End If
    'K Key
    If KeyCode = 75 Then
        K(75).BackColor = vbGreen
    End If
    'L Key
    If KeyCode = 76 Then
        K(76).BackColor = vbGreen
    End If
    'M Key
    If KeyCode = 77 Then
        K(77).BackColor = vbGreen
    End If
    'N Key
    If KeyCode = 78 Then
        K(78).BackColor = vbGreen
    End If
    'O Key
    If KeyCode = 79 Then
        K(79).BackColor = vbGreen
    End If
    'P Key
    If KeyCode = 80 Then
        K(80).BackColor = vbGreen
    End If
    'Q Key
    If KeyCode = 81 Then
        K(81).BackColor = vbGreen
    End If
    'R Key
    If KeyCode = 82 Then
        K(82).BackColor = vbGreen
    End If
    'S Key
    If KeyCode = 83 Then
        K(83).BackColor = vbGreen
    End If
    'T Key
    If KeyCode = 84 Then
        K(84).BackColor = vbGreen
    End If
    'U Key
    If KeyCode = 85 Then
        K(85).BackColor = vbGreen
    End If
    'V Key
    If KeyCode = 86 Then
        K(86).BackColor = vbGreen
    End If
    'W Key
    If KeyCode = 87 Then
        K(87).BackColor = vbGreen
    End If
    'X Key
    If KeyCode = 88 Then
        K(88).BackColor = vbGreen
    End If
    'Y Key
    If KeyCode = 89 Then
        K(89).BackColor = vbGreen
    End If
    'Z Key
    If KeyCode = 90 Then
        K(90).BackColor = vbGreen
    End If
    'alt Key
    If KeyCode = 18 Then
        K(18).BackColor = vbGreen
    End If
    'ctrl Key
    If KeyCode = 17 Then
        K(17).BackColor = vbGreen
    End If
    'shift Key
    If KeyCode = 16 Then
        K(16).BackColor = vbGreen
    End If
    'Caps Key
    If KeyCode = 20 Then
        If Not caps Then
            K(20).BackColor = vbGreen
            caps = True
        Else
            K(20).BackColor = &H8000000A
            caps = False
        End If
    End If
    '~ Key
    If KeyCode = 192 Then
        K(192).BackColor = vbGreen
    End If
    'ESC Key
    If KeyCode = 27 Then
        K(27).BackColor = vbGreen
    End If
    '1 Key
    If KeyCode = 49 Then
        K(49).BackColor = vbGreen
    End If
    '2 Key
    If KeyCode = 50 Then
        K(50).BackColor = vbGreen
    End If
    '3 Key
    If KeyCode = 51 Then
        K(51).BackColor = vbGreen
    End If
    '4 Key
    If KeyCode = 52 Then
        K(52).BackColor = vbGreen
    End If
    '5 Key
    If KeyCode = 53 Then
        K(53).BackColor = vbGreen
    End If
    '6 Key
    If KeyCode = 54 Then
        K(54).BackColor = vbGreen
    End If
    '7 Key
    If KeyCode = 55 Then
        K(55).BackColor = vbGreen
    End If
    '8 Key
    If KeyCode = 56 Then
        K(56).BackColor = vbGreen
    End If
    '9 Key
    If KeyCode = 57 Then
        K(57).BackColor = vbGreen
    End If
    '0 Key
    If KeyCode = 48 Then
        K(48).BackColor = vbGreen
    End If
    '- Key
    If KeyCode = 189 Then
        K(189).BackColor = vbGreen
    End If
    '= Key
    If KeyCode = 187 Then
        K(187).BackColor = vbGreen
    End If
    '[ Key
    If KeyCode = 219 Then
        K(219).BackColor = vbGreen
    End If
    '] Key
    If KeyCode = 221 Then
        K(221).BackColor = vbGreen
    End If
    '\ Key
    If KeyCode = 220 Then
        K(220).BackColor = vbGreen
    End If
    ': Key
    If KeyCode = 186 Then
        K(186).BackColor = vbGreen
    End If
    '" Key
    If KeyCode = 222 Then
        K(222).BackColor = vbGreen
    End If
    'Enter Key
    If KeyCode = 13 Then
        K(13).BackColor = vbGreen
    End If
    ', Key
    If KeyCode = 188 Then
        K(188).BackColor = vbGreen
    End If
    '. Key
    If KeyCode = 190 Then
        K(190).BackColor = vbGreen
    End If
    '/ Key
    If KeyCode = 191 Then
        K(191).BackColor = vbGreen
    End If
    'BackSpace Key
    If KeyCode = 8 Then
        K(8).BackColor = vbGreen
    End If
End Sub

Private Sub T_KeyUp(KeyCode As Integer, Shift As Integer)
    'Space Key
    'MsgBox "UP"
    If KeyCode = 32 Then
        K(32).BackColor = &H8000000A
    End If
    'A Key
    If KeyCode = 65 Then
        K(65).BackColor = &H8000000A
    End If
    'B Key
    If KeyCode = 66 Then
        K(66).BackColor = &H8000000A
    End If
    'C Key
    If KeyCode = 67 Then
        K(67).BackColor = &H8000000A
    End If
    'D Key
    If KeyCode = 68 Then
        K(68).BackColor = &H8000000A
    End If
    'E Key
    If KeyCode = 69 Then
        K(69).BackColor = &H8000000A
    End If
    'F Key
    If KeyCode = 70 Then
        K(70).BackColor = &H8000000A
    End If
    'G Key
    If KeyCode = 71 Then
        K(71).BackColor = &H8000000A
    End If
    'H Key
    If KeyCode = 72 Then
        K(72).BackColor = &H8000000A
    End If
    'I Key
    If KeyCode = 73 Then
        K(73).BackColor = &H8000000A
    End If
    'J Key
    If KeyCode = 74 Then
        K(74).BackColor = &H8000000A
    End If
    'K Key
    If KeyCode = 75 Then
        K(75).BackColor = &H8000000A
    End If
    'L Key
    If KeyCode = 76 Then
        K(76).BackColor = &H8000000A
    End If
    'M Key
    If KeyCode = 77 Then
        K(77).BackColor = &H8000000A
    End If
    'N Key
    If KeyCode = 78 Then
        K(78).BackColor = &H8000000A
    End If
    'O Key
    If KeyCode = 79 Then
        K(79).BackColor = &H8000000A
    End If
    'P Key
    If KeyCode = 80 Then
        K(80).BackColor = &H8000000A
    End If
    'Q Key
    If KeyCode = 81 Then
        K(81).BackColor = &H8000000A
    End If
    'R Key
    If KeyCode = 82 Then
        K(82).BackColor = &H8000000A
    End If
    'S Key
    If KeyCode = 83 Then
        K(83).BackColor = &H8000000A
    End If
    'T Key
    If KeyCode = 84 Then
        K(84).BackColor = &H8000000A
    End If
    'U Key
    If KeyCode = 85 Then
        K(85).BackColor = &H8000000A
    End If
    'V Key
    If KeyCode = 86 Then
        K(86).BackColor = &H8000000A
    End If
    'W Key
    If KeyCode = 87 Then
        K(87).BackColor = &H8000000A
    End If
    'X Key
    If KeyCode = 88 Then
        K(88).BackColor = &H8000000A
    End If
    'Y Key
    If KeyCode = 89 Then
        K(89).BackColor = &H8000000A
    End If
    'Z Key
    If KeyCode = 90 Then
        K(90).BackColor = &H8000000A
    End If
    'alt Key
    If KeyCode = 18 Then
        K(18).BackColor = &H8000000A
    End If
    'ctrl Key
    If KeyCode = 17 Then
        K(17).BackColor = &H8000000A
    End If
    'shift Key
    If KeyCode = 16 Then
        K(16).BackColor = &H8000000A
    End If
    '~ Key
    If KeyCode = 192 Then
        K(192).BackColor = &H8000000A
    End If
    'ESC Key
    If KeyCode = 27 Then
        K(27).BackColor = &H8000000A
    End If
    '1 Key
    If KeyCode = 49 Then
        K(49).BackColor = &H8000000A
    End If
    '2 Key
    If KeyCode = 50 Then
        K(50).BackColor = &H8000000A
    End If
    '3 Key
    If KeyCode = 51 Then
        K(51).BackColor = &H8000000A
    End If
    '4 Key
    If KeyCode = 52 Then
        K(52).BackColor = &H8000000A
    End If
    '5 Key
    If KeyCode = 53 Then
        K(53).BackColor = &H8000000A
    End If
    '6 Key
    If KeyCode = 54 Then
        K(54).BackColor = &H8000000A
    End If
    '7 Key
    If KeyCode = 55 Then
        K(55).BackColor = &H8000000A
    End If
    '8 Key
    If KeyCode = 56 Then
        K(56).BackColor = &H8000000A
    End If
    '9 Key
    If KeyCode = 57 Then
        K(57).BackColor = &H8000000A
    End If
    '0 Key
    If KeyCode = 48 Then
        K(48).BackColor = &H8000000A
    End If
    '- Key
    If KeyCode = 189 Then
        K(189).BackColor = &H8000000A
    End If
    '= Key
    If KeyCode = 187 Then
        K(187).BackColor = &H8000000A
    End If
    '[ Key
    If KeyCode = 219 Then
        K(219).BackColor = &H8000000A
    End If
    '] Key
    If KeyCode = 221 Then
        K(221).BackColor = &H8000000A
    End If
    '\ Key
    If KeyCode = 220 Then
        K(220).BackColor = &H8000000A
    End If
    ': Key
    If KeyCode = 186 Then
        K(186).BackColor = &H8000000A
    End If
    '" Key
    If KeyCode = 222 Then
        K(222).BackColor = &H8000000A
    End If
    'Enter Key
    If KeyCode = 13 Then
        K(13).BackColor = &H8000000A
    End If
    ', Key
    If KeyCode = 188 Then
        K(188).BackColor = &H8000000A
    End If
    '. Key
    If KeyCode = 190 Then
        K(190).BackColor = &H8000000A
    End If
    '/ Key
    If KeyCode = 191 Then
        K(191).BackColor = &H8000000A
    End If
    'BackSpace Key
    If KeyCode = 8 Then
        K(8).BackColor = &H8000000A
    End If
End Sub
