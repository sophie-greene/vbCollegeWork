VERSION 5.00
Begin VB.Form frmManual 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Manual"
   ClientHeight    =   8235
   ClientLeft      =   2010
   ClientTop       =   915
   ClientWidth     =   14220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Change Number of Matches"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   159
      Top             =   7680
      Width           =   3735
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   203
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.Frame fraScrabble 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Scrabble"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   10605
      TabIndex        =   119
      Top             =   240
      Width           =   3495
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   158
         Top             =   3240
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   157
         Top             =   3240
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   153
         Top             =   720
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   152
         Top             =   720
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   151
         Top             =   1080
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   150
         Top             =   1080
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   149
         Top             =   1440
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   148
         Top             =   1440
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   147
         Top             =   1800
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   146
         Top             =   1800
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   2160
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   144
         Top             =   2160
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   143
         Top             =   2520
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   2520
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   141
         Top             =   2880
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   140
         Top             =   2880
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   139
         Top             =   3600
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   138
         Top             =   3600
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   3960
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   136
         Top             =   3960
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   4320
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Top             =   4320
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   133
         Top             =   4680
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   132
         Top             =   4680
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   5040
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   130
         Top             =   5040
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   129
         Top             =   5400
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   128
         Top             =   5400
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   127
         Top             =   5760
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   126
         Top             =   5760
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   6120
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   6120
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Top             =   6480
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   6480
         Width           =   1600
      End
      Begin VB.ComboBox cmbGScrabbleStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   6840
         Width           =   1600
      End
      Begin VB.ComboBox cmbYScrabbleStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   6840
         Width           =   1600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
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
         Index           =   7
         Left            =   120
         TabIndex        =   155
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
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
         Index           =   6
         Left            =   1680
         TabIndex        =   154
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.Frame fraSpillikins 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Spillikins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   7110
      TabIndex        =   80
      Top             =   240
      Width           =   3495
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   720
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   720
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   1080
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   1080
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   1440
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   1440
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   1800
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   1800
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   2160
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   2160
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   2520
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   2520
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   2880
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   2880
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   3240
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   3240
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   3600
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   3600
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   3960
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   3960
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   4320
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   4320
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   4680
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   4680
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   5040
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   5040
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   5400
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   5400
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   5760
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   5760
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   6120
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   6120
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   6480
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   6480
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSpillikinsStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   6840
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSpillikinsStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   6840
         Width           =   1600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
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
         Index           =   11
         Left            =   240
         TabIndex        =   118
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
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
         Index           =   10
         Left            =   1680
         TabIndex        =   117
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.Frame fraSnap 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Snap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   120
      TabIndex        =   41
      Top             =   240
      Width           =   3495
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   720
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   720
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   1080
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   1080
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1440
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1440
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   1800
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1800
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2160
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   2160
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   2520
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   2520
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   2880
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   2880
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   3240
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   3240
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   3600
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   3600
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   3960
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   3960
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   4320
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   4320
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   4680
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   4680
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   5040
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   5040
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   5400
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   5400
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   5760
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   5760
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   6120
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   6120
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   6480
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   6480
         Width           =   1600
      End
      Begin VB.ComboBox cmbGSnapStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   6840
         Width           =   1600
      End
      Begin VB.ComboBox cmbYSnapStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   6840
         Width           =   1600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
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
         Index           =   2
         Left            =   1680
         TabIndex        =   79
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
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
         Index           =   0
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.Frame fraCribbage 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cribbage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   3615
      TabIndex        =   2
      Top             =   240
      Width           =   3495
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   6840
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   6840
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   6480
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   6480
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   6120
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   6120
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   5760
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   5760
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   5400
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   5400
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   5040
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   5040
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   4680
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4680
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   4320
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   4320
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3960
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3960
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3600
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3600
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3240
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3240
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2880
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2880
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2520
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2520
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2160
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2160
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1800
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1600
      End
      Begin VB.ComboBox cmbGCribbageStudent 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1600
      End
      Begin VB.ComboBox cmbYCribbageStudent 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "frmManual.frx":0000
         Left            =   1725
         List            =   "frmManual.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   1600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
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
         Index           =   15
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
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
         Index           =   14
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Go to the Main Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   4455
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   12323
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
End
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'************************** Tournament Organising System************************
'**********************************frmManual Code*******************************
'*******************************Programer: S. Saqfelhait************************
'***********************************Date:07/04/2007*****************************
'*******************************************************************************
'this form is loaded from frmIntro "select game generation method and number of
'matches form, it is also loaded from random form when Edit button is clicked
Option Explicit

'***************** Green Cribbage***********************************************
'generate the list of players that are available to play
'*******************************************************************************
Private Sub cmbGCribbageStudent_DropDown(Index As Integer)

    Call Manual(cmbGCribbageStudent(Index), cmbYCribbageStudent(Index).Text, Index, 2)
    
End Sub

'***************** Green Cribbage***********************************************
'update the game array and database once a player is selected
'*******************************************************************************
Private Sub cmbGCribbageStudent_Click(Index As Integer)

    If cmbGCribbageStudent(Index).ListIndex <> -1 Then
        gadoCommand.CommandText = "UPDATE Students  Set PlayedCribbage=true " _
        & "where studentName='" & cmbGCribbageStudent(Index).Text & "'"
        Set gadoRecordSet = gadoCommand.Execute
        strGame(Index, 2) = cmbGCribbageStudent(Index).Text
        cmbYCribbageStudent(Index).Enabled = True
    End If
    
End Sub
'***************** Green Scrabble**********************************************
'generate the list of players that are available to play
'*******************************************************************************
Private Sub cmbGScrabbleStudent_DropDown(Index As Integer)

    Call Manual(cmbGScrabbleStudent(Index), cmbYScrabbleStudent(Index).Text, Index, 6)
    
End Sub

'***************** Green Scrabble**********************************************
'update the game array and database once a player is selected
'*******************************************************************************
Private Sub cmbGScrabbleStudent_Click(Index As Integer)

    If cmbGScrabbleStudent(Index).ListIndex <> -1 Then
        gadoCommand.CommandText = "UPDATE Students  Set PlayedScrabble=true " _
        & " where studentName='" & cmbGScrabbleStudent(Index).Text & "'"
        Set gadoRecordSet = gadoCommand.Execute
        strGame(Index, 6) = cmbGScrabbleStudent(Index).Text
        End If
        
End Sub

'***************** Green Snap**************************************************
'generate the list of players that are available to play
'*******************************************************************************
Private Sub cmbGSnapStudent_DropDown(Index As Integer)

    Call Manual(cmbGSnapStudent(Index), cmbYSnapStudent(Index).Text, Index, 0)
    
End Sub

'***************** Green Snap***************************************************
'update the game array and database once a player is selected
Private Sub cmbGSnapStudent_Click(Index As Integer)

  '  If cmbGSnapStudent(Index).ListIndex <> -1 Then
        gadoCommand.CommandText = "UPDATE Students  Set PlayedSnap=true " _
        & " where studentName='" & cmbGSnapStudent(Index).Text & "'"
        Set gadoRecordSet = gadoCommand.Execute
        strGame(Index, 0) = cmbGSnapStudent(Index).Text
   ' End If
    
End Sub

'***************** Green Spillikins*********************************************
'generate the list of players that are available to play
'*******************************************************************************
Private Sub cmbGSpillikinsStudent_DropDown(Index As Integer)

    Call Manual(cmbGSpillikinsStudent(Index), _
        cmbYSpillikinsStudent(Index).Text, Index, 4)
    
End Sub

'***************** Green Spillikins**********************************************
'update the game array and database once a player is selected
'*******************************************************************************
Private Sub cmbGSpillikinsStudent_Click(Index As Integer)
    If cmbGSpillikinsStudent(Index).ListIndex <> -1 Then
        gadoCommand.CommandText = "UPDATE Students  Set PlayedSpillikins=true " _
        & " where studentName='" & cmbGSpillikinsStudent(Index).Text & "'"
        Set gadoRecordSet = gadoCommand.Execute
        strGame(Index, 4) = cmbGSpillikinsStudent(Index).Text
    End If
End Sub

'***************** Yellow Cribbage************************************************
'generate the list of players that are available to play
'*******************************************************************************
Private Sub cmbYCribbageStudent_DropDown(Index As Integer)
     Manual cmbYCribbageStudent(Index), cmbGCribbageStudent(Index).Text, Index, 3
End Sub

'***************** Yellow Cribbage*********************************************
'update the game array and database once a player is selected
'*******************************************************************************
Private Sub cmbYCribbageStudent_Click(Index As Integer)

    If cmbYCribbageStudent(Index).ListIndex <> -1 Then
        gadoCommand.CommandText = "UPDATE Students  Set PlayedCribbage=True " _
        & "where studentName='" & cmbYCribbageStudent(Index).Text & "'"
        Set gadoRecordSet = gadoCommand.Execute
        strGame(Index, 3) = cmbYCribbageStudent(Index).Text
    End If
    
End Sub

'***************** Yellow Scrabble************************************************
'generate the list of players that are available to play
'*******************************************************************************
Private Sub cmbYScrabbleStudent_DropDown(Index As Integer)
 Manual cmbYScrabbleStudent(Index), cmbGScrabbleStudent(Index).Text, Index, 7
    
End Sub

'***************** Yellow Scrabble***********************************************
'update the game array and database once a player is selected
'*******************************************************************************
Private Sub cmbYScrabbleStudent_Click(Index As Integer)

    If cmbYScrabbleStudent(Index).ListIndex <> -1 Then
        gadoCommand.CommandText = "UPDATE Students  Set PlayedScrabble=True " _
        & "where studentName='" & cmbYScrabbleStudent(Index).Text & "'"
        Set gadoRecordSet = gadoCommand.Execute
        strGame(Index, 7) = cmbYScrabbleStudent(Index).Text
    End If
    
End Sub

'***************** Yellow Spillikins********************************************
'generate the list of players that are available to play
'*******************************************************************************
Private Sub cmbYSpillikinsStudent_DropDown(Index As Integer)

 Manual cmbYSpillikinsStudent(Index), cmbGSpillikinsStudent(Index).Text, Index, 5
    
End Sub

'***************** Yellow Spillikins********************************************
'update the game array and database once a player is selected
'*******************************************************************************
Private Sub cmbYSpillikinsStudent_Click(Index As Integer)

    If cmbYSpillikinsStudent(Index).ListIndex <> -1 Then
        gadoCommand.CommandText = "UPDATE Students  Set PlayedSpillikins=True " _
        & "where studentName='" & cmbYSpillikinsStudent(Index).Text & "'"
        Set gadoRecordSet = gadoCommand.Execute
        strGame(Index, 5) = cmbYSpillikinsStudent(Index).Text
    End If
    
End Sub

'***************** Yellow Snap*************************************************
'generate the list of players that are available to play
'*******************************************************************************
Private Sub cmbYSnapStudent_DropDown(Index As Integer)

   Manual cmbYSnapStudent(Index), cmbGSnapStudent(Index).Text, Index, 1
    
End Sub

'***************** Yellow Snap**************************************************
'update the game array and database once a player is selected
'*******************************************************************************
Private Sub cmbYSnapStudent_Click(Index As Integer)

    If cmbYSnapStudent(Index).ListIndex <> -1 Then
        gadoCommand.CommandText = "UPDATE Students  Set PlayedSnap=True " _
        & "where studentName='" & cmbYSnapStudent(Index).Text & "'"
        Set gadoRecordSet = gadoCommand.Execute
        strGame(Index, 1) = cmbYSnapStudent(Index).Text
    End If
    
End Sub
'*****************************************************************************
'subroutine will be executed when the Go to main window command button is clicked
'*****************************************************************************
Private Sub cmdMain_Click()
    frmMain.Show
    Unload Me
End Sub

'*****************************************************************************
'subroutine will be executed when the Change Number of matches command button is clicked
'*****************************************************************************
Private Sub cmdChange_Click()
    frmIntro.Show
    Unload Me
End Sub

'*****************************************************************************
'subroutine will be executed when the Exit command button is clicked
'*****************************************************************************
Private Sub cmdExit_Click()
       
  'define a variable to capture the msgbox response
    Dim varResponce As Variant
    varResponce = myMsgBox("Are You Sure you Want to Exit", "YesNo", "Attention")
    'exit program if response if Yes
    If varResponce = vbYes Then
        DB_Disconnect
        End
    End If
End Sub

'*****************************************************************************
'subroutine will be executed when the Save command button is clicked
'*****************************************************************************
Private Sub cmdSave_Click()
    'define a variable "flag" which is set to true if any
    'enabled combo box is empty and initialise it to false
    Dim emptyResults As Boolean
    emptyResults = False
    
    Dim intX As Integer
    'check if any enabled combo box is empty
    'Snap
    For intX = 0 To gintSnapMatch - 1
        If cmbGSnapStudent(intX).Text = "" Then
            cmbGSnapStudent(intX).BackColor = &HC0E0FF
            emptyResults = True
        End If
        
        If cmbYSnapStudent(intX).Text = "" Then
            cmbYSnapStudent(intX).BackColor = &HC0E0FF
            emptyResults = True
        End If
    Next intX
    'Cribbage
    For intX = 0 To gintCribbageMatch - 1
        If cmbGCribbageStudent(intX).Text = "" Then
            cmbGCribbageStudent(intX).BackColor = &HC0E0FF
            emptyResults = True
        End If
        
        If cmbYCribbageStudent(intX).Text = "" Then
            cmbYCribbageStudent(intX).BackColor = &HC0E0FF
            emptyResults = True
        End If
    Next intX
    'Spillikins
    For intX = 0 To gintSpillikinsMatch - 1
        If cmbGSpillikinsStudent(intX).Text = "" Then
            cmbGSpillikinsStudent(intX).BackColor = &HC0E0FF
            emptyResults = True
        End If
        
        If cmbYSpillikinsStudent(intX).Text = "" Then
            cmbYSpillikinsStudent(intX).BackColor = &HC0E0FF
            emptyResults = True
        End If
    Next intX
    'Scrabble
    For intX = 0 To gintScrabbleMatch - 1
        If cmbGScrabbleStudent(intX).Text = "" Then
            cmbGScrabbleStudent(intX).BackColor = &HC0E0FF
            emptyResults = True
        End If
        
        If cmbYScrabbleStudent(intX).Text = "" Then
            cmbYScrabbleStudent(intX).BackColor = &HC0E0FF
            emptyResults = True
        End If
    Next intX
    
    If emptyResults = True Then
    myMsgBox "Please fill in the highlighted game boxes", "Ok", "Attention"
    Exit Sub
    End If
    
    'save and load the random form with the selected games
    gboolRandom = False
    frmRandom.Show
    Unload Me
End Sub
'*****************************************************************************
'subroutine will be executed when the form is loaded
'*****************************************************************************
Private Sub Form_Load()
  Dim intY As Integer
    If gboolEdit = True Then
    
        For intY = 0 To gintSnapMatch - 1
            cmbGSnapStudent(intY).AddItem strGame(intY, 0)
            cmbGSnapStudent(intY).ListIndex = 0
            cmbYSnapStudent(intY).AddItem strGame(intY, 1)
            cmbYSnapStudent(intY).ListIndex = 0
        Next intY
        
        For intY = 0 To gintCribbageMatch - 1
            cmbGCribbageStudent(intY).AddItem strGame(intY, 2)
            cmbGCribbageStudent(intY).ListIndex = 0
            cmbYCribbageStudent(intY).AddItem strGame(intY, 3)
            cmbYCribbageStudent(intY).ListIndex = 0
        Next intY
        
        For intY = 0 To gintSpillikinsMatch - 1
            cmbGSpillikinsStudent(intY).AddItem strGame(intY, 4)
            cmbGSpillikinsStudent(intY).ListIndex = 0
            cmbYSpillikinsStudent(intY).AddItem strGame(intY, 5)
            cmbYSpillikinsStudent(intY).ListIndex = 0
        Next intY
        
        For intY = 0 To gintScrabbleMatch - 1
            cmbGScrabbleStudent(intY).AddItem strGame(intY, 6)
            cmbGScrabbleStudent(intY).ListIndex = 0
            cmbYScrabbleStudent(intY).AddItem strGame(intY, 7)
            cmbYScrabbleStudent(intY).ListIndex = 0
        Next intY
        
    Else
        'clear games
        clearGameArray
        'clear the game schedule in the database
        clearDatabase
        'read students list from the database
        getStudentArray strGstudent, "Green"
        getStudentArray strYstudent(), "Yellow"
    End If
  
    'enable the combo boxes based on the number of matches
    For intY = gintSnapMatch To 17
        cmbGSnapStudent(intY).Enabled = False
        cmbYSnapStudent(intY).Enabled = False
         
    Next intY
    For intY = gintCribbageMatch To 17
        cmbGCribbageStudent(intY).Enabled = False
        cmbYCribbageStudent(intY).Enabled = False
        
    Next intY
    For intY = gintSpillikinsMatch To 17
        cmbGSpillikinsStudent(intY).Enabled = False
        cmbYSpillikinsStudent(intY).Enabled = False
      
    Next intY
    For intY = gintScrabbleMatch To 17
        cmbGScrabbleStudent(intY).Enabled = False
        cmbYScrabbleStudent(intY).Enabled = False
       
    Next intY

End Sub

