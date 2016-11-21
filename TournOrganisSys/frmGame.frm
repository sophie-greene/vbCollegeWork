VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00E0E0E0&
   Caption         =   "TOS"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6240
      TabIndex        =   80
      Top             =   10200
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   6000
      TabIndex        =   79
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtMatch 
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame fraGame 
      BackColor       =   &H00E0E0E0&
      Height          =   7575
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   1440
         TabIndex        =   76
         Top             =   7095
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   3360
         TabIndex        =   78
         Top             =   7095
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   72
         Top             =   6720
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   3360
         TabIndex        =   74
         Top             =   6720
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   1440
         TabIndex        =   68
         Top             =   6345
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   3360
         TabIndex        =   70
         Top             =   6345
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   1440
         TabIndex        =   64
         Top             =   5970
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   3360
         TabIndex        =   66
         Top             =   5970
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   1440
         TabIndex        =   60
         Top             =   5595
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   3360
         TabIndex        =   62
         Top             =   5595
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   56
         Top             =   5220
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   3360
         TabIndex        =   58
         Top             =   5220
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   52
         Top             =   4845
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   3360
         TabIndex        =   54
         Top             =   4845
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   1440
         TabIndex        =   48
         Top             =   4470
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   3360
         TabIndex        =   50
         Top             =   4470
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   1440
         TabIndex        =   44
         Top             =   4095
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3360
         TabIndex        =   46
         Top             =   4095
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   40
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   3360
         TabIndex        =   42
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   36
         Top             =   3345
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3360
         TabIndex        =   38
         Top             =   3345
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   32
         Top             =   2970
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3360
         TabIndex        =   34
         Top             =   2970
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   28
         Top             =   2595
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3360
         TabIndex        =   30
         Top             =   2595
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   24
         Top             =   2220
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   26
         Top             =   2220
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   20
         Top             =   1845
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   22
         Top             =   1845
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   16
         Top             =   1470
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   18
         Top             =   1470
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Top             =   1095
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   14
         Top             =   1095
         Width           =   615
      End
      Begin VB.TextBox txtYResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3360
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtGResult 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow Students"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   9
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Green Students"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   75
         Top             =   7095
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   2160
         TabIndex        =   77
         Top             =   7095
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   240
         TabIndex        =   71
         Top             =   6720
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   2160
         TabIndex        =   73
         Top             =   6720
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   240
         TabIndex        =   67
         Top             =   6345
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   2160
         TabIndex        =   69
         Top             =   6345
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   240
         TabIndex        =   63
         Top             =   5970
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   2160
         TabIndex        =   65
         Top             =   5970
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   240
         TabIndex        =   59
         Top             =   5595
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   2160
         TabIndex        =   61
         Top             =   5595
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   240
         TabIndex        =   55
         Top             =   5220
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   2160
         TabIndex        =   57
         Top             =   5220
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   51
         Top             =   4845
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   2160
         TabIndex        =   53
         Top             =   4845
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   240
         TabIndex        =   47
         Top             =   4470
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   2160
         TabIndex        =   49
         Top             =   4470
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   43
         Top             =   4095
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   2160
         TabIndex        =   45
         Top             =   4095
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   39
         Top             =   3720
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   41
         Top             =   3720
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   35
         Top             =   3345
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   37
         Top             =   3345
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   2970
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   33
         Top             =   2970
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   27
         Top             =   2595
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   29
         Top             =   2595
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   23
         Top             =   2220
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   25
         Top             =   2220
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1845
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   21
         Top             =   1845
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1470
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   17
         Top             =   1470
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1095
         Width           =   1200
      End
      Begin VB.Label lblYStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   13
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lblGStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1200
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter number of &Matches:"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerate_Click()
    Dim g() As game
    Dim intCount As Integer
    Dim count_Records1 As Integer
    Dim intTemp As Integer
    'clear the lables and textboxes
    Call clearControlArrays
    gadoCommand.CommandText = "UPDATE Students  Set Played" & gstrGameName & "=false"
    Set gadoRecordSet = gadoCommand.Execute
    gadoCommand.CommandText = "UPDATE Students  Set " & gstrGameName & "Oponent=0"
    Set gadoRecordSet = gadoCommand.Execute
    'initialize the count
    intCount = 0
    
    'validate the number of matches
    If Val(txtMatch.Text) > 18 Or Val(txtMatch.Text) < 1 Or txtMatch.Text = "" Then
        MsgBox "Invalid number of matches, please enter a number between 1 and 18"
        txtMatch.Text = ""
        txtMatch.SetFocus
        Exit Sub
    End If
    'allocate game records dynamically
    ReDim g(0 To Val(txtMatch.Text) - 1)
    
    Do While (intCount <> Val(txtMatch.Text))
    'select a student from the green house
        Do
            'generate a random student number
            Randomize (Timer)
            intTemp = Int(Rnd(Timer) * 36) + 1
            'read the student record from the access database
            gadoCommand.CommandText = "SELECT * FROM Students WHERE House='Green'" & _
                "AND Played" & gstrGameName & "=False AND studentID = " & intTemp
            Set gadoRecordSet = gadoCommand.Execute
        Loop While (gadoRecordSet.EOF = True)
        'put student info in the array
        copyRecord g(intCount).strGstudent, gadoRecordSet
        g(intCount).strGstudent.boolPlaySnap = True
        gadoCommand.CommandText = "UPDATE Students  Set Played" & gstrGameName & "=true WHERE StudentID=" _
            & g(intCount).strGstudent.intStudentID
        Set gadoRecordSet = gadoCommand.Execute
        lblGStudent(intCount) = g(intCount).strGstudent.strStudentName
        'yellow student
        Do
            Randomize (Timer)
            intTemp = Int(Rnd(Timer) * 36) + 1
            gadoCommand.CommandText = "SELECT * FROM Students WHERE House='Yellow' AND Played" & gstrGameName & "=False AND " & gstrGameName & "Oponent <>" _
            & g(intCount).strGstudent.intStudentID & " AND Class <> '" & g(intCount).strGstudent.strClass & "' AND studentID = " & intTemp
            Set gadoRecordSet = gadoCommand.Execute
        Loop While (gadoRecordSet.EOF = True)
        
        copyRecord g(intCount).strYstudent, gadoRecordSet
        g(intCount).strGstudent.boolPlaySnap = True
        gadoCommand.CommandText = "UPDATE Students  Set Played" & gstrGameName & "=true WHERE StudentID=" & g(intCount).strYstudent.intStudentID
        lblYStudent(intCount) = g(intCount).strYstudent.strStudentName
       
        gadoCommand.CommandText = "UPDATE Students  Set Played" & gstrGameName & "=true WHERE StudentID=" & gadoRecordSet.Fields(0)
        Set gadoRecordSet = gadoCommand.Execute
        gadoCommand.CommandText = "UPDATE Students  Set " & gstrGameName & "Oponent=" & _
            g(intCount).strGstudent.intStudentID & " WHERE StudentID=" _
            & g(intCount).strYstudent.intStudentID
        Set gadoRecordSet = gadoCommand.Execute
        
        gadoCommand.CommandText = "UPDATE Students  Set " & gstrGameName & "Oponent=" & _
            g(intCount).strYstudent.intStudentID & " WHERE StudentID=" _
            & g(intCount).strGstudent.intStudentID
        Set gadoRecordSet = gadoCommand.Execute
     'gadoCommand.CommandText = "SELECT COUNT(*) as recordcount FROM Students Where House='Yellow' and Played" & gstrGameName & "=False"
     'Set gadoRecordSet = gadoCommand.Execute
     'count_Records1 = gadoRecordSet("recordcount")
     intCount = intCount + 1
     Loop
     Beep
End Sub


Private Sub Command1_Click()
Unload Me
frmIntro.Show
End Sub

Private Sub Form_Load()
   
    'fix form position relative to the screen
    Left = Screen.Width / 2 - Width / 2
    Top = Screen.Height / 2 - Height / 2 - 500
    Caption = gstrGameName
    fraGame.Caption = gstrGameName
End Sub



Private Sub txtGResult_Change(Index As Integer)
    If txtGResult(Index) = "" Then Exit Sub
    If Val(txtGResult(Index).Text) < 0 Or Val(txtGResult(Index).Text) > 3 Then
        MsgBox "Invalid Result, Enter a value between 0 and 3"
        txtGResult(Index) = ""
        txtGResult(Index).SetFocus
        Exit Sub
    End If
    If Val(txtYResult(Index).Text) + Val(txtGResult(Index).Text) <> 3 _
            And txtYResult(Index).Text <> "" Then
        MsgBox "The results of both students shoould add to 3"
        txtGResult(Index) = ""
        txtGResult(Index).SetFocus
    End If
End Sub

Private Sub txtYResult_Change(Index As Integer)
    If txtYResult(Index) = "" Then Exit Sub
    If Val(txtYResult(Index).Text) < 0 Or Val(txtYResult(Index).Text) > 3 Then
        MsgBox "Invalid Result, Enter a value between 0 and 3"
        txtYResult(Index) = ""
        txtYResult(Index).SetFocus
        Exit Sub
    End If
    If Val(txtYResult(Index).Text) + Val(txtGResult(Index).Text) <> 3 _
            And txtGResult(Index).Text <> "" Then
        MsgBox "The results of both students should add to 3"
        txtYResult(Index) = ""
        txtYResult(Index).SetFocus
    End If
End Sub
Private Sub clearControlArrays()
    Dim intX As Integer
    For intX = 0 To 17
        lblGStudent(intX).Caption = ""
    Next intX
    For intX = 0 To 17
        lblYStudent(intX).Caption = ""
    Next intX
    For intX = 0 To 17
        txtGResult(intX).Text = ""
    Next intX
    For intX = 0 To 17
        txtYResult(intX).Text = ""
    Next intX
End Sub


