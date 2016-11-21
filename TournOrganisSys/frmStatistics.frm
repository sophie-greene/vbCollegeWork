VERSION 5.00
Begin VB.Form frmStatistics 
   BackColor       =   &H8000000C&
   Caption         =   "Students' Statistics"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&OK"
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   481
      Top             =   8280
      Width           =   3975
   End
   Begin VB.CommandButton cmdStudentName 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Student Name"
      Height          =   600
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdSnapCount 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Snap Count"
      Height          =   600
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdCribbageCount 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cribbage Count"
      Height          =   600
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdSpillikinsCount 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Spillikins Count"
      Height          =   600
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdScrabbleCount 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Scrabble Count"
      Height          =   600
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdWonCribbage 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cribbage Won Count"
      Height          =   600
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdWonSpillikins 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Spillikins Won Count"
      Height          =   600
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdWonScrabble 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Scrabble Won Count"
      Height          =   600
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdSnapMax 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Max Snap Score"
      Height          =   600
      Left            =   9855
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdCribbageMax 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Max Cribbage Score"
      Height          =   600
      Left            =   10950
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdSpillikinsMax 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Max Spillikins Score"
      Height          =   600
      Left            =   12045
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdScrabbleMax 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Max Scrabble Score"
      Height          =   600
      Left            =   13140
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdWonSnap 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Snap Won Count"
      Height          =   600
      Left            =   5475
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1100
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   482
      Top             =   8160
      Width           =   14355
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   13140
      TabIndex        =   480
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   12045
      TabIndex        =   479
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   10950
      TabIndex        =   478
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   9855
      TabIndex        =   477
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   8760
      TabIndex        =   476
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   7665
      TabIndex        =   475
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   6570
      TabIndex        =   474
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   5475
      TabIndex        =   473
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   4380
      TabIndex        =   472
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   3285
      TabIndex        =   471
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   2190
      TabIndex        =   470
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   1095
      TabIndex        =   469
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   35
      Left            =   0
      TabIndex        =   468
      Top             =   7950
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   13140
      TabIndex        =   467
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   13140
      TabIndex        =   466
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   13140
      TabIndex        =   465
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   13140
      TabIndex        =   464
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   13140
      TabIndex        =   463
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   13140
      TabIndex        =   462
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   13140
      TabIndex        =   461
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   13140
      TabIndex        =   460
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   13140
      TabIndex        =   459
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   13140
      TabIndex        =   458
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   13140
      TabIndex        =   457
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   13140
      TabIndex        =   456
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   13140
      TabIndex        =   455
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   13140
      TabIndex        =   454
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   13140
      TabIndex        =   453
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   13140
      TabIndex        =   452
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   13140
      TabIndex        =   451
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   13140
      TabIndex        =   450
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   13140
      TabIndex        =   449
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   13140
      TabIndex        =   448
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   13140
      TabIndex        =   447
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   13140
      TabIndex        =   446
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   13140
      TabIndex        =   445
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   13140
      TabIndex        =   444
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   13140
      TabIndex        =   443
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   13140
      TabIndex        =   442
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   13140
      TabIndex        =   441
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   13140
      TabIndex        =   440
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   13140
      TabIndex        =   439
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   13140
      TabIndex        =   438
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   13140
      TabIndex        =   437
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   13140
      TabIndex        =   436
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   13140
      TabIndex        =   435
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   13140
      TabIndex        =   434
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblScrabbleMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   13140
      TabIndex        =   433
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   0
      TabIndex        =   432
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   1095
      TabIndex        =   431
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   2190
      TabIndex        =   430
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   3285
      TabIndex        =   429
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   4380
      TabIndex        =   428
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   5475
      TabIndex        =   427
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   6570
      TabIndex        =   426
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   7665
      TabIndex        =   425
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   8760
      TabIndex        =   424
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   9855
      TabIndex        =   423
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   10950
      TabIndex        =   422
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   34
      Left            =   12045
      TabIndex        =   421
      Top             =   7740
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   0
      TabIndex        =   420
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   1095
      TabIndex        =   419
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   2190
      TabIndex        =   418
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   3285
      TabIndex        =   417
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   4380
      TabIndex        =   416
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   5475
      TabIndex        =   415
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   6570
      TabIndex        =   414
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   7665
      TabIndex        =   413
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   8760
      TabIndex        =   412
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   9855
      TabIndex        =   411
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   10950
      TabIndex        =   410
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   33
      Left            =   12045
      TabIndex        =   409
      Top             =   7530
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   0
      TabIndex        =   408
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   1095
      TabIndex        =   407
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   2190
      TabIndex        =   406
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   3285
      TabIndex        =   405
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   4380
      TabIndex        =   404
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   5475
      TabIndex        =   403
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   6570
      TabIndex        =   402
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   7665
      TabIndex        =   401
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   8760
      TabIndex        =   400
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   9855
      TabIndex        =   399
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   10950
      TabIndex        =   398
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   32
      Left            =   12045
      TabIndex        =   397
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   0
      TabIndex        =   396
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   1095
      TabIndex        =   395
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   2190
      TabIndex        =   394
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   3285
      TabIndex        =   393
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   4380
      TabIndex        =   392
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   5475
      TabIndex        =   391
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   6570
      TabIndex        =   390
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   7665
      TabIndex        =   389
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   8760
      TabIndex        =   388
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   9855
      TabIndex        =   387
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   10950
      TabIndex        =   386
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   31
      Left            =   12045
      TabIndex        =   385
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   0
      TabIndex        =   384
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   1095
      TabIndex        =   383
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   2190
      TabIndex        =   382
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   3285
      TabIndex        =   381
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   4380
      TabIndex        =   380
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   5475
      TabIndex        =   379
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   6570
      TabIndex        =   378
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   7665
      TabIndex        =   377
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   8760
      TabIndex        =   376
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   9855
      TabIndex        =   375
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   10950
      TabIndex        =   374
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   30
      Left            =   12045
      TabIndex        =   373
      Top             =   6900
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   0
      TabIndex        =   372
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   1095
      TabIndex        =   371
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   2190
      TabIndex        =   370
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   3285
      TabIndex        =   369
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   4380
      TabIndex        =   368
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   5475
      TabIndex        =   367
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   6570
      TabIndex        =   366
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   7665
      TabIndex        =   365
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   8760
      TabIndex        =   364
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   9855
      TabIndex        =   363
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   10950
      TabIndex        =   362
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   29
      Left            =   12045
      TabIndex        =   361
      Top             =   6690
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   0
      TabIndex        =   360
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   1095
      TabIndex        =   359
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   2190
      TabIndex        =   358
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   3285
      TabIndex        =   357
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   4380
      TabIndex        =   356
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   5475
      TabIndex        =   355
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   6570
      TabIndex        =   354
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   7665
      TabIndex        =   353
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   8760
      TabIndex        =   352
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   9855
      TabIndex        =   351
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   10950
      TabIndex        =   350
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   28
      Left            =   12045
      TabIndex        =   349
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   0
      TabIndex        =   348
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   1095
      TabIndex        =   347
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   2190
      TabIndex        =   346
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   3285
      TabIndex        =   345
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   4380
      TabIndex        =   344
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   5475
      TabIndex        =   343
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   6570
      TabIndex        =   342
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   7665
      TabIndex        =   341
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   8760
      TabIndex        =   340
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   9855
      TabIndex        =   339
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   10950
      TabIndex        =   338
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   27
      Left            =   12045
      TabIndex        =   337
      Top             =   6270
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   0
      TabIndex        =   336
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   1095
      TabIndex        =   335
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   2190
      TabIndex        =   334
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   3285
      TabIndex        =   333
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   4380
      TabIndex        =   332
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   5475
      TabIndex        =   331
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   6570
      TabIndex        =   330
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   7665
      TabIndex        =   329
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   8760
      TabIndex        =   328
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   9855
      TabIndex        =   327
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   10950
      TabIndex        =   326
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   26
      Left            =   12045
      TabIndex        =   325
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   0
      TabIndex        =   324
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   1095
      TabIndex        =   323
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   2190
      TabIndex        =   322
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   3285
      TabIndex        =   321
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   4380
      TabIndex        =   320
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   5475
      TabIndex        =   319
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   6570
      TabIndex        =   318
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   7665
      TabIndex        =   317
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   8760
      TabIndex        =   316
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   9855
      TabIndex        =   315
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   10950
      TabIndex        =   314
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   25
      Left            =   12045
      TabIndex        =   313
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   0
      TabIndex        =   312
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   1095
      TabIndex        =   311
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   2190
      TabIndex        =   310
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   3285
      TabIndex        =   309
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   4380
      TabIndex        =   308
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   5475
      TabIndex        =   307
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   6570
      TabIndex        =   306
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   7665
      TabIndex        =   305
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   8760
      TabIndex        =   304
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   9855
      TabIndex        =   303
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   10950
      TabIndex        =   302
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   24
      Left            =   12045
      TabIndex        =   301
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   0
      TabIndex        =   300
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   1095
      TabIndex        =   299
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   2190
      TabIndex        =   298
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   3285
      TabIndex        =   297
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   4380
      TabIndex        =   296
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   5475
      TabIndex        =   295
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   6570
      TabIndex        =   294
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   7665
      TabIndex        =   293
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   8760
      TabIndex        =   292
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   9855
      TabIndex        =   291
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   10950
      TabIndex        =   290
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   23
      Left            =   12045
      TabIndex        =   289
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   0
      TabIndex        =   288
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   1095
      TabIndex        =   287
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   2190
      TabIndex        =   286
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   3285
      TabIndex        =   285
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   4380
      TabIndex        =   284
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   5475
      TabIndex        =   283
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   6570
      TabIndex        =   282
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   7665
      TabIndex        =   281
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   8760
      TabIndex        =   280
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   9855
      TabIndex        =   279
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   10950
      TabIndex        =   278
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   22
      Left            =   12045
      TabIndex        =   277
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   0
      TabIndex        =   276
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   1095
      TabIndex        =   275
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   2190
      TabIndex        =   274
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   3285
      TabIndex        =   273
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   4380
      TabIndex        =   272
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   5475
      TabIndex        =   271
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   6570
      TabIndex        =   270
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   7665
      TabIndex        =   269
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   8760
      TabIndex        =   268
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   9855
      TabIndex        =   267
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   10950
      TabIndex        =   266
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   21
      Left            =   12045
      TabIndex        =   265
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   0
      TabIndex        =   264
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   1095
      TabIndex        =   263
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   2190
      TabIndex        =   262
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   3285
      TabIndex        =   261
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   4380
      TabIndex        =   260
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   5475
      TabIndex        =   259
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   6570
      TabIndex        =   258
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   7665
      TabIndex        =   257
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   8760
      TabIndex        =   256
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   9855
      TabIndex        =   255
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   10950
      TabIndex        =   254
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   20
      Left            =   12045
      TabIndex        =   253
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   0
      TabIndex        =   252
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   1095
      TabIndex        =   251
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   2190
      TabIndex        =   250
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   3285
      TabIndex        =   249
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   4380
      TabIndex        =   248
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   5475
      TabIndex        =   247
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   6570
      TabIndex        =   246
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   7665
      TabIndex        =   245
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   8760
      TabIndex        =   244
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   9855
      TabIndex        =   243
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   10950
      TabIndex        =   242
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   19
      Left            =   12045
      TabIndex        =   241
      Top             =   4590
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   0
      TabIndex        =   240
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   1095
      TabIndex        =   239
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   2190
      TabIndex        =   238
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   3285
      TabIndex        =   237
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   4380
      TabIndex        =   236
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   5475
      TabIndex        =   235
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   6570
      TabIndex        =   234
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   7665
      TabIndex        =   233
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   8760
      TabIndex        =   232
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   9855
      TabIndex        =   231
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   10950
      TabIndex        =   230
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   18
      Left            =   12045
      TabIndex        =   229
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   5475
      TabIndex        =   228
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   12045
      TabIndex        =   227
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   10950
      TabIndex        =   226
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   9855
      TabIndex        =   225
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   8760
      TabIndex        =   224
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   7665
      TabIndex        =   223
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   6570
      TabIndex        =   222
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   5475
      TabIndex        =   221
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   4380
      TabIndex        =   220
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   3285
      TabIndex        =   219
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   2190
      TabIndex        =   218
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   1095
      TabIndex        =   217
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   12045
      TabIndex        =   216
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   10950
      TabIndex        =   215
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   9855
      TabIndex        =   214
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   8760
      TabIndex        =   213
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   7665
      TabIndex        =   212
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   6570
      TabIndex        =   211
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   5475
      TabIndex        =   210
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   4380
      TabIndex        =   209
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   3285
      TabIndex        =   208
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   2190
      TabIndex        =   207
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   1095
      TabIndex        =   206
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   12045
      TabIndex        =   205
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   10950
      TabIndex        =   204
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   9855
      TabIndex        =   203
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   8760
      TabIndex        =   202
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   7665
      TabIndex        =   201
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   6570
      TabIndex        =   200
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   5475
      TabIndex        =   199
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   4380
      TabIndex        =   198
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   3285
      TabIndex        =   197
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   2190
      TabIndex        =   196
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   1095
      TabIndex        =   195
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   12045
      TabIndex        =   194
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   10950
      TabIndex        =   193
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   9855
      TabIndex        =   192
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   8760
      TabIndex        =   191
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   7665
      TabIndex        =   190
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   6570
      TabIndex        =   189
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   5475
      TabIndex        =   188
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   4380
      TabIndex        =   187
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   3285
      TabIndex        =   186
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   2190
      TabIndex        =   185
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   1095
      TabIndex        =   184
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   12045
      TabIndex        =   183
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   10950
      TabIndex        =   182
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   9855
      TabIndex        =   181
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   8760
      TabIndex        =   180
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   7665
      TabIndex        =   179
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   6570
      TabIndex        =   178
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   5475
      TabIndex        =   177
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   4380
      TabIndex        =   176
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   3285
      TabIndex        =   175
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   2190
      TabIndex        =   174
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   1095
      TabIndex        =   173
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   12045
      TabIndex        =   172
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   10950
      TabIndex        =   171
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   9855
      TabIndex        =   170
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   8760
      TabIndex        =   169
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   7665
      TabIndex        =   168
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   6570
      TabIndex        =   167
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   5475
      TabIndex        =   166
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   4380
      TabIndex        =   165
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   3285
      TabIndex        =   164
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   2190
      TabIndex        =   163
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   1095
      TabIndex        =   162
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   12045
      TabIndex        =   161
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   10950
      TabIndex        =   160
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   9855
      TabIndex        =   159
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   8760
      TabIndex        =   158
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   7665
      TabIndex        =   157
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   6570
      TabIndex        =   156
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   5475
      TabIndex        =   155
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   4380
      TabIndex        =   154
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   3285
      TabIndex        =   153
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   2190
      TabIndex        =   152
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   1095
      TabIndex        =   151
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   12045
      TabIndex        =   150
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   10950
      TabIndex        =   149
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   9855
      TabIndex        =   148
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   8760
      TabIndex        =   147
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   7665
      TabIndex        =   146
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   6570
      TabIndex        =   145
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   5475
      TabIndex        =   144
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   4380
      TabIndex        =   143
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   3285
      TabIndex        =   142
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   2190
      TabIndex        =   141
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   1095
      TabIndex        =   140
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   12045
      TabIndex        =   139
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   10950
      TabIndex        =   138
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   9855
      TabIndex        =   137
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   8760
      TabIndex        =   136
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   7665
      TabIndex        =   135
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   6570
      TabIndex        =   134
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   5475
      TabIndex        =   133
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   4380
      TabIndex        =   132
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   3285
      TabIndex        =   131
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   2190
      TabIndex        =   130
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   1095
      TabIndex        =   129
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   12045
      TabIndex        =   128
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   10950
      TabIndex        =   127
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   9855
      TabIndex        =   126
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   8760
      TabIndex        =   125
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   7665
      TabIndex        =   124
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   6570
      TabIndex        =   123
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   5475
      TabIndex        =   122
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   4380
      TabIndex        =   121
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   3285
      TabIndex        =   120
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   2190
      TabIndex        =   119
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   1095
      TabIndex        =   118
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   12045
      TabIndex        =   117
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   10950
      TabIndex        =   116
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   9855
      TabIndex        =   115
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   8760
      TabIndex        =   114
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   7665
      TabIndex        =   113
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   6570
      TabIndex        =   112
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   5475
      TabIndex        =   111
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   4380
      TabIndex        =   110
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   3285
      TabIndex        =   109
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   2190
      TabIndex        =   108
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   1095
      TabIndex        =   107
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   12045
      TabIndex        =   106
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   10950
      TabIndex        =   105
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   9855
      TabIndex        =   104
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   8760
      TabIndex        =   103
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   7665
      TabIndex        =   102
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   6570
      TabIndex        =   101
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   5475
      TabIndex        =   100
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   4380
      TabIndex        =   99
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   3285
      TabIndex        =   98
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   2190
      TabIndex        =   97
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   1095
      TabIndex        =   96
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   12045
      TabIndex        =   95
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   10950
      TabIndex        =   94
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   9855
      TabIndex        =   93
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   8760
      TabIndex        =   92
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   7665
      TabIndex        =   91
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   6570
      TabIndex        =   90
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   5475
      TabIndex        =   89
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   4380
      TabIndex        =   88
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   3285
      TabIndex        =   87
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   2190
      TabIndex        =   86
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   1095
      TabIndex        =   85
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   12045
      TabIndex        =   84
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   10950
      TabIndex        =   83
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   9855
      TabIndex        =   82
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   8760
      TabIndex        =   81
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   7665
      TabIndex        =   80
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   6570
      TabIndex        =   79
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   5475
      TabIndex        =   78
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   4380
      TabIndex        =   77
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   3285
      TabIndex        =   76
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   2190
      TabIndex        =   75
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   1095
      TabIndex        =   74
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   12045
      TabIndex        =   73
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   10950
      TabIndex        =   72
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   9855
      TabIndex        =   71
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   8760
      TabIndex        =   70
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   7665
      TabIndex        =   69
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   6570
      TabIndex        =   68
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   5475
      TabIndex        =   67
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   4380
      TabIndex        =   66
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   3285
      TabIndex        =   65
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   2190
      TabIndex        =   64
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   1095
      TabIndex        =   63
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   12045
      TabIndex        =   62
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   10950
      TabIndex        =   61
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   9855
      TabIndex        =   60
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   8760
      TabIndex        =   59
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   7665
      TabIndex        =   58
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   6570
      TabIndex        =   57
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   5475
      TabIndex        =   56
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   4380
      TabIndex        =   55
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   3285
      TabIndex        =   54
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   2190
      TabIndex        =   53
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   1095
      TabIndex        =   52
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   12045
      TabIndex        =   51
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   10950
      TabIndex        =   50
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   9855
      TabIndex        =   49
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   8760
      TabIndex        =   48
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   7665
      TabIndex        =   47
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   6570
      TabIndex        =   46
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblWonSnap 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   5475
      TabIndex        =   45
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   4380
      TabIndex        =   44
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   3285
      TabIndex        =   43
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   2190
      TabIndex        =   42
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   1095
      TabIndex        =   41
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   12045
      TabIndex        =   40
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblCribbageMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   10950
      TabIndex        =   39
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblSnapMax 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   9855
      TabIndex        =   38
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblWonScrabble 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   8760
      TabIndex        =   37
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblWonSpillikins 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   7665
      TabIndex        =   36
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblWonCribbage 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   6570
      TabIndex        =   35
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblScrabbleCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   4380
      TabIndex        =   34
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblSpillikinsCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   3285
      TabIndex        =   33
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblCribbageCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   2190
      TabIndex        =   32
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblSnapCount 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   1095
      TabIndex        =   31
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   17
      Left            =   0
      TabIndex        =   30
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   16
      Left            =   0
      TabIndex        =   29
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   15
      Left            =   0
      TabIndex        =   28
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   14
      Left            =   0
      TabIndex        =   27
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   13
      Left            =   0
      TabIndex        =   26
      Top             =   3330
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   12
      Left            =   0
      TabIndex        =   25
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   11
      Left            =   0
      TabIndex        =   24
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   10
      Left            =   0
      TabIndex        =   23
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   0
      TabIndex        =   22
      Top             =   2490
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   0
      TabIndex        =   21
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   0
      TabIndex        =   20
      Top             =   2070
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   0
      TabIndex        =   19
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   0
      TabIndex        =   18
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   0
      TabIndex        =   17
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   0
      TabIndex        =   16
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   0
      TabIndex        =   15
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
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
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************** Tournament Organising System************************
'**********************************frmHouseSt Code******************************
'****************************Programer: Somoud Saqfelhait***********************
'***********************************Date:07/06/2007*****************************
'*******************************************************************************
'this form will be loaded from the main menu form when House Achievments
'Statistics  command button is clicked

Option Explicit
'a module level variable which holds the sorting (Ascending or Descending)
Dim mstrOrder As String
'a module level variable which holds the last command button clicked
Dim mstrButton As String

'*****************************************************************************
'subroutine will be executed when the Cribbage Count command button is clicked
'*****************************************************************************
Private Sub cmdCribbageCount_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdCribbageCount" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    'order by maximum
    Else
        mstrOrder = "DESC"
    End If
    Call sortStatistics("CribbageMatches")
    mstrButton = "cmdCribbageCount"
End Sub

'*****************************************************************************
'subroutine will be executed when the Cribbage Max Score command button is clicked
'*****************************************************************************
Private Sub cmdCribbageMax_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdCribbageMax" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    'order by maximum
    Else
        mstrOrder = "DESC"
    End If
    Call sortStatistics("CribbageMax")
    mstrButton = "cmdCribbageMax"
End Sub

'*****************************************************************************
'subroutine will be executed when the Ok command button is clicked
'*****************************************************************************
Private Sub cmdOk_Click()
    frmMain.Show
    Unload Me
End Sub

'*****************************************************************************
'subroutine will be executed when the Scrabble Count command button is clicked
'*****************************************************************************
Private Sub cmdScrabbleCount_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdScrabbleCount" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    'order by maximum
    Else
        mstrOrder = "DESC"
    End If
    Call sortStatistics("ScrabbleMatches")
    mstrButton = "cmdScrabbleCount"
End Sub

'*****************************************************************************
'subroutine will be executed when the Scrabble Max Score command button is clicked
'*****************************************************************************
Private Sub cmdScrabbleMax_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdScrabbleMax" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    Else
        'order by maximum
        mstrOrder = "DESC"
    End If
    Call sortStatistics("ScrabbleMax")
    mstrButton = "cmdScrabbleMax"
End Sub

'*****************************************************************************
'subroutine will be executed when the Snap Count command button is clicked
'*****************************************************************************
Private Sub cmdSnapCount_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdSnapCount" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    Else
        'order by max
        mstrOrder = "DESC"
    End If
    Call sortStatistics("SnapMatches")
    mstrButton = "cmdSnapCount"
End Sub

'*****************************************************************************
'subroutine will be executed when the Snap Max Score command button is clicked
'*****************************************************************************
Private Sub cmdSnapMax_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdSnapMax" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    Else
        mstrOrder = "DESC"
    End If
    Call sortStatistics("SnapMax")
    mstrButton = "cmdSnapMax"
    
End Sub

'*****************************************************************************
'subroutine will be executed when the Spillikins Count command button is clicked
'*****************************************************************************
Private Sub cmdSpillikinsCount_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdSpillikinsCount" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    Else
        mstrOrder = "DESC"
    End If
     Call sortStatistics("SpillikinsMatches")
    mstrButton = "cmdSpillikinsCount"
  
End Sub

'*****************************************************************************
'subroutine will be executed when the Spillikins Max Score command button is clicked
'*****************************************************************************
Private Sub cmdSpillikinsMax_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdSpillikinsMax" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    Else
        mstrOrder = "DESC"
    End If
    Call sortStatistics("SpillikinsMax")
    mstrButton = "cmdSpillikinsMax"
    
End Sub

'*****************************************************************************
'subroutine will be executed when the Student name command button is clicked
'*****************************************************************************
Private Sub cmdStudentName_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdStudentName" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        'order alphabetically starting from a-z
        Else
            mstrOrder = "ASC"
        End If
    Else
        mstrOrder = "ASC"
    End If
    Call sortStatistics("StudentName")
    mstrButton = "cmdStudentName"
End Sub

'*****************************************************************************
'subroutine will be executed when the Cribbage Won Count command button is clicked
'*****************************************************************************
Private Sub cmdWonCribbage_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdWonCribbage" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    Else
        mstrOrder = "DESC"
    End If
    Call sortStatistics("WonCribbage")
    mstrButton = "cmdWonCribbage"
    
End Sub

'*****************************************************************************
'subroutine will be executed when the Scrabble Won Count command button is clicked
'*****************************************************************************
Private Sub cmdWonScrabble_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdWonScrabble" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    Else
        mstrOrder = "DESC"
    End If
    Call sortStatistics("WonScrabble")
    mstrButton = "cmdWonScrabble"
    
End Sub

'*****************************************************************************
'subroutine will be executed when the Snap Won count command button is clicked
'*****************************************************************************
Private Sub cmdWonSnap_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdWonSnap" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    Else
        mstrOrder = "DESC"
    End If
    Call sortStatistics("WonSnap")
    mstrButton = "cmdWonSnap"

End Sub

'*****************************************************************************
'subroutine will be executed when the Spillikins Won Countcommand button is clicked
'*****************************************************************************
Private Sub cmdWonSpillikins_Click()
    'if this button was the last to be clicked reverse the sorting
    If mstrButton = "cmdWonSpillikins" Then
        If mstrOrder = "ASC" Then
            mstrOrder = "DESC"
        Else
            mstrOrder = "ASC"
        End If
    Else
        mstrOrder = "DESC"
    End If
    Call sortStatistics("WonSpillikins")
    mstrButton = "cmdWonSpillikins"
    
End Sub

'*****************************************************************************
'subroutine will be executed when the form is loaded
'*****************************************************************************
Private Sub Form_Load()
    mstrOrder = "ASC"
    Call sortStatistics("StudentName")
    mstrButton = "cmdStudentName"
End Sub
'*****************************************************************************
'subroutine, read student statistics  from the database
'and sort specific field depending on the value of mstrOrder
'input: the attribute to be sorted:strAttr
'*****************************************************************************
Private Sub sortStatistics(strAttr As String)
 Dim intX As Integer
 'hold the colour of a specific introw of label boxes label box
 Dim colour As Variant
 
 'read the first record
    gadoCommand.CommandText = "SELECT * FROM statistics ORDER BY " _
    & strAttr & " " & mstrOrder
        Set gadoRecordSet = gadoCommand.Execute
        'dispaly reord
        lblName(0).Caption = gadoRecordSet.Fields(0)
        lblSnapCount(0).Caption = gadoRecordSet.Fields(1)
        lblCribbageCount(0).Caption = gadoRecordSet.Fields(2)
        lblSpillikinsCount(0).Caption = gadoRecordSet.Fields(3)
        lblScrabbleCount(0).Caption = gadoRecordSet.Fields(4)
        lblWonSnap(0).Caption = gadoRecordSet.Fields(5)
        lblWonCribbage(0).Caption = gadoRecordSet.Fields(6)
        lblWonSpillikins(0).Caption = gadoRecordSet.Fields(7)
        lblWonScrabble(0).Caption = gadoRecordSet.Fields(8)
        lblSnapMax(0).Caption = gadoRecordSet.Fields(9)
        lblCribbageMax(0).Caption = gadoRecordSet.Fields(10)
        lblSpillikinsMax(0).Caption = gadoRecordSet.Fields(11)
        lblScrabbleMax(0).Caption = gadoRecordSet.Fields(12)
        
        'change colours of the labels based on student's house colour
        If gadoRecordSet.Fields(13) = "Yellow" Then
            colour = &HC0FFFF
        Else
            colour = &HC0FFC0
        End If
        'change the colours
        Call setColour(0, colour)
'read the remaining records
    For intX = 1 To 35
    
        gadoCommand.CommandText = "SELECT * FROM statistics ORDER BY " _
        & strAttr & " " & mstrOrder
        Set gadoRecordSet = gadoCommand.Execute
        gadoRecordSet.GetRows (intX)
        lblName(intX).Caption = gadoRecordSet.Fields(0)
        lblSnapCount(intX).Caption = gadoRecordSet.Fields(1)
        lblCribbageCount(intX).Caption = gadoRecordSet.Fields(2)
        lblSpillikinsCount(intX).Caption = gadoRecordSet.Fields(3)
        lblScrabbleCount(intX).Caption = gadoRecordSet.Fields(4)
        lblWonSnap(intX).Caption = gadoRecordSet.Fields(5)
        lblWonCribbage(intX).Caption = gadoRecordSet.Fields(6)
        lblWonSpillikins(intX).Caption = gadoRecordSet.Fields(7)
        lblWonScrabble(intX).Caption = gadoRecordSet.Fields(8)
        lblSnapMax(intX).Caption = gadoRecordSet.Fields(9)
        lblCribbageMax(intX).Caption = gadoRecordSet.Fields(10)
        lblSpillikinsMax(intX).Caption = gadoRecordSet.Fields(11)
        lblScrabbleMax(intX).Caption = gadoRecordSet.Fields(12)
        
        If gadoRecordSet.Fields(13) = "Yellow" Then
            colour = &HC0FFFF
        Else
            colour = &HC0FFC0
        End If
        Call setColour(intX, colour)
    Next intX
End Sub

'*****************************************************************************
'subroutine, change the colour of specific introw of labels to varClr
'input: the number of introw of labels  to be changed:intRow
'input:the desired colour :varClr
'*****************************************************************************
Private Sub setColour(intRow As Integer, varClr As Variant)
        lblName(intRow).BackColor = varClr
        lblSnapCount(intRow).BackColor = varClr
        lblCribbageCount(intRow).BackColor = varClr
        lblSpillikinsCount(intRow).BackColor = varClr
        lblScrabbleCount(intRow).BackColor = varClr
        lblWonSnap(intRow).BackColor = varClr
        lblWonCribbage(intRow).BackColor = varClr
        lblWonSpillikins(intRow).BackColor = varClr
        lblWonScrabble(intRow).BackColor = varClr
        lblSnapMax(intRow).BackColor = varClr
        lblCribbageMax(intRow).BackColor = varClr
        lblSpillikinsMax(intRow).BackColor = varClr
        lblScrabbleMax(intRow).BackColor = varClr
End Sub

