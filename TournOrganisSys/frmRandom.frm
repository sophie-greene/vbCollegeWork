VERSION 5.00
Begin VB.Form frmRandom 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Random"
   ClientHeight    =   8235
   ClientLeft      =   450
   ClientTop       =   735
   ClientWidth     =   14235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00C0FFC0&
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
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   388
      Top             =   7680
      Width           =   3735
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00C0FFC0&
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
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   387
      Top             =   7680
      Width           =   3255
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFC0&
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   386
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Edit"
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
      Left            =   2535
      Style           =   1  'Graphical
      TabIndex        =   384
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Frame fraScrabble 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Scrabble"
      Height          =   7335
      Left            =   10680
      TabIndex        =   97
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   1080
         TabIndex        =   71
         Top             =   6840
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   2400
         TabIndex        =   306
         Top             =   6840
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   16
         Left            =   1080
         TabIndex        =   70
         Top             =   6480
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   16
         Left            =   2400
         TabIndex        =   303
         Top             =   6480
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   1080
         TabIndex        =   69
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   2400
         TabIndex        =   300
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   1080
         TabIndex        =   68
         Top             =   5760
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   2400
         TabIndex        =   297
         Top             =   5760
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   1080
         TabIndex        =   67
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   2400
         TabIndex        =   294
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   1080
         TabIndex        =   66
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   2400
         TabIndex        =   291
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   1080
         TabIndex        =   65
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   2400
         TabIndex        =   288
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   1080
         TabIndex        =   64
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   2400
         TabIndex        =   285
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   1080
         TabIndex        =   63
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   2400
         TabIndex        =   282
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   62
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   2400
         TabIndex        =   279
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   61
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   276
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   60
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   273
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   59
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   270
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   58
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   267
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   57
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   264
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   56
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   261
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   55
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   258
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtGScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   54
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtYScrabbleResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   255
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   17
         Left            =   2880
         TabIndex        =   385
         Top             =   6840
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   16
         Left            =   2880
         TabIndex        =   383
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   15
         Left            =   2880
         TabIndex        =   382
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   14
         Left            =   2880
         TabIndex        =   381
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   13
         Left            =   2880
         TabIndex        =   380
         Top             =   5400
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   12
         Left            =   2880
         TabIndex        =   379
         Top             =   5040
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   11
         Left            =   2880
         TabIndex        =   378
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   10
         Left            =   2880
         TabIndex        =   377
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   9
         Left            =   2880
         TabIndex        =   376
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   375
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   374
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   373
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   372
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   371
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   370
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   369
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   368
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblScrabbleWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   367
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Winner"
         Height          =   375
         Index           =   19
         Left            =   2760
         TabIndex        =   365
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   308
         Top             =   6840
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   17
         Left            =   1560
         TabIndex        =   307
         Top             =   6840
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   16
         Left            =   240
         TabIndex        =   305
         Top             =   6480
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   16
         Left            =   1560
         TabIndex        =   304
         Top             =   6480
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   15
         Left            =   240
         TabIndex        =   302
         Top             =   6120
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   15
         Left            =   1560
         TabIndex        =   301
         Top             =   6120
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   14
         Left            =   240
         TabIndex        =   299
         Top             =   5760
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   14
         Left            =   1560
         TabIndex        =   298
         Top             =   5760
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   13
         Left            =   240
         TabIndex        =   296
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   13
         Left            =   1560
         TabIndex        =   295
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   12
         Left            =   240
         TabIndex        =   293
         Top             =   5040
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   12
         Left            =   1560
         TabIndex        =   292
         Top             =   5040
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   290
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   11
         Left            =   1560
         TabIndex        =   289
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   10
         Left            =   240
         TabIndex        =   287
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   10
         Left            =   1560
         TabIndex        =   286
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   284
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   9
         Left            =   1560
         TabIndex        =   283
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   281
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   280
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   278
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   277
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   275
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   274
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   272
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   271
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   269
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   268
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   266
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   265
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   263
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   262
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   260
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   259
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblGScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   257
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblYScrabbleStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   256
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   101
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
         Height          =   375
         Index           =   6
         Left            =   1560
         TabIndex        =   100
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   99
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         Height          =   375
         Index           =   4
         Left            =   960
         TabIndex        =   98
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdWinner 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Winner"
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
      Left            =   11940
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Frame fraCribbage 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cribbage"
      Height          =   7335
      Left            =   3615
      TabIndex        =   92
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   2505
         TabIndex        =   198
         Top             =   6840
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   1125
         TabIndex        =   35
         Top             =   6840
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   16
         Left            =   2505
         TabIndex        =   195
         Top             =   6480
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   16
         Left            =   1125
         TabIndex        =   34
         Top             =   6480
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   2505
         TabIndex        =   192
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   1125
         TabIndex        =   33
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   2505
         TabIndex        =   189
         Top             =   5760
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   1125
         TabIndex        =   32
         Top             =   5760
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   2505
         TabIndex        =   186
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   1125
         TabIndex        =   31
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   2520
         TabIndex        =   183
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   1125
         TabIndex        =   30
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   2505
         TabIndex        =   180
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   1125
         TabIndex        =   29
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   2505
         TabIndex        =   177
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   1125
         TabIndex        =   28
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   2505
         TabIndex        =   174
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   1125
         TabIndex        =   27
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   2520
         TabIndex        =   171
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   1125
         TabIndex        =   26
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   2505
         TabIndex        =   168
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   1125
         TabIndex        =   25
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   2505
         TabIndex        =   165
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1125
         TabIndex        =   24
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   2505
         TabIndex        =   162
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1125
         TabIndex        =   23
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   2505
         TabIndex        =   159
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1125
         TabIndex        =   22
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   2505
         TabIndex        =   156
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1125
         TabIndex        =   21
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   2505
         TabIndex        =   153
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1125
         TabIndex        =   20
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   2505
         TabIndex        =   150
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1125
         TabIndex        =   19
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtYCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   2505
         TabIndex        =   147
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtGCribbageResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1125
         TabIndex        =   18
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   16
         Left            =   3000
         TabIndex        =   363
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   362
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   361
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   360
         Top             =   5400
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   359
         Top             =   5040
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   358
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   10
         Left            =   3000
         TabIndex        =   357
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   356
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   355
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   354
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   353
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   352
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   351
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   350
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   349
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   348
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   17
         Left            =   3000
         TabIndex        =   347
         Top             =   6840
         Width           =   375
      End
      Begin VB.Label lblCribbageWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   346
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Winner"
         Height          =   375
         Index           =   17
         Left            =   2880
         TabIndex        =   328
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   17
         Left            =   1620
         TabIndex        =   200
         Top             =   6840
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   199
         Top             =   6840
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   16
         Left            =   1620
         TabIndex        =   197
         Top             =   6480
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   16
         Left            =   240
         TabIndex        =   196
         Top             =   6480
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   15
         Left            =   1620
         TabIndex        =   194
         Top             =   6120
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   15
         Left            =   240
         TabIndex        =   193
         Top             =   6120
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   14
         Left            =   1620
         TabIndex        =   191
         Top             =   5760
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   14
         Left            =   240
         TabIndex        =   190
         Top             =   5760
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   13
         Left            =   1620
         TabIndex        =   188
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   13
         Left            =   240
         TabIndex        =   187
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   12
         Left            =   1620
         TabIndex        =   185
         Top             =   5040
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   12
         Left            =   240
         TabIndex        =   184
         Top             =   5040
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   11
         Left            =   1620
         TabIndex        =   182
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   181
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   10
         Left            =   1620
         TabIndex        =   179
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   10
         Left            =   240
         TabIndex        =   178
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   9
         Left            =   1620
         TabIndex        =   176
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   175
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   8
         Left            =   1620
         TabIndex        =   173
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   172
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   7
         Left            =   1620
         TabIndex        =   170
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   169
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   6
         Left            =   1620
         TabIndex        =   167
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   166
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   5
         Left            =   1620
         TabIndex        =   164
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   163
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   4
         Left            =   1620
         TabIndex        =   161
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   160
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   3
         Left            =   1620
         TabIndex        =   158
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   157
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   2
         Left            =   1620
         TabIndex        =   155
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   154
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   152
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   151
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblYCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   149
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblGCribbageStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   148
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   96
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
         Height          =   375
         Index           =   14
         Left            =   1680
         TabIndex        =   95
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         Height          =   375
         Index           =   13
         Left            =   2400
         TabIndex        =   94
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         Height          =   375
         Index           =   12
         Left            =   1080
         TabIndex        =   93
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraSpillikins 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Spillikins"
      Height          =   7335
      Left            =   7110
      TabIndex        =   87
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   2400
         TabIndex        =   252
         Top             =   6840
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   1080
         TabIndex        =   53
         Top             =   6840
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   16
         Left            =   2400
         TabIndex        =   249
         Top             =   6480
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   16
         Left            =   1080
         TabIndex        =   52
         Top             =   6480
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   2400
         TabIndex        =   246
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   1080
         TabIndex        =   51
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   2400
         TabIndex        =   243
         Top             =   5760
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   1080
         TabIndex        =   50
         Top             =   5760
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   2400
         TabIndex        =   240
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   1080
         TabIndex        =   49
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   2400
         TabIndex        =   237
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   1080
         TabIndex        =   48
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   2400
         TabIndex        =   234
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   1080
         TabIndex        =   47
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   2400
         TabIndex        =   231
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   1080
         TabIndex        =   46
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   2400
         TabIndex        =   228
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   1080
         TabIndex        =   45
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   2400
         TabIndex        =   225
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   44
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   222
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   43
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   219
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   42
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   216
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   41
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   213
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   40
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   210
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   39
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   207
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   38
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   204
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   37
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtYSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   201
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtGSpillikinsResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   36
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   17
         Left            =   2880
         TabIndex        =   366
         Top             =   6840
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Winner"
         Height          =   375
         Index           =   18
         Left            =   2760
         TabIndex        =   364
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   16
         Left            =   2880
         TabIndex        =   345
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   15
         Left            =   2880
         TabIndex        =   344
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   14
         Left            =   2880
         TabIndex        =   343
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   13
         Left            =   2880
         TabIndex        =   342
         Top             =   5400
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   12
         Left            =   2880
         TabIndex        =   341
         Top             =   5040
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   11
         Left            =   2880
         TabIndex        =   340
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   10
         Left            =   2880
         TabIndex        =   339
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   9
         Left            =   2880
         TabIndex        =   338
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   337
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   336
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   335
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   334
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   333
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   332
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   331
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   330
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblSpillikinsWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   329
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   17
         Left            =   1560
         TabIndex        =   254
         Top             =   6840
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   253
         Top             =   6840
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   16
         Left            =   1560
         TabIndex        =   251
         Top             =   6480
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   16
         Left            =   240
         TabIndex        =   250
         Top             =   6480
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   15
         Left            =   1560
         TabIndex        =   248
         Top             =   6120
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   15
         Left            =   240
         TabIndex        =   247
         Top             =   6120
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   14
         Left            =   1560
         TabIndex        =   245
         Top             =   5760
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   14
         Left            =   240
         TabIndex        =   244
         Top             =   5760
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   13
         Left            =   1560
         TabIndex        =   242
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   13
         Left            =   240
         TabIndex        =   241
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   12
         Left            =   1560
         TabIndex        =   239
         Top             =   5040
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   12
         Left            =   240
         TabIndex        =   238
         Top             =   5040
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   11
         Left            =   1560
         TabIndex        =   236
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   235
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   10
         Left            =   1560
         TabIndex        =   233
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   10
         Left            =   240
         TabIndex        =   232
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   9
         Left            =   1560
         TabIndex        =   230
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   229
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   227
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   226
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   224
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   223
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   221
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   220
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   218
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   217
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   215
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   214
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   212
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   211
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   209
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   208
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   206
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   205
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblYSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   203
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblGSpillikinsStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   202
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         Height          =   375
         Index           =   11
         Left            =   360
         TabIndex        =   91
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
         Height          =   375
         Index           =   10
         Left            =   1560
         TabIndex        =   90
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         Height          =   375
         Index           =   9
         Left            =   2280
         TabIndex        =   89
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         Height          =   375
         Index           =   8
         Left            =   960
         TabIndex        =   88
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraSnap 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Snap"
      Height          =   7335
      Left            =   120
      TabIndex        =   73
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   1125
         TabIndex        =   17
         Top             =   6840
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   2505
         TabIndex        =   144
         Top             =   6840
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   16
         Left            =   1125
         TabIndex        =   16
         Top             =   6480
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   16
         Left            =   2505
         TabIndex        =   141
         Top             =   6480
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   1125
         TabIndex        =   15
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   2505
         TabIndex        =   138
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   1125
         TabIndex        =   14
         Top             =   5760
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   2505
         TabIndex        =   135
         Top             =   5760
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   1125
         TabIndex        =   13
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   2505
         TabIndex        =   132
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   1125
         TabIndex        =   12
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   2505
         TabIndex        =   129
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   1125
         TabIndex        =   11
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   2505
         TabIndex        =   126
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   1125
         TabIndex        =   10
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   2505
         TabIndex        =   123
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   1125
         TabIndex        =   9
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   2505
         TabIndex        =   120
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   1125
         TabIndex        =   8
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   2505
         TabIndex        =   117
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   1125
         TabIndex        =   7
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   2505
         TabIndex        =   114
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1125
         TabIndex        =   6
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   2505
         TabIndex        =   111
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1125
         TabIndex        =   5
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   2505
         TabIndex        =   108
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1125
         TabIndex        =   4
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   2505
         TabIndex        =   105
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1125
         TabIndex        =   3
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   2505
         TabIndex        =   102
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1125
         TabIndex        =   2
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   2505
         TabIndex        =   82
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1125
         TabIndex        =   1
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   2505
         TabIndex        =   79
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtYSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   2505
         TabIndex        =   76
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtGSnapResult 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1125
         TabIndex        =   0
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   17
         Left            =   3000
         TabIndex        =   327
         Top             =   6840
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   16
         Left            =   3000
         TabIndex        =   326
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   325
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   324
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   323
         Top             =   5400
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   322
         Top             =   5040
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   321
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   10
         Left            =   3000
         TabIndex        =   320
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   319
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   318
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   317
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   316
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   315
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   314
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   313
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   312
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   311
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Winner"
         Height          =   375
         Index           =   16
         Left            =   2880
         TabIndex        =   310
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSnapWinner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   309
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   146
         Top             =   6840
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   17
         Left            =   1620
         TabIndex        =   145
         Top             =   6840
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   16
         Left            =   240
         TabIndex        =   143
         Top             =   6480
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   16
         Left            =   1620
         TabIndex        =   142
         Top             =   6480
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   15
         Left            =   240
         TabIndex        =   140
         Top             =   6120
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   15
         Left            =   1620
         TabIndex        =   139
         Top             =   6120
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   14
         Left            =   240
         TabIndex        =   137
         Top             =   5760
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   14
         Left            =   1620
         TabIndex        =   136
         Top             =   5760
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   13
         Left            =   240
         TabIndex        =   134
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   13
         Left            =   1620
         TabIndex        =   133
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   12
         Left            =   240
         TabIndex        =   131
         Top             =   5040
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   12
         Left            =   1620
         TabIndex        =   130
         Top             =   5040
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   128
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   11
         Left            =   1620
         TabIndex        =   127
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   10
         Left            =   240
         TabIndex        =   125
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   10
         Left            =   1620
         TabIndex        =   124
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   122
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   9
         Left            =   1620
         TabIndex        =   121
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   119
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   8
         Left            =   1620
         TabIndex        =   118
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   116
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   7
         Left            =   1620
         TabIndex        =   115
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   113
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   6
         Left            =   1620
         TabIndex        =   112
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   110
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   5
         Left            =   1620
         TabIndex        =   109
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   107
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   4
         Left            =   1620
         TabIndex        =   106
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   104
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   3
         Left            =   1620
         TabIndex        =   103
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   80
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   2
         Left            =   1620
         TabIndex        =   81
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   77
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   78
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   86
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         Height          =   375
         Index           =   3
         Left            =   2400
         TabIndex        =   85
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   84
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblYSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   75
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   83
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblGSnapStudent 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   74
         Top             =   720
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmRandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'************************** Tournament Organising System************************
'**********************************frmRandom Code*******************************
'*******************************Programer: S. Saqfelhait************************
'***********************************Date:07/04/2007*****************************
'*******************************************************************************
'this form is loaded from frmIntro "select game generation method and number of
'matches form, it is also loaded from manual form when save button is clicked
Option Explicit

'*****************************************************************************
'subroutine will be executed when the Change command button is clicked
'*****************************************************************************
Private Sub cmdChange_Click()
    frmIntro.Show
    Unload Me
End Sub

'*****************************************************************************
'subroutine will be executed when the Edit command button is clicked
'*****************************************************************************
Private Sub cmdEdit_Click()
    gboolEdit = True
    frmManual.Show
    Unload Me
End Sub

'*****************************************************************************
'subroutine will be executed when the Exit command button is clicked
'*****************************************************************************
Private Sub cmdExit_Click()
    'define a variable to capture the msgbox response
    Dim varResponce As Variant
    varResponce = myMsgBox("Are You Sure you Want to Exit", _
                    "YesNo", "Attention")
    'exit program if response if Yes
    If varResponce = vbYes Then
        DB_Disconnect
        End
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
'subroutine will be executed when the Winner command button is clicked
'*****************************************************************************
Private Sub cmdWinner_Click()
    'define a variable to count the green house points
    Dim GPoints As Integer
    'define a variable to count the yellow house points
    Dim YPoints As Integer
    Dim intX As Integer
    'define a variable to be used as a flag to track any empty result textbox
    Dim emptyResults As Boolean
    'initialise the variables
    GPoints = 0
    YPoints = 0
    emptyResults = False
    
    'find and highlight empty results text boxes
     For intX = 0 To gintSnapMatch - 1
        If txtGSnapResult(intX).Text = "" Then
            txtGSnapResult(intX).BackColor = &HC0E0FF
            txtYSnapResult(intX).BackColor = &HC0E0FF
            emptyResults = True
        Else
            txtGSnapResult(intX).BackColor = vbWhite
            txtYSnapResult(intX).BackColor = vbWhite
        End If
    Next intX
    
    For intX = 0 To gintCribbageMatch - 1
        If txtGCribbageResult(intX).Text = "" Then
            txtGCribbageResult(intX).BackColor = &HC0E0FF
            txtYCribbageResult(intX).BackColor = &HC0E0FF
            emptyResults = True
        Else
            txtGCribbageResult(intX).BackColor = vbWhite
            txtYCribbageResult(intX).BackColor = vbWhite
        End If
    Next intX
    
    For intX = 0 To gintSpillikinsMatch - 1
        If txtGSpillikinsResult(intX).Text = "" Then
            txtGSpillikinsResult(intX).BackColor = &HC0E0FF
            txtYSpillikinsResult(intX).BackColor = &HC0E0FF
            emptyResults = True
        Else
            txtGSpillikinsResult(intX).BackColor = vbWhite
            txtYSpillikinsResult(intX).BackColor = vbWhite
        End If
       
    Next intX
    
    For intX = 0 To gintScrabbleMatch - 1
        If txtGScrabbleResult(intX).Text = "" Then
            txtGScrabbleResult(intX).BackColor = &HC0E0FF
            txtYScrabbleResult(intX).BackColor = &HC0E0FF
            emptyResults = True
        Else
            txtGScrabbleResult(intX).BackColor = vbWhite
            txtYScrabbleResult(intX).BackColor = vbWhite
        End If
    Next intX
    
    'if any empty result text boxes were found tell the user and exit the sub
    If emptyResults = True Then
        myMsgBox "Please fill in the highlighted result boxes" _
        , "Ok", "Attention" '
        Exit Sub
    End If
    'calculate the results
    For intX = 0 To gintSnapMatch - 1
        If Val(txtGSnapResult(intX).Text) > _
        Val(txtYSnapResult(intX).Text) Then
            Call UpdateMatchWonCount(strGame(intX, 0), "Snap", _
                Val(txtGSnapResult(intX).Text))
            GPoints = GPoints + 1
        Else
            Call UpdateMatchWonCount(strGame(intX, 1), "Snap", _
            Val(txtYSnapResult(intX).Text))
            YPoints = YPoints + 1
        End If
    Next intX
    
    For intX = 0 To gintCribbageMatch - 1
        If Val(txtGCribbageResult(intX).Text) > _
        Val(txtYCribbageResult(intX).Text) Then
            Call UpdateMatchWonCount(strGame(intX, 2), "Cribbage" _
                , Val(txtGCribbageResult(intX).Text))
            GPoints = GPoints + 1
        Else
            Call UpdateMatchWonCount(strGame(intX, 3), "Cribbage" _
                , Val(txtYCribbageResult(intX).Text))
            YPoints = YPoints + 1
        End If
    Next intX
    
    For intX = 0 To gintSpillikinsMatch - 1
        If Val(txtGSpillikinsResult(intX).Text) > _
        Val(txtYSpillikinsResult(intX).Text) Then
            Call UpdateMatchWonCount(strGame(intX, 4), "Spillikins" _
                , Val(txtGSpillikinsResult(intX).Text))
            GPoints = GPoints + 1
        Else
            Call UpdateMatchWonCount(strGame(intX, 5), "Spillikins" _
                , Val(txtYSpillikinsResult(intX).Text))
            YPoints = YPoints + 1
        End If
    Next intX
    
    For intX = 0 To gintScrabbleMatch - 1
        If Val(txtGScrabbleResult(intX).Text) > _
        Val(txtYScrabbleResult(intX).Text) Then
            Call UpdateMatchWonCount(strGame(intX, 6), "Scrabble" _
                , Val(txtGScrabbleResult(intX).Text))
            GPoints = GPoints + 1
        Else
            Call UpdateMatchWonCount(strGame(intX, 7), "Scrabble" _
                , Val(txtGScrabbleResult(intX).Text))
            YPoints = YPoints + 1
        End If
    Next intX
    
    'determine the winner
    If GPoints > YPoints Then
        gstrWinner = "Green"
    Else
        gstrWinner = "Yellow"
    End If
    Dim tempCount As Integer
    'store game results and players in the database
    For intX = 0 To gintSnapMatch - 1
        Call UpdateMatchCount(strGame(intX, 0), "Snap")
        Call UpdateMatchCount(strGame(intX, 1), "Snap")
    Next intX
    
    For intX = 0 To gintCribbageMatch - 1
        Call UpdateMatchCount(strGame(intX, 2), "Cribbage")
        Call UpdateMatchCount(strGame(intX, 3), "Cribbage")
    Next intX
    
    For intX = 0 To gintSpillikinsMatch - 1
        Call UpdateMatchCount(strGame(intX, 4), "Spillikins")
        Call UpdateMatchCount(strGame(intX, 5), "Spillikins")
    Next intX
    
    For intX = 0 To gintScrabbleMatch - 1
        Call UpdateMatchCount(strGame(intX, 6), "Scrabble")
        Call UpdateMatchCount(strGame(intX, 7), "Scrabble")
    Next intX
    Call UpdateTournaments(gstrWinner)
    
    Unload Me
    frmWinner.Show
End Sub

'*****************************************************************************
'subroutine will be executed when the form is activated
'*****************************************************************************
Private Sub Form_Activate()
    'if the form was loaded from Manual form i.e gboolrandom=false
    If gboolRandom = False Then
        Call displayManual
    End If
    'find the class of each student and store it in each label box tiptool
    Dim intX As Integer
    'snap
    For intX = 0 To gintSnapMatch - 1
    'green opponent
    gadoCommand.CommandText = "SELECT * FROM Students WHERE studentName = '" & _
        lblGSnapStudent(intX).Caption & "'"
                Set gadoRecordSet = gadoCommand.Execute
        lblGSnapStudent(intX).ToolTipText = "Class " & gadoRecordSet.Fields(2)
     'yellow opponent
        gadoCommand.CommandText = "SELECT * FROM Students WHERE studentName = '" & _
        lblYSnapStudent(intX).Caption & "'"
                Set gadoRecordSet = gadoCommand.Execute
        lblYSnapStudent(intX).ToolTipText = "Class " & gadoRecordSet.Fields(2)
     Next intX
     'cribbage
     For intX = 0 To gintCribbageMatch - 1
        'green
        gadoCommand.CommandText = "SELECT * FROM Students WHERE studentName = '" & _
        lblGCribbageStudent(intX).Caption & "'"
                Set gadoRecordSet = gadoCommand.Execute
        lblGCribbageStudent(intX).ToolTipText = "Class " & gadoRecordSet.Fields(2)
        
         'yellow
         gadoCommand.CommandText = "SELECT * FROM Students WHERE studentName = '" & _
        lblYCribbageStudent(intX).Caption & "'"
                Set gadoRecordSet = gadoCommand.Execute
        lblYCribbageStudent(intX).ToolTipText = "Class " & gadoRecordSet.Fields(2)
    Next intX
    'spillikins
    For intX = 0 To gintSpillikinsMatch - 1
        'green
        gadoCommand.CommandText = "SELECT * FROM Students WHERE studentName = '" & _
        lblGSpillikinsStudent(intX).Caption & "'"
                Set gadoRecordSet = gadoCommand.Execute
        lblGSpillikinsStudent(intX).ToolTipText = "Class " & gadoRecordSet.Fields(2)
         
         'yellow
         gadoCommand.CommandText = "SELECT * FROM Students WHERE studentName = '" & _
        lblYSpillikinsStudent(intX).Caption & "'"
                Set gadoRecordSet = gadoCommand.Execute
        lblYSpillikinsStudent(intX).ToolTipText = "Class " & gadoRecordSet.Fields(2)
    Next intX
    'scrabble
    For intX = 0 To gintScrabbleMatch - 1
    'green
        gadoCommand.CommandText = "SELECT * FROM Students WHERE studentName = '" & _
        lblGScrabbleStudent(intX).Caption & "'"
                Set gadoRecordSet = gadoCommand.Execute
        lblGScrabbleStudent(intX).ToolTipText = "Class " & gadoRecordSet.Fields(2)
         
         'yellow
         gadoCommand.CommandText = "SELECT * FROM Students WHERE studentName = '" & _
        lblYScrabbleStudent(intX).Caption & "'"
                Set gadoRecordSet = gadoCommand.Execute
        lblYScrabbleStudent(intX).ToolTipText = "Class " & gadoRecordSet.Fields(2)
        
    Next intX
End Sub

'*****************************************************************************
'subroutine will be executed when the from is loaded
'*****************************************************************************
Private Sub Form_Load()

    Dim intX As Integer
    'disable result boxes corresponding to no opponents
    For intX = gintSnapMatch To 17
        txtGSnapResult(intX).Enabled = False
        txtYSnapResult(intX).Enabled = False
    Next intX
         
    For intX = gintCribbageMatch To 17
        txtGCribbageResult(intX).Enabled = False
        txtYCribbageResult(intX).Enabled = False
    Next intX
       
    For intX = gintSpillikinsMatch To 17
        txtGSpillikinsResult(intX).Enabled = False
        txtYSpillikinsResult(intX).Enabled = False
    Next intX
    
    For intX = gintScrabbleMatch To 17
        txtGScrabbleResult(intX).Enabled = False
        txtYScrabbleResult(intX).Enabled = False
    Next intX
    
    'if the form was loaded from Manual form i.e gboolrandom=false
    'display the players selected in the manual form
    If gboolRandom = False Then
        Call displayManual
    Else
    'generate a game schedule randomly
        Call GenerateRandom
    End If
End Sub



'*****************************************************************************
'subroutine will be executed when the mouse is moved on label number "Index" green
'*****************************************************************************
Private Sub lblGSnapStudent_MouseMove(Index As Integer, Button As Integer, _
Shift As Integer, X As Single, Y As Single)
    
    Dim intX As Integer
    
    For intX = 0 To gintSnapMatch - 1
        If intX = Index Then
            lblGSnapStudent(intX).ForeColor = vbRed
            lblYSnapStudent(intX).ForeColor = vbRed
            lblGSnapStudent(intX).FontBold = True
            lblYSnapStudent(intX).FontBold = True
        Else
            lblGSnapStudent(intX).ForeColor = vbBlack
            lblYSnapStudent(intX).ForeColor = vbBlack
            lblGSnapStudent(intX).FontBold = False
            lblYSnapStudent(intX).FontBold = False
        End If
    Next intX
    'highlight corresponding labels from the rest of the games and
    'highlight opponents
    For intX = 0 To gintCribbageMatch - 1
        If lblGCribbageStudent(intX).Caption = lblGSnapStudent(Index).Caption Then
            lblGCribbageStudent(intX).ForeColor = vbRed
            lblYCribbageStudent(intX).ForeColor = vbRed
            lblGCribbageStudent(intX).FontBold = True
            lblYCribbageStudent(intX).FontBold = True
        Else
            lblGCribbageStudent(intX).ForeColor = vbBlack
            lblYCribbageStudent(intX).ForeColor = vbBlack
            lblGCribbageStudent(intX).FontBold = False
            lblYCribbageStudent(intX).FontBold = False
        End If
    Next intX
    
     For intX = 0 To gintSpillikinsMatch - 1
        If lblGSpillikinsStudent(intX).Caption = lblGSnapStudent(Index).Caption Then
            lblGSpillikinsStudent(intX).ForeColor = vbRed
            lblYSpillikinsStudent(intX).ForeColor = vbRed
            lblGSpillikinsStudent(intX).FontBold = True
            lblYSpillikinsStudent(intX).FontBold = True
        Else
            lblGSpillikinsStudent(intX).ForeColor = vbBlack
            lblYSpillikinsStudent(intX).ForeColor = vbBlack
            lblGSpillikinsStudent(intX).FontBold = False
            lblYSpillikinsStudent(intX).FontBold = False
        End If
    Next intX
    
    For intX = 0 To gintScrabbleMatch - 1
        If lblGScrabbleStudent(intX).Caption = lblGSnapStudent(Index).Caption Then
            lblGScrabbleStudent(intX).ForeColor = vbRed
            lblYScrabbleStudent(intX).ForeColor = vbRed
            lblGScrabbleStudent(intX).FontBold = True
            lblYScrabbleStudent(intX).FontBold = True
        Else
            lblGScrabbleStudent(intX).ForeColor = vbBlack
            lblYScrabbleStudent(intX).ForeColor = vbBlack
            lblGScrabbleStudent(intX).FontBold = False
            lblYScrabbleStudent(intX).FontBold = False
        End If
    Next intX
    
End Sub

'*****************************************************************************
'subroutine will be executed when the mouse is moved on label number "Index" yellow
'*****************************************************************************
Private Sub lblYSnapStudent_MouseMove(Index As Integer, Button As Integer, _
Shift As Integer, X As Single, Y As Single)
       
    Dim intX As Integer
    
    For intX = 0 To gintSnapMatch - 1
        If intX = Index Then
            lblGSnapStudent(intX).ForeColor = vbRed
            lblYSnapStudent(intX).ForeColor = vbRed
            lblGSnapStudent(intX).FontBold = True
            lblYSnapStudent(intX).FontBold = True
        Else
            lblGSnapStudent(intX).ForeColor = vbBlack
            lblYSnapStudent(intX).ForeColor = vbBlack
            lblGSnapStudent(intX).FontBold = False
            lblYSnapStudent(intX).FontBold = False
        End If
    Next intX
    'highlight corresponding labels from the rest of the games and
    'highlight opponents
    For intX = 0 To gintCribbageMatch - 1
        If lblYCribbageStudent(intX).Caption = lblYSnapStudent(Index).Caption Then
            lblGCribbageStudent(intX).ForeColor = vbRed
            lblYCribbageStudent(intX).ForeColor = vbRed
            lblGCribbageStudent(intX).FontBold = True
            lblYCribbageStudent(intX).FontBold = True
        Else
            lblGCribbageStudent(intX).ForeColor = vbBlack
            lblYCribbageStudent(intX).ForeColor = vbBlack
            lblGCribbageStudent(intX).FontBold = False
            lblYCribbageStudent(intX).FontBold = False
        End If
    Next intX
    
     For intX = 0 To gintSpillikinsMatch - 1
        If lblYSpillikinsStudent(intX).Caption = lblYSnapStudent(Index).Caption Then
            lblGSpillikinsStudent(intX).ForeColor = vbRed
            lblYSpillikinsStudent(intX).ForeColor = vbRed
            lblGSpillikinsStudent(intX).FontBold = True
            lblYSpillikinsStudent(intX).FontBold = True
        Else
            lblGSpillikinsStudent(intX).ForeColor = vbBlack
            lblYSpillikinsStudent(intX).ForeColor = vbBlack
            lblGSpillikinsStudent(intX).FontBold = False
            lblYSpillikinsStudent(intX).FontBold = False
        End If
    Next intX
    
    For intX = 0 To gintScrabbleMatch - 1
        If lblYScrabbleStudent(intX).Caption = lblYSnapStudent(Index).Caption Then
            lblGScrabbleStudent(intX).ForeColor = vbRed
            lblYScrabbleStudent(intX).ForeColor = vbRed
            lblGScrabbleStudent(intX).FontBold = True
            lblYScrabbleStudent(intX).FontBold = True
        Else
            lblGScrabbleStudent(intX).ForeColor = vbBlack
            lblYScrabbleStudent(intX).ForeColor = vbBlack
            lblGScrabbleStudent(intX).FontBold = False
            lblYScrabbleStudent(intX).FontBold = False
        End If
    Next intX
    
End Sub

'*****************************************************************************
'subroutine will be executed when the text is changed in any text box of the
'Green students Snap results text boxes control array
'*****************************************************************************
Private Sub txtGSnapResult_Change(Index As Integer)
    'if the change results empty box, ignore the change
    If txtGSnapResult(Index) = "" Then Exit Sub
    
    'check the validity of input text
    If Val(txtGSnapResult(Index).Text) < 0 Or Val(txtGSnapResult(Index).Text) > 3 Or _
        Not IsNumeric((txtGSnapResult(Index).Text)) Then
        myMsgBox "Invalid Result, Enter a value between 0 and 3" _
        , "Ok", "Attention"
        txtGSnapResult(Index) = ""
        txtYSnapResult(Index) = ""
        txtGSnapResult(Index).SetFocus
        Exit Sub
    End If
    'compute the score of the opponent based on the value entered
    txtYSnapResult(Index).Text = _
        str$(3 - Val(txtGSnapResult(Index).Text))
        
    'determine the winner of the match
    If Val(txtGSnapResult(Index).Text) >= 2 Then
        lblSnapWinner(Index).BackColor = vbGreen
        lblSnapWinner(Index).Caption = "G"
    End If
End Sub

'*****************************************************************************
'subroutine will be executed when the text is changed in any text box of the
'yellow students Snap results text boxes control array
'*****************************************************************************
Private Sub txtYSnapResult_Change(Index As Integer)
    'if the change results empty box, ignore the change
    If txtYSnapResult(Index) = "" Then Exit Sub
    
    'check the validity of input text
    If Val(txtYSnapResult(Index).Text) < 0 Or Val(txtYSnapResult(Index).Text) > 3 Or _
    Not IsNumeric(txtYSnapResult(Index).Text) Then
        myMsgBox "Invalid Result, Enter a value between 0 and 3" _
        , "Ok", "Attention"
        txtYSnapResult(Index) = ""
        txtGSnapResult(Index) = ""
        txtYSnapResult(Index).SetFocus
        Exit Sub
    End If
    'compute the score of the opponent based on the value entered
    txtGSnapResult(Index).Text = _
        str$(3 - Val(txtYSnapResult(Index).Text))
        
    'determine the winner of the match
    If Val(txtYSnapResult(Index).Text) >= 2 Then
        lblSnapWinner(Index).BackColor = vbYellow
        lblSnapWinner(Index).Caption = "Y"
    End If
End Sub

'*****************************************************************************
'subroutine will be executed when the text is changed in any text box of the
'yellow students Snap results text boxes control array
'*****************************************************************************
Private Sub txtGcribbageResult_Change(Index As Integer)
    'if the change results empty box, ignore the change
    If txtGCribbageResult(Index) = "" Then Exit Sub
    txtYCribbageResult(Index) = ""
    
    'check the validity of input text
    If Val(txtGCribbageResult(Index).Text) < 0 Or Val(txtGCribbageResult(Index).Text) > 3 Then
        myMsgBox "Invalid Result, Enter a value between 0 and 3" _
        , "Ok", "Attention"
        txtGCribbageResult(Index) = ""
        txtGCribbageResult(Index).SetFocus
        Exit Sub
    End If
    'compute the score of the opponent based on the value entered
    txtYCribbageResult(Index).Text = _
    str$(3 - Val(txtGCribbageResult(Index).Text))
            
    'determine the winner of the match
    If Val(txtGCribbageResult(Index).Text) >= 2 Then
        lblCribbageWinner(Index).BackColor = vbGreen
        lblCribbageWinner(Index).Caption = "G"
    End If
    
End Sub

'*****************************************************************************
'subroutine will be executed when the text is changed in any text box of the
'yellow students Cribbage results text boxes control array
'*****************************************************************************
Private Sub txtYCribbageResult_Change(Index As Integer)
    'if the change results empty box, ignore the change
    If txtYCribbageResult(Index) = "" Then Exit Sub
    
    'check the validity of input text
    If Val(txtYCribbageResult(Index).Text) < 0 Or Val(txtYCribbageResult(Index).Text) > 3 Then
        myMsgBox "Invalid Result, Enter a value between 0 and 3" _
        , "Ok", "Attention"
        txtYCribbageResult(Index) = ""
        txtGCribbageResult(Index) = ""
        txtYCribbageResult(Index).SetFocus
        Exit Sub
    End If
    
    'compute the score of the opponent based on the value entered
    txtGCribbageResult(Index).Text = _
        str$(3 - Val(txtYCribbageResult(Index).Text))
        
    'determine the winner of the match
    If Val(txtYCribbageResult(Index).Text) >= 2 Then
            lblCribbageWinner(Index).BackColor = vbYellow
            lblCribbageWinner(Index).Caption = "Y"
    End If
End Sub

'*****************************************************************************
'subroutine will be executed when the text is changed in any text box of the
' Green students Spillikins results text boxes control array
'*****************************************************************************
Private Sub txtGspillikinsResult_Change(Index As Integer)
    'if the change results empty box, ignore the change
    If txtGSpillikinsResult(Index) = "" Then Exit Sub
    
    'check the validity of input text
    If Val(txtGSpillikinsResult(Index).Text) < 0 Or Val(txtGSpillikinsResult(Index).Text) > 3 Then
        myMsgBox "Invalid Result, Enter a value between 0 and 3" _
        , "Ok", "Attention"
        txtGSpillikinsResult(Index) = ""
        txtYSpillikinsResult(Index) = ""
        txtGSpillikinsResult(Index).SetFocus
        Exit Sub
    End If
    
    'compute the score of the opponent based on the value entered
    txtYSpillikinsResult(Index).Text = _
        str$(3 - Val(txtGSpillikinsResult(Index).Text))
        
    'determine the winner of the match
    If Val(txtGSpillikinsResult(Index).Text) >= 2 Then
        lblSpillikinsWinner(Index).BackColor = vbGreen
        lblSpillikinsWinner(Index).Caption = "G"
    End If
End Sub

'*****************************************************************************
'subroutine will be executed when the text is changed in any text box of the
'yellow students Spillikins results text boxes control array
'*****************************************************************************
Private Sub txtYSpillikinsResult_Change(Index As Integer)
    'if the change results empty box, ignore the change
    If txtYSpillikinsResult(Index) = "" Then Exit Sub
    
    'check the validity of input text
    If Val(txtYSpillikinsResult(Index).Text) < 0 Or Val(txtYSpillikinsResult(Index).Text) > 3 Then
        myMsgBox "Invalid Result, Enter a value between 0 and 3" _
        , "Ok", "Attention"
        txtYSpillikinsResult(Index) = ""
        txtGSpillikinsResult(Index) = ""
        txtYSpillikinsResult(Index).SetFocus
        Exit Sub
    End If
    
    'compute the score of the opponent based on the value entered
    txtGSpillikinsResult(Index).Text = _
        str$(3 - Val(txtYSpillikinsResult(Index).Text))
        
    'determine the winner of the match
     If Val(txtYSpillikinsResult(Index).Text) >= 2 Then
        lblSpillikinsWinner(Index).BackColor = vbYellow
        lblSpillikinsWinner(Index).Caption = "Y"
    End If
End Sub

'*****************************************************************************
'subroutine will be executed when the text is changed in any text box of the
'green students Scrabble results text boxes control array
'*****************************************************************************
Private Sub txtGScrabbleResult_Change(Index As Integer)
    'if the change results empty box, ignore the change
    If txtGScrabbleResult(Index) = "" Then Exit Sub
    
    'check the validity of input text
    If Val(txtGScrabbleResult(Index).Text) < 0 Or Val(txtGScrabbleResult(Index).Text) > 3 Then
        myMsgBox "Invalid Result, Enter a value between 0 and 3" _
        , "Ok", "Attention"
        txtGScrabbleResult(Index) = ""
        txtYScrabbleResult(Index) = ""
        txtGScrabbleResult(Index).SetFocus
        Exit Sub
    End If
    
    'compute the score of the opponent based on the value entered
    txtYScrabbleResult(Index).Text = _
        str$(3 - Val(txtGScrabbleResult(Index).Text))
        
    'determine the winner of the match
    If Val(txtGScrabbleResult(Index).Text) >= 2 Then
        lblScrabbleWinner(Index).BackColor = vbGreen
        lblScrabbleWinner(Index).Caption = "G"
    End If
End Sub

'*****************************************************************************
'subroutine will be executed when the text is changed in any text box of the
'yellow students Scrabble results text boxes control array
'*****************************************************************************
Private Sub txtYscrabbleResult_Change(Index As Integer)
    'if the change results empty box, ignore the change
    If txtYScrabbleResult(Index) = "" Then Exit Sub
    
    'check the validity of input text
    If Val(txtYScrabbleResult(Index).Text) < 0 Or Val(txtYScrabbleResult(Index).Text) > 3 Then
        myMsgBox "Invalid Result, Enter a value between 0 and 3" _
        , "Ok", "Attention"
        txtYScrabbleResult(Index) = ""
        txtGScrabbleResult(Index) = ""
        txtYScrabbleResult(Index).SetFocus
        Exit Sub
    End If
    
    'compute the score of the opponent based on the value entered
    txtGScrabbleResult(Index).Text = _
        str$(3 - Val(txtYScrabbleResult(Index).Text))
        
    'determine the winner of the match
    If Val(txtYScrabbleResult(Index).Text) >= 2 Then
        lblScrabbleWinner(Index).BackColor = vbYellow
        lblScrabbleWinner(Index).Caption = "Y"
    End If
End Sub

'*****************************************************************************
'subroutine will generate random players
'*****************************************************************************
Public Sub GenerateRandom()
    'variable to hold the number of potential opponents
    Dim intOpt As Integer
    Dim intX As Integer
    
    Dim intW As Integer
    Dim intS As Integer
    
    'array to store the available openent for each student
    Dim op(18) As String
    
    'read student records from the database
    getStudentArray strGstudent, "Green"
    getStudentArray strYstudent(), "Yellow"
add1:
    'clear games
    clearGameArray
    
    'snap columns in the game array are
    'green col =0
    'yellow col=1
    For intS = 0 To gintSnapMatch - 1
        Do
            Randomize (Timer)
                intX = Int(Rnd(Timer) * 18)
                If IsInGame(strGstudent(intX, 0), 0) = False Then
                    strGame(intS, 0) = strGstudent(intX, 0)
                    lblGSnapStudent(intS) = strGame(intS, 0)
                    Exit Do
                End If
        Loop
        'find potential opponents
        intOpt = potentialOponents(strGame(intS, 0), 1, op())
        
        'if no potential opponents were found restart the whole generation processes
        If intOpt = 0 Then GoTo add1
        'assign a random opponent
        Randomize (Timer)
        strGame(intS, 1) = op(Int(Rnd(Timer) * intOpt))
        lblYSnapStudent(intS) = strGame(intS, 1)
    Next intS
            
    'cribbage columns in the game array are
    'green col =2
    'yellow col=3
    For intS = 0 To gintCribbageMatch - 1
        Do
            Randomize (Timer)
                intX = Int(Rnd(Timer) * 18)
                If IsInGame(strGstudent(intX, 0), 2) = False Then
                    strGame(intS, 2) = strGstudent(intX, 0)
                  lblGCribbageStudent(intS) = strGame(intS, 2)
                    Exit Do
                End If
        Loop
        'find potential opponents
        intOpt = potentialOponents(strGame(intS, 2), 3, op())
        'if no potential opponents were found restart the whole generation processes
        If intOpt = 0 Then GoTo add1
        'assign a random opponent
        Randomize (Timer)
        strGame(intS, 3) = op(Int(Rnd(Timer) * intOpt))
        lblYCribbageStudent(intS) = strGame(intS, 3)
    Next intS
    'spillikins columns in the game array are
    'green col =4
    'yellow col=5
    For intS = 0 To gintSpillikinsMatch - 1
        Do
            Randomize (Timer)
                intX = Int(Rnd(Timer) * 18)
                If IsInGame(strGstudent(intX, 0), 4) = False Then
                    strGame(intS, 4) = strGstudent(intX, 0)
                    
                    lblGSpillikinsStudent(intS) = strGame(intS, 4)
                    Exit Do
                End If
        Loop
        'find potential opponents
        intOpt = potentialOponents(strGame(intS, 4), 5, op())
        'if no potential opponents were found restart the whole generation processes
        If intOpt = 0 Then GoTo add1
        'assign a random opponent
        Randomize (Timer)
        strGame(intS, 5) = op(Int(Rnd(Timer) * intOpt))
         lblYSpillikinsStudent(intS) = strGame(intS, 5)
    Next intS
           
    'scrabble columns in the game array are
    'green col =6
    'yellow col=7
    For intS = 0 To gintScrabbleMatch - 1
        Do
            Randomize (Timer)
                intX = Int(Rnd(Timer) * 18)
                If IsInGame(strGstudent(intX, 0), 6) = False Then
                    strGame(intS, 6) = strGstudent(intX, 0)
                    lblGScrabbleStudent(intS) = strGame(intS, 6)
                    Exit Do
                End If
        Loop
        'find potential opponents
        intOpt = potentialOponents(strGame(intS, 6), 7, op())
        'if no potential opponents were found restart the whole generation processes
        If intOpt = 0 Then GoTo add1
        'assign a random opponent
        Randomize (Timer)
        strGame(intS, 7) = op(Int(Rnd(Timer) * intOpt))
        lblYScrabbleStudent(intS) = strGame(intS, 7)
    Next intS
   
End Sub

'*****************************************************************************
'subroutine will display players selected manuallly
'*****************************************************************************
Private Sub displayManual()
    Dim intX As Integer
    Dim intY As Integer
    For intX = 0 To 17
            lblGSnapStudent(intX).Caption = strGame(intX, 0)
    Next intX
    For intX = 0 To 17
            lblYSnapStudent(intX).Caption = strGame(intX, 1)
    Next intX
    For intX = 0 To 17
            lblGCribbageStudent(intX).Caption = strGame(intX, 2)
    Next intX
    For intX = 0 To 17
            lblYCribbageStudent(intX).Caption = strGame(intX, 3)
    Next intX
    For intX = 0 To 17
            lblGSpillikinsStudent(intX).Caption = strGame(intX, 4)
    Next intX
    For intX = 0 To 17
            lblYSpillikinsStudent(intX).Caption = strGame(intX, 5)
    Next intX
    For intX = 0 To 17
            lblGScrabbleStudent(intX).Caption = strGame(intX, 6)
    Next intX
    For intX = 0 To 17
            lblYScrabbleStudent(intX).Caption = strGame(intX, 7)
    Next intX
End Sub
