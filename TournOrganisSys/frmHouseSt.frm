VERSION 5.00
Begin VB.Form frmHouseSt 
   BackColor       =   &H80000009&
   Caption         =   "Houses' Statistics"
   ClientHeight    =   8235
   ClientLeft      =   3555
   ClientTop       =   1875
   ClientWidth     =   14235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
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
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7050
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   5655
      Left            =   3030
      TabIndex        =   2
      Top             =   1050
      Width           =   8175
      Begin VB.Label lblTournaments 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label lblScrabbleWon 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   19
         Top             =   4140
         Width           =   1815
      End
      Begin VB.Label lblSpillikinsWon 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label lblCribbageWon 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   17
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label lblSnapWon 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblScrabbleCount 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   1980
         Width           =   1815
      End
      Begin VB.Label lblSpillikinsCount 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblCribbageCount 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   13
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblSnapCount 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tournaments Won:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   4680
         Width           =   4575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scrabble Matches Won Count: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   4140
         Width           =   4575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Spillikins Matches Won Count: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   3600
         Width           =   4575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cribbage Matches Won Count: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   3060
         Width           =   3975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Snap Matches Won Count: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scrabble Matches Played Count: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1980
         Width           =   4575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Spillikins Matches Played Count: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cribbage Matches Played Count: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   900
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Snap Matches Played Count: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.ComboBox cmbHouse 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6690
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   450
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "House Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4770
      TabIndex        =   1
      Top             =   450
      Width           =   1695
   End
End
Attribute VB_Name = "frmHouseSt"
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

'*****************************************************************************
'subroutine will be executed when the house combo box is clicked to select item
'*****************************************************************************
Private Sub cmbHouse_Click()

    If cmbHouse.ListIndex <> -1 Then
        lblSnapCount.Caption = ""
        lblCribbageCount.Caption = ""
        lblSpillikinsCount.Caption = ""
        lblScrabbleCount.Caption = ""
        lblSnapWon.Caption = ""
        lblCribbageWon.Caption = ""
        lblSpillikinsWon.Caption = ""
        lblScrabbleWon.Caption = ""
        lblTournaments.Caption = ""
        
        'the statistics function will be called with appropriate house
        'and database attribute order
        'colours change based on the house combo box value
        If cmbHouse.Text = "Green" Then
            Me.BackColor = &HC0FFC0
            Frame1.BackColor = &HC0FFC0
            cmdOk.BackColor = &HC0FFFF
            lblSnapCount.Caption = Statistics("Green", 1)
            lblCribbageCount.Caption = Statistics("Green", 2)
            lblSpillikinsCount.Caption = Statistics("Green", 3)
            lblScrabbleCount.Caption = Statistics("Green", 4)
            lblSnapWon.Caption = Statistics("Green", 5)
            lblCribbageWon.Caption = Statistics("Green", 6)
            lblSpillikinsWon.Caption = Statistics("Green", 7)
            lblScrabbleWon.Caption = Statistics("Green", 8)
            'call Tournamet won counting function
            lblTournaments.Caption = getTournaments("Green")
        ElseIf cmbHouse.Text = "Yellow" Then
            Me.BackColor = &HC0FFFF
            Frame1.BackColor = &HC0FFFF
            cmdOk.BackColor = &HC0FFC0
            lblSnapCount.Caption = Statistics("Yellow", 1)
            lblCribbageCount.Caption = Statistics("Yellow", 2)
            lblSpillikinsCount.Caption = Statistics("Yellow", 3)
            lblScrabbleCount.Caption = Statistics("Yellow", 4)
            lblSnapWon.Caption = Statistics("Yellow", 5)
            lblCribbageWon.Caption = Statistics("Yellow", 6)
            lblSpillikinsWon.Caption = Statistics("Yellow", 7)
            lblScrabbleWon.Caption = Statistics("Yellow", 8)
            'call Tournamet won counting function
            lblTournaments.Caption = getTournaments("Yellow")
        End If
    End If
End Sub

'*****************************************************************************
'a function that reads the number of tournaments won by strHouse from the database
'input: House name: strHouse
'return the number of tournaments won by strHouse
'******************************************************************************
Private Function getTournaments(strHouse As String) As Integer
    gadoCommand.CommandText = "SELECT * FROM Tournaments" & _
        " WHERE House='" & strHouse & "'"
    Set gadoRecordSet = gadoCommand.Execute
    getTournaments = Val(gadoRecordSet.Fields(1))
End Function

'*****************************************************************************
'subroutine will be executed when the Ok command button is clicked
'*****************************************************************************
Private Sub cmdOk_Click()
    frmMain.Show
    Unload Me
End Sub

'*****************************************************************************
'subroutine will be executed when the form is loaded
'*****************************************************************************
Private Sub Form_Load()
    cmbHouse.AddItem "Green"
    cmbHouse.AddItem "Yellow"
    'pre-select green
    cmbHouse.ListIndex = 0
    Me.BackColor = &HC0FFC0
End Sub

'*****************************************************************************
'a function that calculate the sum of a specifi field in the database
'input: Class number: intClass
'input: the attribute location within the database (column position):intCol
'return the sum of the field
'******************************************************************************
Private Function Statistics(strHouse As String, intCol As Integer) As Integer
    Dim intX As Integer
    'integer value used to accumulate the values of the col
    Dim intCount As Integer
    
    'get the first record
    gadoCommand.CommandText = "SELECT * FROM Statistics WHERE House= '" _
    & strHouse & "'"
    Set gadoRecordSet = gadoCommand.Execute
     'intialise the accumulator
    intCount = Val(gadoRecordSet.Fields(intCol))
    
    'read the rest of the records
    For intX = 1 To 35
        gadoCommand.CommandText = "SELECT * FROM Statistics WHERE House= '" _
        & strHouse & "'"
        Set gadoRecordSet = gadoCommand.Execute
        gadoRecordSet.GetRows (intX)
        'if no more records belong to intClass exit the for loop
        If gadoRecordSet.EOF = True Then Exit For
        'add the current record field value to the accumulator
        intCount = intCount + Val(gadoRecordSet.Fields(intCol))
    Next intX
    'return the result
    Statistics = intCount
End Function
