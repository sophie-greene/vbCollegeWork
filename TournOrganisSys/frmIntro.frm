VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00C0FFC0&
   Caption         =   "TOS"
   ClientHeight    =   8235
   ClientLeft      =   3165
   ClientTop       =   2460
   ClientWidth     =   14220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   4455
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Next >>"
      Default         =   -1  'True
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
      Left            =   11400
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter the number of &matches"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   720
      TabIndex        =   9
      Top             =   840
      Width           =   6375
      Begin VB.ComboBox cmbCribbage 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1845
         Width           =   3000
      End
      Begin VB.ComboBox cmbSpillikins 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2715
         Width           =   3000
      End
      Begin VB.ComboBox cmbScrabble 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3600
         Width           =   3000
      End
      Begin VB.ComboBox cmbSnap 
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
         ItemData        =   "frmIntro.frx":0000
         Left            =   3000
         List            =   "frmIntro.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   3000
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sc&rabble matches:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   2400
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "S&pillikins matches:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   2715
         Width           =   2400
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Crib&bage matches:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   1845
         Width           =   2280
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Snap matches:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   10
         Top             =   960
         Width           =   2160
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Choose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7560
      TabIndex        =   8
      Top             =   1800
      Width           =   6015
      Begin VB.OptionButton optManual 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Let me choose the competitors"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   5760
      End
      Begin VB.OptionButton optRandom 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Generate the competitors list randomly "
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
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Value           =   -1  'True
         Width           =   5760
      End
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
      Height          =   735
      Left            =   720
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create Games Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4523
      TabIndex        =   15
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************** Tournament Organising System************************
'**********************************frmIntro Code********************************
'****************************Programer: Somoud Saqfelhait***********************
'***********************************Date:07/04/2007*****************************
'*******************************************************************************
'this form will be loaded from the main menu form when Game button is clicked
Option Explicit

'*****************************************************************************
'subroutine will be executed when the "Go to Main window" button is clicked
'*****************************************************************************
Private Sub cmdMain_Click()
     frmMain.Show
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
'subroutine will be executed when the Next command button is clicked
'*****************************************************************************
Private Sub cmdNext_Click()
    'check for validity of number of matches intered
    If Not (IsNumeric(cmbSnap.Text) And IsNumeric(cmbCribbage.Text) _
        And IsNumeric(cmbSpillikins.Text) And IsNumeric(cmbScrabble.Text)) Then
        
            myMsgBox "Invalid input,please enter a number between 1 and 18" _
            , "Ok", "Attention"
            Exit Sub
    End If
    'check if entered number are in the range 1-18
    If Not ((1 <= Val(cmbSnap.Text) Or Val(cmbSnap.Text) <= 18) _
        And (1 <= Val(cmbCribbage.Text) And Val(cmbCribbage.Text) <= 18) _
        And (1 <= Val(cmbSpillikins.Text) And Val(cmbSpillikins.Text) <= 18) _
        And (1 <= Val(cmbScrabble.Text) And Val(cmbScrabble.Text) <= 18)) Then
        
            myMsgBox "Invalid number,please enter a number between 1 and 18" _
            , "Ok", "Attention"
            Exit Sub
    End If
    If (Val(cmbSnap.Text) + Val(cmbCribbage.Text) + _
        Val(cmbSpillikins.Text) + Val(cmbScrabble.Text)) Mod 2 = 0 Then
            myMsgBox "Please ensure that the sum of the match numbers is odd" _
            , "Ok", "Attention"
            Exit Sub
    End If
    'store the seleceted numbers of matches
    gintSnapMatch = Val(cmbSnap.Text)
    gintCribbageMatch = Val(cmbCribbage.Text)
    gintSpillikinsMatch = Val(cmbSpillikins.Text)
    gintScrabbleMatch = Val(cmbScrabble.Text)
 
    'load random or manual form based on the option button selected
    If gboolRandom = True Then
        frmRandom.Show
        'Me.Hide
        Unload Me
    Else
        frmManual.Show
        'Me.Hide
        Unload Me
    End If
    
End Sub

'*****************************************************************************
'this subroutine will be executed when the form is loaded
'*****************************************************************************
Private Sub Form_Load()
    
    Dim intX As Integer
    'fix form position relative to the screen
    Left = Screen.Width / 2 - Width / 2
    Top = Screen.Height / 2 - Height / 2 - 2000
    'initialise the combo boxes
    For intX = 1 To 18
        cmbSnap.AddItem intX
        cmbCribbage.AddItem intX
        cmbSpillikins.AddItem intX
        cmbScrabble.AddItem intX
    Next intX
    gboolRandom = True
    gboolEdit = False
End Sub

'*****************************************************************************
'this subroutine will be executed the mouse is moved over frame2
'*****************************************************************************
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

    optRandom.ForeColor = vbBlack
    optRandom.FontBold = False
    optManual.ForeColor = vbBlack
    optManual.FontBold = False
End Sub

'*****************************************************************************
'this subroutine will be executed the mouse is moved over Manual option button
'*****************************************************************************
Private Sub optManual_MouseMove(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

    optManual.ForeColor = vbRed
    optManual.FontBold = True
End Sub

'*****************************************************************************
'this subroutine will be executed the mouse is moved over Random option button
'*****************************************************************************
Private Sub optRandom_MouseMove(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

    optRandom.ForeColor = vbRed
    optRandom.FontBold = True
End Sub

'*****************************************************************************
'this subroutine will be executed the Manual option button is clicked
'*****************************************************************************
Private Sub optManual_Click()
    gboolRandom = False
End Sub

'*****************************************************************************
'this subroutine will be executed the Random option button is clicked
'*****************************************************************************
Private Sub optRandom_Click()
    gboolRandom = True
End Sub
