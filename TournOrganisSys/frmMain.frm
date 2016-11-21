VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Main Form"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAnimate 
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   7095
      Left            =   383
      TabIndex        =   6
      Top             =   570
      Width           =   13455
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
         Height          =   615
         Left            =   983
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11183
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CommandButton cmdstClasses 
         BackColor       =   &H00C0FFC0&
         Caption         =   "C&lasses' Achievement Statistics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4943
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "click to view classes acheviement statistics"
         Top             =   3960
         Width           =   4335
      End
      Begin VB.CommandButton cmdstHouse 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Houses' &Achievement Statistics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4943
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "click to view houses acheviement statistics"
         Top             =   3000
         Width           =   4335
      End
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Games"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4943
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Click to create Games schedule randomly or manually, eneter results and acknowledge the winner"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CommandButton cmdStatistics 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Students' Achievement Statistics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4943
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Click to view each student's achievement in all games"
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Window"
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
         Left            =   4920
         TabIndex        =   7
         Top             =   360
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************** Tournament Organising System************************
'**********************************frmMain Code*********************************
'****************************Programer: Somoud Saqfelhait***********************
'***********************************Date:07/04/2007*****************************
'*******************************************************************************
'this is the main menu form loaded on successful login
Option Explicit

'*******************************************************************************
'subroutine,executed when the Classes Achievment Statistics button is clicked
'*******************************************************************************
Private Sub cmdstClasses_Click()
    frmClassSt.Show
    Unload Me
End Sub
'*******************************************************************************
'subroutine,executed when the Houses Achievment Statistics button is clicked
'*******************************************************************************
Private Sub cmdstHouse_Click()
    frmHouseSt.Show
    Unload Me
End Sub

'*******************************************************************************
'this subroutine is exceuted after the timer is enabled by the interval value
'*******************************************************************************
Private Sub tmrAnimate_Timer()
    'toggle form and frame colours
    If Me.BackColor = &HC0FFC0 Then
        Me.BackColor = &HC0FFFF
    Else
        Me.BackColor = &HC0FFC0
    End If
    If Frame1.BackColor = &HC0FFC0 Then
        Frame1.BackColor = &HC0FFFF
    Else
        Frame1.BackColor = &HC0FFC0
    End If
End Sub

'*******************************************************************************
'subroutine,executed when the Games button is clicked
'*******************************************************************************
Private Sub cmdCreate_Click()
    frmIntro.Show
    Unload Me
End Sub
'*******************************************************************************
'subroutine,executed when the Exit button is clicked
'*******************************************************************************
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

'*******************************************************************************
'subroutine,executed when the Help button is clicked
'*******************************************************************************
Private Sub cmdHelp_Click()
    frmHelp.Show (1)
End Sub

'*******************************************************************************
'subroutine,executed when the Students Achievment Statistics button is clicked
'*******************************************************************************
Private Sub cmdStatistics_Click()
    frmStatistics.Show
    Unload Me
End Sub

