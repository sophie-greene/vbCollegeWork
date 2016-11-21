VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Help"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13275
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&OK"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000040&
      Height          =   5175
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   11535
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************Tournament Organising System************************
'***********************************frmHelp Code*******************************
'*******************************Programer: S. Saqfelhait***********************
'***********************************Date:07/06/2006****************************
'******************************************************************************
'this form is loaded from the main menu form when the help button is clicked
Option Explicit

'*****************************************************************************
'this subroutine is executed when the OK button is clicked
'*****************************************************************************
Private Sub cmdOk_Click()
    Unload Me
End Sub
'*****************************************************************************
'this subroutine is executed when the form is loaded
'*****************************************************************************
Private Sub Form_Load()
    Dim strFileName As String
    Dim intX As Integer
    Dim strHelpText As String
    
    'get the full path of the help file
    strFileName = App.Path & "\help.txt"
    'open the help file in read mode
    Open strFileName For Input As #1
    'read first line
    Input #1, strHelpText
    lblHelp.Caption = strHelpText
    'read all the lines of the file until the end of the file is reached
    Do While Not EOF(1)
        Input #1, strHelpText
        lblHelp.Caption = lblHelp.Caption & vbCrLf & strHelpText
    Loop
    Close #1
End Sub
