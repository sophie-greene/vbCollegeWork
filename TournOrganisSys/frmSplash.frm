VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8550
   ClientLeft      =   2925
   ClientTop       =   2370
   ClientWidth     =   14265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   14265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAnimate 
      Interval        =   2000
      Left            =   480
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   6315
      Left            =   2100
      TabIndex        =   0
      Top             =   1118
      Width           =   10065
      Begin VB.Timer tmrSplash 
         Left            =   240
         Top             =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright: June 2007"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8040
         TabIndex        =   5
         Top             =   5760
         Width           =   1770
      End
      Begin VB.Image imgLogo 
         Height          =   945
         Left            =   120
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   8520
         TabIndex        =   2
         Top             =   5400
         Width           =   1140
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   3000
         Width           =   1290
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   2040
         TabIndex        =   4
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label lblLicenseTo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "licenced to:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   960
         TabIndex        =   1
         Top             =   3960
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'************************** Tournament Organising System************************
'***********************************frmSplash Code******************************
'*******************************Programer: S. Saqfelhait************************
'***********************************Date:07/04/2007*****************************
'*******************************************************************************
'startup form
Option Explicit

'*******************************************************************************
'subroutine,executed if any key is pressed
'input: the operating system pass the ascii value of the key pressed
'*******************************************************************************
Private Sub Form_KeyPress(KeyAscii As Integer)
    'if a key is pressed the introfrm will be shown
    '& the form will be unloaded
    frmLogin.Show
    Unload Me
End Sub

'*****************************************************************************
'subroutine will be executed when the form is loaded
'*****************************************************************************
Private Sub Form_Load()
    'set the timer to 5 seconds
    tmrSplash.Interval = 5000
    'enable the timer control
    tmrSplash.Enabled = True
    'allow the operating system to deal with other events other than the timeout
    DoEvents
    lblVersion.Caption = "Version " & App.Major & "." & _
        App.Minor & "." & App.Revision
    lblProductName.Caption = "Tournament Organising System (TOS)"
    lblLicenseTo.Caption = " Licenced to Salchaster Primary school"
    lblCompany.Caption = " S. Saqfelhait"
End Sub

'*****************************************************************************
'subroutine will be executed when frame1 is clicked
'*****************************************************************************
Private Sub Frame1_Click()
    frmLogin.Show
    Unload Me
End Sub

'*****************************************************************************
'subroutine will be executed when the timer is enabled and every  "interval"
'*****************************************************************************
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

'*****************************************************************************
'subroutine will be executed when the timer is enabled and every  "interval"
'*****************************************************************************
Private Sub tmrSplash_Timer()
    'this subroutine is exceuted after _
    'the timer is enabled by the interval value
    frmLogin.Show
    Unload Me
    tmrSplash.Enabled = False
End Sub
