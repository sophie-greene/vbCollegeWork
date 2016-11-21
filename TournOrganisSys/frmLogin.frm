VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3225
   ClientLeft      =   2580
   ClientTop       =   3015
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905.437
   ScaleMode       =   0  'User
   ScaleWidth      =   7816.726
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   1695
      Left            =   1155
      TabIndex        =   6
      Top             =   397
      Width           =   6015
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   885
         Width           =   1440
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2332
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2317
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4612
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2317
      Width           =   1380
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************** Tournament Organising System************************
'**********************************frmLogin Code********************************
'****************************Programer: Somoud Saqfelhait***********************
'***********************************Date:07/04/2007*****************************
'*******************************************************************************
'this is the login form loaded from the frmSplash form
Option Explicit

'*******************************************************************************
'subroutine,executed when the Cancel button is clicked
'*******************************************************************************
Private Sub cmdCancel_Click()
    End
End Sub

'*******************************************************************************
'subroutine,executed when the OK button is clicked
'*******************************************************************************
Private Sub cmdOk_Click()
    'check for username and correct password
    If UCase(txtUserName.Text) = "ADMIN" Then
        If txtPassword = "aloha" Then
            DB_Connect ("students.mdb")
            frmMain.Show
            Unload Me
        Else
            myMsgBox "Invalid Password, try again!", "OK", "Login"
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
        End If
    Else
        myMsgBox "Invalid UserName, try again!", "OK", "Login"
        txtUserName.SetFocus
            'SendKeys "{Home}+{End}"
    End If
End Sub

