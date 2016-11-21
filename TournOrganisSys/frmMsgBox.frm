VERSION 5.00
Begin VB.Form frmMsgBox 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attention"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   7875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
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
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3750
      TabIndex        =   1
      Top             =   840
      Width           =   105
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************Tournament Organising System************************
'***********************************Message Box Form***************************
'*******************************Programer: S. Saqfelhait***********************
'***********************************Date:07/06/2006****************************
'******************************************************************************
'this form is loaded from myMsgBox function "mdlTOS"
Option Explicit

'*****************************************************************************
'this subroutine will be executed the OK button is clicked
'*****************************************************************************
Private Sub cmdOk_Click()
    'store the response based on the type of the button yes,ok,..
    If msgResponse = "OKCancel" Then
        msgResponse = "OK"
    ElseIf msgResponse = "YesNo" Then
        msgResponse = "Yes"
    ElseIf msgResponse = "OK" Then
        msgResponse = "OK"
    ElseIf msgResponse = "Yes" Then
        msgResponse = "Yes"
    End If
    
    Unload Me
End Sub

'*****************************************************************************
'this subroutine will be executed the cancel button is clicked
'*****************************************************************************
Private Sub cmdCancel_Click()
    If msgResponse = "OKCancel" Then
        msgResponse = "Cancel"
    ElseIf msgResponse = "YesNo" Then
        msgResponse = "No"
    End If
    Unload Me
End Sub

'*****************************************************************************
'this subroutine will be executed the form is loaded into memory "RAM"
'*****************************************************************************
Private Sub Form_Load()
    'show the buttons on the form based on the value passed by myMsgBox func
    If msgResponse = "OKCancel" Then
        cmdCancel.Visible = True
        cmdOk.Visible = True
        cmdOk.Caption = "&Ok"
        cmdCancel.Caption = "&Cancel"
    ElseIf msgResponse = "YesNo" Then
        cmdCancel.Visible = True
        cmdOk.Visible = True
        cmdOk.Caption = "&Yes"
        cmdCancel.Caption = "&No"
    ElseIf msgResponse = "Ok" Then
        cmdCancel.Visible = False
        cmdOk.Visible = True
        cmdOk.Left = (frmMsgBox.ScaleWidth / 2) - (frmMsgBox.cmdOk.Width / 2)
        cmdOk.Caption = "&Ok"
    ElseIf msgResponse = "Yes" Then
        cmdCancel.Visible = False
        cmdOk.Visible = True
        cmdOk.Left = (frmMsgBox.ScaleWidth / 2) - (frmMsgBox.cmdOk.Width / 2)
        cmdOk.Caption = "&Yes"
    End If
  
End Sub

