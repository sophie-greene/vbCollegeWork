VERSION 5.00
Begin VB.Form frmWinner 
   BackColor       =   &H00000000&
   Caption         =   "Congratulation"
   ClientHeight    =   8235
   ClientLeft      =   2970
   ClientTop       =   1875
   ClientWidth     =   14220
   ControlBox      =   0   'False
   DrawWidth       =   3
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   145.256
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   250.825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMain 
      Caption         =   "Go to Main Window"
      Height          =   495
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Timer tmrTrophy 
      Interval        =   100
      Left            =   1080
      Top             =   840
   End
   Begin VB.Timer tmrExplosion 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   120
   End
   Begin VB.Image picTrophy 
      Height          =   4935
      Left            =   4800
      Picture         =   "frmWinner.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "frmWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************** Tournament Organising System************************
'**********************************frmWinner Code*******************************
'****************************Programer: Somoud Saqfelhait***********************
'***********************************Date:07/06/2007*****************************
'*******************************************************************************
'this form will be loaded from the random form when winner command button is clicked

Option Explicit
Dim gErase As String
'circle radius
Dim intR As Integer

Private Sub cmdMain_Click()
    Unload Me
    frmMain.Show
End Sub

'*****************************************************************************
'subroutine will be executed when the form is activated
'*****************************************************************************
Private Sub Form_Activate()
    gErase = "Winner"
    picTrophy.Left = (Me.ScaleWidth - picTrophy.Width) / 2
    picTrophy.Top = (Me.ScaleHeight - picTrophy.Height) / 2
    picTrophy.Width = 1
    picTrophy.Height = 1
End Sub

'*****************************************************************************
'subroutine will be executed when the form is de-activated
'*****************************************************************************
Private Sub Form_Deactivate()
    tmrExplosion.Enabled = False
    
End Sub

'*****************************************************************************
'subroutine will be executed when the form is loaded
'*****************************************************************************
Private Sub Form_Load()
    If gstrWinner = "Green" Then
        ForeColor = vbGreen
        'load the picture of the trophy with green rippon
        picTrophy.Picture = LoadPicture(App.Path & "\greentrophy.jpg")
        cmdMain.BackColor = &HC0FFFF
    ElseIf gstrWinner = "Yellow" Then
        ForeColor = vbYellow
        cmdMain.BackColor = &HC0FFC0
        'load the picture of the trophy with yellow rippon
        picTrophy.Picture = LoadPicture(App.Path & "\yellowtrophy.jpg")
    Else
        ForeColor = &HC0FFC0
    End If
    Caption = Caption + " for the " + gstrWinner + " House"

End Sub

'*****************************************************************************
'subroutine will be executed when the form is clicked
'*****************************************************************************
Sub Form_Click()
  Unload Me
  frmMain.Show
End Sub

'*****************************************************************************
'subroutine will be executed when the timer is enabled and every  "interval"
'*****************************************************************************
Private Sub tmrExplosion_Timer()
'if no circles draw else erase
 If gErase = "Winner" Then
        drawCircle 50, 50, intR
        drawCircle 210, 50, intR
        drawCircle 40, 40, intR + 1
        drawCircle 200, 40, intR + 1
        drawCircle 200, 100, intR
        drawCircle 50, 100, intR
        drawCircle 210, 90, intR
        drawCircle 60, 90, intR
        intR = intR + 2
    ElseIf gErase = "Var" Then
        varCircle 50, 50, intR
        varCircle 210, 50, intR
        varCircle 40, 40, intR + 1
        varCircle 200, 40, intR + 1
        varCircle 200, 100, intR
        varCircle 50, 100, intR
        varCircle 210, 90, intR
        varCircle 60, 90, intR
        intR = intR + 2
    Else
        eraseCircle 50, 50, intR
        eraseCircle 210, 50, intR
        eraseCircle 40, 40, intR + 1
        eraseCircle 200, 40, intR + 1
        eraseCircle 200, 100, intR
        eraseCircle 50, 100, intR
        eraseCircle 210, 90, intR
        eraseCircle 60, 90, intR
        intR = intR + 2
    End If
    'when the max radius is reached change status from erase to draw
    'or from draw to erase
    If intR >= 30 Then
        intR = 0
        If gErase = "Winner" Then
            gErase = "Var"
        ElseIf gErase = "Var" Then
            gErase = "Erase"
        Else
            gErase = "Winner"
        End If
    End If
End Sub

'*****************************************************************************
'subroutine will be executed when the timer is enabled and every  "interval"
'*****************************************************************************
Private Sub tmrTrophy_Timer()
'while the width is less than 80 increase the pic dimensions
    If picTrophy.Width > 80 Then
        tmrExplosion.Enabled = True
        tmrTrophy.Enabled = False
    End If
    picTrophy.Width = picTrophy.Width + 6
    picTrophy.Height = picTrophy.Height + 8
    'keep the pic centered
    picTrophy.Left = (Me.ScaleWidth - picTrophy.Width) / 2
    picTrophy.Top = (Me.ScaleHeight - picTrophy.Height) / 2
End Sub

'*****************************************************************************
'subroutine, draw a dotted circle with the current forecolor
'input:X- cordinate of the center of the circle:intX
'input:Y- cordinate of the center of the circle:intY
'input:the raduis of the circle:intRad
'*****************************************************************************
Private Sub drawCircle(intX As Integer, intY As Integer, intRad As Integer)
    Dim intZ As Double

    For intZ = -intRad To intRad Step 0.9
        CurrentX = intZ + intX
        CurrentY = Abs(Sqr(Abs((intRad * intRad) - (intZ * intZ)))) + intY
        frmWinner.PSet (CurrentX, CurrentY)
    Next intZ
    For intZ = -intRad To intRad Step 0.9
        CurrentX = intZ + intX
        CurrentY = -Abs(Sqr(Abs((intRad * intRad) - (intZ * intZ)))) + intY
        frmWinner.PSet (CurrentX, CurrentY)
    Next intZ
End Sub

'*****************************************************************************
'subroutine, draw a dotted circle with random colour
'input:X- cordinate of the center of the circle:intX
'input:Y- cordinate of the center of the circle:intY
'input:the raduis of the circle:intRad
'*****************************************************************************
Private Sub varCircle(intX As Integer, intY As Integer, intRad As Integer)
    Dim intZ As Double
    Dim temp As Variant
    temp = ForeColor
    ForeColor = QBColor((Rnd * 10) + 5)
    For intZ = -intRad To intRad Step 0.9
        CurrentX = intZ + intX
        CurrentY = Abs(Sqr(Abs((intRad * intRad) - (intZ * intZ)))) + intY
        frmWinner.PSet (CurrentX, CurrentY)
    Next intZ
    For intZ = -intRad To intRad Step 0.9
        CurrentX = intZ + intX
        CurrentY = -Abs(Sqr(Abs((intRad * intRad) - (intZ * intZ)))) + intY
        frmWinner.PSet (CurrentX, CurrentY)
    Next intZ
    ForeColor = temp
End Sub

'*****************************************************************************
'subroutine, draw a dotted circle with the current backcolor i.e erase any
'circle with the same coordinates
'input:X- cordinate of the center of the circle:intX
'input:Y- cordinate of the center of the circle:intY
'input:the raduis of the circle:intRad
'*****************************************************************************
Private Sub eraseCircle(intX As Integer, intY As Integer, intRad As Integer)
    Dim intZ As Double
    Dim temp As Variant
    temp = ForeColor
    ForeColor = BackColor
    For intZ = -intRad To intRad Step 0.9
        CurrentX = intZ + intX
        CurrentY = Abs(Sqr(Abs((intRad * intRad) - (intZ * intZ)))) + intY
        frmWinner.PSet (CurrentX, CurrentY)
    Next intZ
    For intZ = -intRad To intRad Step 0.9
        CurrentX = intZ + intX
        CurrentY = -Abs(Sqr(Abs((intRad * intRad) - (intZ * intZ)))) + intY
        frmWinner.PSet (CurrentX, CurrentY)
    Next intZ
    ForeColor = temp
End Sub
