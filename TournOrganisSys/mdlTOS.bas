Attribute VB_Name = "mdlTOS"
'*******************************************************************************
'************************** Tournament Organising System************************
'**********************************modTOS Code**********************************
'*******************************Programer: S. Saqfelhait************************
'***********************************Date:07/04/2007*****************************
'this module contains functions and procedures that can be called from any form
'within the project
'*******************************************************************************

'enforce declaration of all variables used within the module
Option Explicit

'store customised msgbox response
Public msgResponse As String

'a global variable used to store the winning house
Public gstrWinner As String

'a global variable to store the way in which the user choose to generate the games
'true if games to be generated randomly otherwise false
Public gboolRandom As Boolean

'a global variable to indicate that the game is being edited
Public gboolEdit As Boolean

'a global variable used to store the number of snap matches
Public gintSnapMatch As Integer

'a global variable used to store the number of cribbage matches
Public gintCribbageMatch As Integer

'a global variable used to store the number of spillikins matches
Public gintSpillikinsMatch As Integer

'a global variable used to store the number of scrabble matches
Public gintScrabbleMatch As Integer

'define a variable to connect to Access Database
Public gadoConnection As New ADODB.Connection

'define a variable to hold SQL commands
Public gadoCommand As New ADODB.Command

'define a variable to execute SQL commands and to access Access records
Public gadoRecordSet As ADODB.Recordset

'define a variable to hold the connection path
Public gstrConnection As String

'define a variable to hold the number of records in the database
Public gintCountRecord As Integer

'a global variable to hold the game schedule, 2-dimensional array
'even colums contains Green house oponents
'odd columns contain Yellow house oponents
'Snap Green:0,Yellow:1
'Cribbage Green:2,Yellow:3
'Spllikins Green:4,Yellow:5
'Scrabble Green:6,Yellow:7
Public strGame(18, 8) As String

'a global variable to hold the green house studets names and classes,
'2-dimensional array
Public strGstudent(18, 2) As String

'a global variable to hold the yellow house students names and classes,
'2-dimensional array
Public strYstudent(18, 2) As String

'*******************************************************************************
'a function that will open a connection with a database
'input: file name
'*******************************************************************************
Public Sub DB_Connect(strFile As String)
    gstrConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
        App.Path & "\" & strFile
        
    gadoConnection.Open gstrConnection
    Set gadoCommand.ActiveConnection = gadoConnection
  
End Sub

'*******************************************************************************
'a subroutine that will close the connection with a database
'*******************************************************************************
Public Sub DB_Disconnect()

    Set gadoRecordSet = Nothing
    Set gadoCommand = Nothing
 
    gadoConnection.Close
    Set gadoConnection = Nothing
End Sub

'*******************************************************************************
'a subroutine that reads the names of a specific house students
'input: house name:hs
'output: array of students names: str
'*******************************************************************************
Public Sub getStudentArray(str() As String, hs As String)
   
    Dim intCount As Integer
    Dim intX As Integer
    intX = 0
    'count number of students in hs
    gadoCommand.CommandText = _
            "SELECT COUNT (*) as recordc FROM Students WHERE House='" & hs & "'"
    Set gadoRecordSet = gadoCommand.Execute
    'store the count in intCount
    intCount = gadoRecordSet("recordc")
    'read first record
    gadoCommand.CommandText = _
        "SELECT * FROM Students WHERE House='" & hs & "'"
    Set gadoRecordSet = gadoCommand.Execute
    'store it
    str(intX, 0) = gadoRecordSet.Fields(1) 'name
    str(intX, 1) = gadoRecordSet.Fields(2) 'class
    
    'read the rest of the records
    For intX = 1 To intCount - 1
        gadoCommand.CommandText = _
            "SELECT * FROM Students WHERE House='" & hs & "'"
        Set gadoRecordSet = gadoCommand.Execute
        gadoRecordSet.GetRows (intX) 'get the record number intX
        str(intX, 0) = gadoRecordSet.Fields(1) 'name
        str(intX, 1) = gadoRecordSet.Fields(2) 'class
    Next intX
      
End Sub

'*******************************************************************************
'a subroutine that clears game information from the Access database
'*******************************************************************************
Public Sub clearDatabase()

    gadoCommand.CommandText = "UPDATE Students  Set PlayedSnap=false"
    Set gadoRecordSet = gadoCommand.Execute
    gadoCommand.CommandText = "UPDATE Students  Set SnapOponent=0"
    Set gadoRecordSet = gadoCommand.Execute
    gadoCommand.CommandText = "UPDATE Students  Set PlayedCribbage=false"
    Set gadoRecordSet = gadoCommand.Execute
    gadoCommand.CommandText = "UPDATE Students  Set CribbageOponent=0"
    Set gadoRecordSet = gadoCommand.Execute
    gadoCommand.CommandText = "UPDATE Students  Set PlayedScrabble=false"
    Set gadoRecordSet = gadoCommand.Execute
    gadoCommand.CommandText = "UPDATE Students  Set ScrabbleOponent=0"
    Set gadoRecordSet = gadoCommand.Execute
    gadoCommand.CommandText = "UPDATE Students  Set PlayedSpillikins=false"
    Set gadoRecordSet = gadoCommand.Execute
    gadoCommand.CommandText = "UPDATE Students  Set SpillikinsOponent=0"
    Set gadoRecordSet = gadoCommand.Execute
    
End Sub

'*******************************************************************************
'a function that checks if a specific student has played a specific game
'input: student name: str
'input: indicates gameand student house: intGameCol
'return true if str has played the game which col is intGameCol otherwise false
'*******************************************************************************
Public Function IsInGame(str As String, intGameCol As Integer) As Boolean

    Dim intX As Integer
   
    'compare the student name str to all students who played already if a match
    'is found true is returned and the search is terminated
    For intX = 0 To 17
        If strGame(intX, intGameCol) = str Then
            IsInGame = True
            Exit Function
        End If
    Next intX
    'if str did not match any of the students who already played
    'str did not play and a false is returned
    IsInGame = False
    
End Function

'*******************************************************************************
'a function that checks if two students form the yellow and green houses
'has played each other in any of the four games
'input: first student name: str1
'input: second student name: str2
'return true if str1 has already played str2
'*******************************************************************************
Public Function Played(str1 As String, str2 As String, intGameCol As Integer) As Boolean
    
    Dim intX As Integer
    Dim intY As Integer
    If intGameCol Mod 2 = 0 Then
        For intX = 0 To 7 'go through all games
            If intX Mod 2 = 0 Then
                For intY = 0 To 17 'check all matches
                'if the first student is found in the game schedule
                'and if the oponent is the second student-> str1 and str2
                'played each other and true is returned and the search is terminated
                    If strGame(intY, intX) = str1 Then
                        If strGame(intY, intX + 1) = str2 Then
                            Played = True
                            Exit Function
                        End If
                    End If
                Next intY
            End If
        Next intX
    Else
        For intX = 0 To 7 'go through all games
            If intX Mod 2 = 1 Then
                For intY = 0 To 17 'check all matches
                'if the first student is found in the game schedule
                'and if the oponent is the second student-> str1 and str2
                'played each other and true is returned and the search is terminated
                    If strGame(intY, intX) = str1 Then
                        If strGame(intY, intX - 1) = str2 Then
                            Played = True
                            Exit Function
                        End If
                    End If
                Next intY
            End If
        Next intX
    End If
    
    'if the search ended without finding matches false is returned
    Played = False
    
End Function

'*******************************************************************************
'a function that generate a list of all potential yellow house oponents for a
'student from the green house in a specific game
'input: student name: str
'input: indicates game and student house: intGameCol
'output: list of the names of the potential oponents
'return the number of potential oponents
'*******************************************************************************
Public Function potentialOponents(str As String, _
intGameCol As Integer, oponents() As String) As Integer

'integer to hold "str" class
    Dim intClass As Integer
    Dim intX As Integer
    'number of openents
    potentialOponents = 0
    'Clear oponents array
    For intX = LBound(oponents) To UBound(oponents)
        oponents(intX) = ""
    Next intX
    'game number is odd the yellow house student is in consideration
    'yellow
    If intGameCol Mod 2 = 1 Then
        'find the class the student is in
        For intX = 0 To 17
            If strGstudent(intX, 0) <> "" Then
                If strGstudent(intX, 0) = str Then
                    intClass = strGstudent(intX, 1)
                    Exit For
                End If
            End If
        Next intX
        'search through the list to find student who can play
        For intX = 0 To 17
            If strYstudent(intX, 0) <> "" Then
                If strYstudent(intX, 1) <> intClass Then
                    If Not IsInGame(strYstudent(intX, 0), intGameCol) Then
                        If Not Played(strYstudent(intX, 0), str, intGameCol) Then
                            oponents(potentialOponents) = strYstudent(intX, 0)
                            potentialOponents = potentialOponents + 1
                        End If
                    End If
                End If
            End If
       Next intX
    'green
    Else
         'find the class the student is in
      For intX = 0 To 17
           If strYstudent(intX, 0) <> "" Then
                If strYstudent(intX, 0) = str Then
                    intClass = strYstudent(intX, 1)
                    Exit For
                End If
            End If
       Next intX
        'search through the list
        For intX = 0 To 17
            If strGstudent(intX, 0) <> "" Then
                If strGstudent(intX, 1) <> intClass Then
                    If Not IsInGame(strGstudent(intX, 0), intGameCol) Then
                        If Not Played(strGstudent(intX, 0), str, intGameCol) Then
                            oponents(potentialOponents) = strGstudent(intX, 0)
                            potentialOponents = potentialOponents + 1
                        End If
                    End If
               End If
            End If
         Next intX
    End If

End Function
'*******************************************************************************
'a function that clears the game array
'*******************************************************************************
Public Sub clearGameArray()
    Dim intX As Integer
    Dim intY As Integer
    
    For intX = 0 To 7
        For intY = 0 To 17
            strGame(intY, intX) = ""
        Next intY
    Next intX

End Sub

'*******************************************************************************
'a function that generate a list of possible players added to a combo box
'input: the name of the game: strGame
'input: the combo box that will hold the potential players names
'output: a list of all potential players at that point will be entered into
'the combo box
'*******************************************************************************
Public Sub Manual(cmb1 As ComboBox, txtOp As String, intIndex As Integer _
, intGameCol As Integer)

    Dim op(18) As String
    Dim intOp As Integer
    Dim intX As Integer
    Dim strTemp As String
    'remove the current player from game
    If gboolEdit = True Then
        strTemp = ""
    Else
        strTemp = cmb1.Text
    End If
    
    'update the database (remove the player)
    gadoCommand.CommandText = "UPDATE Students Set PlayedCribbage=False " _
    & "where studentName='" & strTemp & "'"
    Set gadoRecordSet = gadoCommand.Execute
    strGame(intIndex, intGameCol) = ""
    
    cmb1.Clear
    'generate a list of all possible players in that place
    intOp = potentialOponents(txtOp, intGameCol, op())
    
    'add the players to the combo box list
    For intX = 0 To intOp - 1
        cmb1.AddItem op(intX)
    Next intX
End Sub

'*******************************************************************************
'a function that increments the number of matches played by a student in a game
'input: the name of the game: strGame
'input:the name of the student: strName
'the process is carried out on the database
'*******************************************************************************
Public Sub UpdateMatchCount(strName As String, strGame As String)

    'define a variable to hold the previous number of matches
    Dim intTempCount As Integer
     'get the previous value from the database
    gadoCommand.CommandText = "SELECT * FROM statistics" & _
        " WHERE StudentName='" & strName & "'"
    Set gadoRecordSet = gadoCommand.Execute
    
    'store the value depending on the game considered
    If strGame = "Snap" Then
        intTempCount = gadoRecordSet.Fields(1)
    ElseIf strGame = "Cribbage" Then
        intTempCount = gadoRecordSet.Fields(2)
    ElseIf strGame = "Spillikins" Then
        intTempCount = gadoRecordSet.Fields(3)
    ElseIf strGame = "Scrabble" Then
        intTempCount = gadoRecordSet.Fields(4)
    Else
        myMsgBox "Error", "Ok", "Error"
    End If
    'increment the games played and update the database
    intTempCount = intTempCount + 1
     gadoCommand.CommandText = "UPDATE statistics  Set " & strGame & "Matches=" _
       & intTempCount & " where studentName='" & strName & "'"
      Set gadoRecordSet = gadoCommand.Execute
      
End Sub

'*******************************************************************************
'a function that increments the number of matches won by a student in a game
'it also store the max score the student achieved
'input: the name of the game: strGame
'input:the name of the student: strName
'input:the score of the student, intScore
'the process is carried out on the database
'*******************************************************************************
Public Sub UpdateMatchWonCount(strName As String, strGame As String, _
intScore As Integer)
    'define a variable to hold the previous number of matches
    Dim intTempCount As Integer
    Dim intTempScore As Integer
    'get the previous value from the database
    gadoCommand.CommandText = "SELECT * FROM statistics" & _
        " WHERE StudentName='" & strName & "'"
    Set gadoRecordSet = gadoCommand.Execute
    
    'store the value depending on the game considered
    If strGame = "Snap" Then
        intTempCount = Val(gadoRecordSet.Fields(5))
        intTempScore = Val(gadoRecordSet.Fields(9))
    ElseIf strGame = "Cribbage" Then
        intTempCount = Val(gadoRecordSet.Fields(6))
        intTempScore = Val(gadoRecordSet.Fields(10))
    ElseIf strGame = "Spillikins" Then
        intTempCount = Val(gadoRecordSet.Fields(7))
        intTempScore = Val(gadoRecordSet.Fields(11))
    ElseIf strGame = "Scrabble" Then
        intTempCount = Val(gadoRecordSet.Fields(8))
        intTempScore = Val(gadoRecordSet.Fields(12))
    Else
        myMsgBox "Error", "Ok", "Error"
    End If
    
    'increment the games won and update the database
    intTempCount = Val(intTempCount) + 1
    gadoCommand.CommandText = "UPDATE statistics  Set Won" & strGame & "=" _
    & intTempCount & " where studentName='" & strName & "'"
    Set gadoRecordSet = gadoCommand.Execute
    
    'if current score is greater than the score stored in the database,store it
    If intScore > intTempScore Then
        gadoCommand.CommandText = "UPDATE statistics  Set " & strGame & "Max =" _
    & intScore & " where studentName='" & strName & "'"
    Set gadoRecordSet = gadoCommand.Execute
    End If
    
End Sub

'*******************************************************************************
'a function that increments the number of compititions won by a specific house
'input: the name of the house: strWinner
'the process is carried out on the database
'*******************************************************************************
Public Sub UpdateTournaments(strWinner As String)
  'define a variable to hold the previous number of tournaments won
    Dim intTempNumber As Integer
    
    'read the previous number
    gadoCommand.CommandText = "SELECT * FROM Tournaments" & _
        " WHERE House='" & strWinner & "'"
    Set gadoRecordSet = gadoCommand.Execute
    
    intTempNumber = gadoRecordSet.Fields(1)
    'increament the number
    intTempNumber = intTempNumber + 1
    
    'update the database
    gadoCommand.CommandText = "UPDATE Tournaments  Set CompititionsWon=" _
    & intTempNumber & " where House='" & strWinner & "'"
    Set gadoRecordSet = gadoCommand.Execute
End Sub

'*******************************************************************************
'a public function use frmMsgBox to diplay a Message
'input: the message contetnts: strMessage
'input:indication of the buttons that should be on the message form: strOption
'input:the title or caption of the message form:strTitle
'*******************************************************************************
Public Function myMsgBox(strMessage As String, strOption As String, _
strTitle As String) As Variant
    'pass the options to the global variable used in the message form
    msgResponse = strOption
    frmMsgBox.lblMessage.Caption = strMessage
    frmMsgBox.Caption = strTitle
    frmMsgBox.Show (1)
    'store the response of the user
    If msgResponse = "Yes" Then
        myMsgBox = vbYes
    ElseIf msgResponse = "No" Then
        myMsgBox = vbNo
    ElseIf msgResponse = "OK" Then
        myMsgBox = vbOK
    ElseIf msgResponse = "Cancel" Then
        myMsgBox = vbCancel
    End If
End Function


