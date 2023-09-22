Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'The enumeration lists the types of questions the program can ask.
Private Enum QuestionTypeE
   Addition = 0     'Addition.
   Division         'Division.
   Multiplication   'Multiplication.
   Subtraction      'Subtraction.
End Enum

'This structure defines a question and its answer.
Private Type QuestionStr
   Answer As String         'Defines the answer.
   Text As String           'Defines the question's text.
End Type

'This structure defines a session of questions.
Private Type SessionStr
   CorrectCount As Long     'Defines the number of questions correctly answered by the user.
   Difficulty As Long       'Defines the difficulty set for the questions.
   IncorrectCount As Long   'Defines the number of questions incorrectly answered by the user.
   QuestionCount As Long    'Defines the number of questions asked during a session.
End Type

Private Const QUESTION_TYPE_COUNT As Long = 4   'Defines the number of different question types.
'This procedure displays the results of a session once it has finished.
Private Sub DisplayResults(Session As SessionStr)
On Error GoTo ErrorTrap
Dim Message As String

   With Session
      Message = "All questions have been answered." & vbCr
      Message = Message & CStr(Session.CorrectCount) & " out of " & CStr(.QuestionCount) & " questions "
      If .QuestionCount = 1 Then Message = Message & " was" Else Message = Message & " were"
      Message = Message & " answered correctly." & vbCr
      Message = Message & "Difficulty: " & .Difficulty
   End With
   
   MsgBox Message, vbInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure returns the operator for the specified question type.
Private Function GetOperatorSymbol(QuestionType As QuestionTypeE) As String
On Error GoTo ErrorTrap
Dim Symbol As String

   Symbol = vbNullString
   Select Case QuestionType
      Case Addition
         Symbol = "+"
      Case Division
         Symbol = "/"
      Case Multiplication
         Symbol = "*"
      Case Subtraction
         Symbol = "-"
   End Select

EndRoutine:
   GetOperatorSymbol = Symbol
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure generates and returns a random question based on the difficulty setting.
Private Function GetQuestion() As QuestionStr
On Error GoTo ErrorTrap
Dim LeftNumber As Long
Dim Question As QuestionStr
Dim QuestionType As QuestionTypeE
Dim RightNumber As Long
 
   With Question
      QuestionType = Int(Rnd * QUESTION_TYPE_COUNT)
      If QuestionType = Division Then
         Do
            LeftNumber = CLng(Rnd * (GetHighestNumber() + 1))
            RightNumber = CLng(Rnd * GetHighestNumber()) + 1
            If LeftNumber Mod RightNumber = 0 Then Exit Do
            DoEvents
         Loop
      Else
         LeftNumber = CLng(Rnd * (GetHighestNumber() + 1))
         RightNumber = CLng(Rnd * (GetHighestNumber() + 1))
      End If
      
      Select Case QuestionType
         Case Addition
            .Answer = CStr(LeftNumber + RightNumber)
         Case Division
            .Answer = CStr(LeftNumber / RightNumber)
         Case Multiplication
            .Answer = CStr(LeftNumber * RightNumber)
         Case Subtraction
            .Answer = CStr(LeftNumber - RightNumber)
      End Select
      
      .Text = "What is the answer to: " & CStr(LeftNumber) & " " & GetOperatorSymbol(QuestionType) & " " & CStr(RightNumber) & "?"
   End With
EndRoutine:
   GetQuestion = Question
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the number of questions specified by the user.
Private Function GetQuestionCount(DefaultCount As Long) As Long
On Error GoTo ErrorTrap
Dim Answer As String
Dim NewCount As Long

   Do
      Answer = Trim$(InputBox$("Number of questions (1-100)?", , CStr(DefaultCount)))
      If Answer = vbNullString Then
         If Quit(Ask:=True) Then Exit Do
      End If
      
      NewCount = 0
      If IsValidIntegralNumber(Answer) Then NewCount = CLng(Val(Answer))
      
      If NewCount > 1 And NewCount < 100 Then
         Exit Do
      Else
         MsgBox "The number of questions can range from 1 to 100.", vbExclamation
      End If
      DoEvents
   Loop
EndRoutine:
   GetQuestionCount = NewCount
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function



'This procedure returns the difficulty level specified by the user.
Private Function GetDifficulty(DefaultDifficulty As Long) As Long
On Error GoTo ErrorTrap
Dim Answer As String
Dim NewDifficulty As Long

   Do
      Answer = Trim$(InputBox$("Difficulty level (1-10)?", , CStr(DefaultDifficulty)))
      If Answer = vbNullString Then
         If Quit(Ask:=True) Then Exit Do
      End If
      
      NewDifficulty = 0
      If IsValidIntegralNumber(Answer) Then NewDifficulty = CLng(Val(Answer))
      
      If NewDifficulty >= 1 And NewDifficulty <= 10 Then
         Exit Do
      Else
         MsgBox "The difficulty level can range from 1 to 10.", vbExclamation
      End If
      DoEvents
   Loop
   
EndRoutine:
   GetDifficulty = NewDifficulty
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


 
'This procedure returns the highest number allowed by the set difficulty level.
Function GetHighestNumber(Optional NewDifficulty As Long = 0) As Long
On Error GoTo ErrorTrap
Static Highest As Long

   If Not NewDifficulty = 0 Then
      Highest = NewDifficulty * 10
   End If

EndRoutine:
   GetHighestNumber = Highest
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure initializes and returns a session.
Private Function GetSession() As SessionStr
On Error GoTo ErrorTrap
Dim Session As SessionStr
   
   Randomize
   
   With Session
      .CorrectCount = 0
      .IncorrectCount = 0
      .QuestionCount = GetQuestionCount(DefaultCount:=10)
      If Not Quit() Then .Difficulty = GetDifficulty(DefaultDifficulty:=1)
      GetMaximum NewDifficulty:=.Difficulty
   End With
   
EndRoutine:
   GetSession = Session
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function



'This procedure returns the status text for the specified session.
Private Function GetStatus(Session As SessionStr) As String
Dim Status As String

   With Session
      Status = "Correct: " & CStr(.CorrectCount)
      Status = Status & "   Incorrect: " & CStr(.IncorrectCount)
      Status = Status & "   Question: " & CStr(.CorrectCount + .IncorrectCount + 1) & "/" & CStr(.QuestionCount)
      Status = Status & "   " & "Difficulty: " & CStr(.Difficulty)
   End With
EndRoutine:
   GetStatus = Status
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles any errors that occur.
Private Sub HandleError()
Dim Choice As Long
Dim ErrorCode As Long
Dim Message As String

   ErrorCode = Err.Number
   Message = Err.Description
   On Error Resume Next
   Choice = MsgBox(Message & vbCr & "Error code: " & ErrorCode, vbInformation Or vbOKCancel)
   If Choice = vbCancel Then End
End Sub


'This procedure checks whether or not the specified text represents a valid integral number and returns the result.
Private Function IsValidIntegralNumber(Text As String) As Boolean
On Error GoTo ErrorTrap
Dim IsValid As Boolean
   
   IsValid = (CStr(CLng(Val(Text))) = Text)
   
EndRoutine:
   IsValidIntegralNumber = IsValid
   Exit Function
   
ErrorTrap:
   IsValid = False
   HandleError
   Resume EndRoutine
End Function


'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
Dim Answer As String
Dim Question As QuestionStr
Dim Session As SessionStr
 
   Session = GetSession()
   Do Until Quit()
      If Session.CorrectCount + Session.IncorrectCount = Session.QuestionCount Then
         DisplayResults Session
         If Quit(Ask:=True) Then Exit Do
         Session = GetSession()
      End If
      
      Question = GetQuestion()
      Answer = Trim$(InputBox$(GetStatus(Session) & vbCr & vbCr & Question.Text))
      If Answer = vbNullString Then
         If Quit(Ask:=True) Then Exit Do
      Else
         If Answer = Question.Answer Then
            Session.CorrectCount = Session.CorrectCount + 1
         Else
            Session.IncorrectCount = Session.IncorrectCount + 1
            MsgBox "Incorrect, the correct answer is: " & Question.Answer & ".", vbExclamation
         End If
      End If
      DoEvents
   Loop
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub
'This procedure returns whether or not the user has indicated the program should be quit.
Private Function Quit(Optional Ask As Boolean = False) As Boolean
On Error GoTo ErrorTrap
Static Answer As Long
   If Ask Then
      Answer = MsgBox("Quit?", vbQuestion Or vbYesNo Or vbDefaultButton2)
   End If
EndRoutine:
   Quit = (Answer = vbYes)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


