Public Class MindfulnessAndArtTherapyWorkshops

'Declare global variables
Dim ageGroup As String 
Dim groupSize As Integer 
Dim numberOfWorkshops As Integer 
Dim weeklySchedule As Boolean 

'Declare constants
Const MIN_AGE_REQUIREMENT As Integer = 10 
Const MAX_AGE_REQUIREMENT As Integer = 21 

Public Sub Initialise()
' code to set initial values of global variables
ageGroup = "Teenagers"
groupSize = 10 
numberOfWorkshops = 2 
weeklySchedule = True 
End Sub

Public Function AgeEligibilitySelection(ByVal ageInput As Integer) As Boolean 
' code to validate age eligibility
If ageInput >= MIN_AGE_REQUIREMENT And ageInput <= MAX_AGE_REQUIREMENT Then
Return True 
Else 
Return False
End If 
End Function 

Public Sub CollectInputRequests(ByRef ageInput As Integer)
' code to collect age input request from user
 ageInput = InputBox("Please enter the age of the participant:", "Age Eligibility") 
End Sub

Public Sub FinaliseSetup() 
'Code to finalise the setup for the workshops
Dim participantsValidated As Boolean = False 
Dim ageInput As Integer 

Do 
CollectInputRequests(ageInput)
participantsValidated = AgeEligibilitySelection(ageInput)
Loop Until participantsValidated = True 

MsgBox("Program setup completed successfully.") 
End Sub

Public Sub RunWorkshops() 
' code to run the workshops
Dim workshop As Integer 

For workshop = 1 To numberOfWorkshops
If weeklySchedule = True Then
MsgBox("The next workshop will start next week.") 
Else 
MsgBox("The next workshop will start in two weeks.") 
End If 

Dim artSession As String 
artSession = InputBox("Please select the type of art session for the workshop:", "Art Session Selection") 

Dim mindfulnessSession As String 
mindfulnessSession = InputBox("Please select the type of mindfulness session for the workshop:", "Mindfulness Session Selection") 

MsgBox("The workshop will now begin.") 

Next 

End Sub

End Class