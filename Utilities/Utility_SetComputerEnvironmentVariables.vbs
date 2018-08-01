'001. SetEnvironmentVariables

Call SetEnvironmentVariables()

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	SetEnvironmentVariables

'Description			 :	Function Used to Set Environment Variables

'Parameters			   :   NA
'
'Return Value		   : 	NA

'Pre-requisite			:	NA

'Examples				:  Call SetEnvironmentVariables()
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												04-Sep-2014								1.0										New						
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function SetEnvironmentVariables()
	On Error Resume next
	Dim UserVar,sType,sEnvName
	Dim objShell,objUsrEnv
	UserVar = ""

	Set objShell = CreateObject("WScript.Shell")
	sType=InputBox("Enter Variable Type  ( User Or System )","Enter Variable Type","User")
	
	If sType="" Then
		Msgbox "Variable Type cannot be blank"
		Exit Function
	End If
	
	Set objUsrEnv = objShell.Environment(sType)
	
	sEnvName=InputBox("Enter Variable Name","Enter Variable Name")
	If sEnvName="" Then
		Msgbox "Variable Name cannot be blank"
		sEnvName=InputBox("Enter Variable Name","Enter Variable Name")
		If sEnvName="" Then
			objShell.Popup "Fail to set " & sType & " level variable as user passed invalid ( empty ) variable name",20,"Fail to set variable values"	
			Exit function
		End IF
	End If
	
	UserVar=InputBox("Enter Variable Value","Enter Variable Value")
	If UserVar="" Then
		Msgbox "Variable value cannot be blank"
		UserVar=InputBox("Enter Variable Value","Enter Variable Value")
		If UserVar="" Then
			objShell.Popup "Fail to set " & sType & " level variable as user passed invalid ( empty ) variable value",20,"Fail to set variable values"	
			Exit function
		End IF
	End If
	objUsrEnv(sEnvName)=UserVar


	If Err.Number<>0 Then
		objShell.Popup "Fail to set " & sType & " level variable",20,"Fail to set variable values"
	Else
		objShell.Popup "Successfully  set " & sType & " level variable " & sEnvName & " to value " & UserVar ,10,"Successfully set variable values"
	End If

	Set objUsrEnv = Nothing
	Set objShell = Nothing
End Function