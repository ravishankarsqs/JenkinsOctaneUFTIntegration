Option Explicit
Err.Clear

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Function Name									|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. 	Fn_CommonUtil_EnvironmentVariablesOperations	|	Sandeep.Navghane@sqs.com	|	16-Jan-2015	|	Function used to perform operations on local machine\systems environment variables 
'002. 	Fn_CommonUtil_DataTableOperations				|	vrushali.sahare@sqs.com		|	24-Feb-2016	|	Function used to perfrom operations on datatable
'003. 	Fn_CommonUtil_KeyBoardOperation					|	vrushali.sahare@sqs.com		|	25-Feb-2016	|	Function used to perfrom the keypress function on selected node
'004. 	Fn_CommonUtil_GetCursorStateState				|	vrushali.sahare@sqs.com		|	25-Feb-2016	|	Function used to get cursor current state
'005. 	Fn_CommonUtil_CursorReadyStatusOperation		|	vrushali.sahare@sqs.com		|	25-Feb-2016	|	Function used to perform operation on cursor ready status
'006. 	Fn_CommonUtil_WindowsApplicationOperations		|	vrushali.sahare@sqs.com		|	25-Feb-2016	|	Function used to perform operations on running processes
'007. 	Fn_CommonUtil_GenerateRandomNumber				|	Sandeep.Navghane@sqs.com	|	09-Mar-2016	|	Function used to generate random number
'008. 	Fn_CommonUtil_GenerateRandomString				|	Sandeep.Navghane@sqs.com	|	09-Mar-2016	|	Function used to Generate Random String of given length
'009. 	Fn_CommonUtil_MouseWheelRotationOperations		|	Sandeep.Navghane@sqs.com	|	09-Mar-2016	|	Function used to scroll mouse wheel up/down
'010. 	Fn_CommonUtil_StringArrayOperations				|	Sandeep.Navghane@sqs.com	|	09-Mar-2016	|	Function used to perform operations on string array
'011. 	Fn_CommonUtil_LocalMachineOperations			|	Sandeep.Navghane@sqs.com	|	09-Mar-2016	|	Function used to perform operations on local computer
'012. 	Fn_CommonUtil_AddDivider						|	Sandeep.Navghane@sqs.com	|	10-Mar-2016	|	Function used to add divider
'013. 	Fn_CommonUtil_ArrayStringContains				|	Sandeep.Navghane@sqs.com	|	13-Sep-2016	|	Function used to verify specific value contains in string array
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_CommonUtil_EnvironmentVariablesOperations
'
'Function Description	 :	Function used to perform operations on local machine\systems environment variables 
'
'Function Parameters	 :  1.sAction		: Function action name to perform
'							2.sVariableType	: Environment variable type
'							3.sVariableName	: Environment variable name
'							4.sVariableValue: Environment variable value
'
'Function Return Value	 : 	True \ False \ Environment Variable value
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Environment variable should exist
'
'Function Usage		     :  bReturn = Fn_CommonUtil_EnvironmentVariablesOperations("Set","User","AutomationDir","C:\GOG_AL_11.2")
'Function Usage		     :  bReturn = Fn_CommonUtil_EnvironmentVariablesOperations("Get","User","AutomationDir","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Jan-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Replacement for function : Fn_GetEnvValue,Fn_SetEnvValue ' Delete this comment once implementation is completed
Function Fn_CommonUtil_EnvironmentVariablesOperations(sAction,sVariableType,sVariableName,sVariableValue)
	'Declaring variables
	Dim objShell,objEnvironment
	Dim sVariableCurrentValue
	
	'Initially set function return value as False
	Fn_CommonUtil_EnvironmentVariablesOperations=False
	
	'Creating [ Shell ] object
	Set objShell = CreateObject("WScript.Shell")
	'Creating Environment Type object
	Set objEnvironment = objShell.Environment(sVariableType)
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set local machine environment variable
		Case "Set"
			objEnvironment(sVariableName) = sVariableValue
			Fn_CommonUtil_EnvironmentVariablesOperations=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get local machine environment variable value
		Case "Get"
			sVariableCurrentValue = objEnvironment(sVariableName)
			If sVariableCurrentValue<>"" Then
				Fn_CommonUtil_EnvironmentVariablesOperations=sVariableCurrentValue
			End If
	End Select
	
	If Err.Number <> 0 Then
		'Fn_CommonUtil_EnvironmentVariablesOperations = False
		'Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_CommonUtil_EnvironmentVariablesOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Releasing above created objects
	Set objEnvironment = Nothing	
	Set objShell = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_DataTableOperations
'
'Function Description	 :	Function used to perfrom operations on datatable
'
'Function Parameters	 :   1.sAction: Action name 
'						    2.sColumnName: Column name in datatable
'							3.sValue: Value to be set in datatable
'							4.iRowNumber: Datatable rownumber where data is to be stored
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_CommonUtil_DataTableOperations("IsColumnExist","User_Role","","")
'Function Usage		     :	Call Fn_CommonUtil_DataTableOperations("AddColumn","User_Group","","")
'Function Usage		     :	Call Fn_CommonUtil_DataTableOperations("SetValue","DesignID","WS_200369",2)
'Function Usage		     :	Call Fn_CommonUtil_DataTableOperations("getvalue","DesignID","",2)
'Function Usage		     :	Call Fn_CommonUtil_DataTableOperations("exportdatatable","","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  24-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CommonUtil_DataTableOperations(sAction, sColumnName, sValue, iRowNumber)
	'Declaring Variables
	Dim iCounter
	Dim bFlag
    Err.Clear

	'Initially set function return value as False
	Fn_CommonUtil_DataTableOperations = False
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to check existance of Column in global sheet datatable
		Case "iscolumnexist"
			For iCounter = 1 To DataTable.GlobalSheet.GetParameterCount
				If DataTable.GlobalSheet.GetParameter(iCounter).Name = sColumnName Then
					Fn_CommonUtil_DataTableOperations = True
					Exit For
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Add Column in global sheet datatable
		Case "addcolumn"
			bFlag = False
			For iCounter = 1 To DataTable.GlobalSheet.GetParameterCount
				If DataTable.GlobalSheet.GetParameter(iCounter).Name = sColumnName Then
					bFlag = True
					Exit For
				End If
			Next
			If bFlag = False Then
				DataTable.GlobalSheet.AddParameter sColumnName,sValue
			End If
			Fn_CommonUtil_DataTableOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set value in global sheet datatable
		Case "setvalue"
			If iRowNumber <> "" Then
				DataTable.GlobalSheet.SetCurrentRow iRowNumber
			End If
			DataTable.Value(sColumnName,"Global") = sValue
			Fn_CommonUtil_DataTableOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get value in global sheet datatable
		Case "getvalue"
			If iRowNumber <> "" Then
				DataTable.GlobalSheet.SetCurrentRow iRowNumber
			End If
			Fn_CommonUtil_DataTableOperations = DataTable.Value(sColumnName,"Global")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to export global sheet datatable
		Case "exportdatatable"
			If sValue = "" Then
				sValue = Fn_Setup_GetAutomationFolderPath("TestData")
			End If
			DataTable.ExportSheet  sValue & "\" &  Environment.Value("TestName") & ".xls","Global"
			Fn_CommonUtil_DataTableOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get column index
		Case "getcolumnindex"
			For iCounter = 1 To DataTable.GlobalSheet.GetParameterCount
				If DataTable.GlobalSheet.GetParameter(iCounter).Name = sColumnName Then
					Fn_CommonUtil_DataTableOperations = iCounter
					Exit For
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to object create security matrix import global sheet datatable
		Case "importobjectcreatesecuritymatrixdatatable"
			sValue = Fn_Setup_GetAutomationFolderPath("TestData")			
			DataTable.ImportSheet sValue & "\SecurityMatrix\ObjectCreateSecurityMatrix.xlsx","CreateSecurityMatrix","Global"
			Fn_CommonUtil_DataTableOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Exit Function
	End Select
	If Err.Number <> 0 Then
		Fn_CommonUtil_DataTableOperations = False
	End If
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_KeyBoardOperation
'
'Function Description	 :	Function used to perfrom the keypress function on selected node
'
'Function Parameters	 :   1.sAction: Action name 
'						   	 2.sKey: Key Name
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_CommonUtil_KeyBoardOperation("SendKey", "^(c)")
'Function Usage		     :	Call Fn_CommonUtil_KeyBoardOperation("SendKeys", "^(c)~^(v)") 'Use ~ as seperator. Do not use ~ for {ENTER}
'Function Usage		     :	Call Fn_CommonUtil_KeyBoardOperation("PressKey", 28) '28 = Enter
'Function Usage		     :	Call Fn_CommonUtil_KeyBoardOperation("PressKeyAndSendString", "%~T~r~p")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  25-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CommonUtil_KeyBoardOperation(sAction, sKey)
	'Declaring Variables
	Dim aKey, aKeySet
	Dim iCount
	Dim WshShell
	Dim objDeviceReplay
	
	'Initially set function return value as False
	Fn_CommonUtil_KeyBoardOperation = False
	
	'Following cases used to send key stroke to the active window
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "sendkey"
			If Fn_UI_Object_Operations("Fn_CommonUtil_KeyBoardOperation","Exist", JavaWindow("DefaultWindow"),"","","") Then
				JavaWindow("DefaultWindow").Click 150, 3, "LEFT"
				'Creating Shell object
				Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys sKey
				wait(GBL_MIN_MICRO_TIMEOUT)
				'Release Shell Object
				Set WshShell = Nothing
				If Err.Number <> 0 Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ Fn_CommonUtil_KeyBoardOperation ] : Failed to Send Keystroke [" + sKey + "] on teamcenter application due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]" )
					Exit function
				Else	
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "PASS >> Successfully Operated Keystroke [" + sKey + "] On Teamcenter Application")
				End If
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>: [ Fn_CommonUtil_KeyBoardOperation ] : Teamcenter Application Window Not Found" )
				Exit function
			End If			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "presskey"
			aKey = Split(sKey, ":", -1, 1)
			If Fn_UI_Object_Operations("Fn_CommonUtil_KeyBoardOperation","Exist", JavaWindow("DefaultWindow"),"","","") Then
				JavaWindow("DefaultWindow").Click 150, 3, "LEFT"
				If ubound(aKey) > 0 Then
					JavaWindow("DefaultWindow").PressKey aKey(0), aKey(1)
				Else
					JavaWindow("DefaultWindow").PressKey aKey(0)
				End If
				If Err.Number <> 0 Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>: [ Fn_CommonUtil_KeyBoardOperation ] : Failed to Send Keystroke [" + sKey + "] on teamcenter application due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]" )
					Exit function
				Else	
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "PASS >> Successfully Operated Keystroke [" + sKey + "] On Teamcenter Application")
				End If
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>: [ Fn_CommonUtil_KeyBoardOperation ] : Teamcenter Application Window Not Found" )
				Exit function
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "sendkeys"
			aKeySet = split(sKey,"~")
			'Creating Shell Object
			Set WshShell = CreateObject("WScript.Shell")
			For iCount = 0 to UBound(aKeySet)
				WshShell.SendKeys aKeySet(iCount)								
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "PASS >> Successfully Operated Keystroke [" + aKeySet(iCount) + "] On Teamcenter Application")
				Wait(GBL_MICRO_TIMEOUT)
			Next
			'Release Shell Object
			Set WshShell = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "presskeyandsendstring"
			aKeySet = split(sKey,"~")
			Select Case aKeySet(0)
				Case "%","LEFT ALT"
					aKeySet(0)="56"
			End Select
			'Creating Device Replay Object
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.PressKey aKeySet(0)
			Wait GBL_MICRO_TIMEOUT
			For iCount = 1 to UBound(aKeySet)
				objDeviceReplay.SendString aKeySet(iCount)
				Wait GBL_MICRO_TIMEOUT	
			Next
			'Release Object
			Set objDeviceReplay = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ Fn_CommonUtil_KeyBoardOperation ] : [No valid case] : No valid case was passed for function [Fn_CommonUtil_KeyBoardOperation]")
	End Select
	
	If Err.Number <> 0 Then
		 Fn_CommonUtil_KeyBoardOperation = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ Fn_CommonUtil_KeyBoardOperation ] : Due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	Else	
		Fn_CommonUtil_KeyBoardOperation = True
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_GetCursorStateState
'
'Function Description	 :	Function used to get cursor current state
'
'Function Parameters	 :  NA
'
'Function Return Value	 : 	Cursor current state number
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn = Fn_CommonUtil_GetCursorState()
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  25-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CommonUtil_GetCursorState()
	'Declaring Variables
	Dim iWindowHandle,iProcessID,iThreadID
	
	extern.Declare micLong,"GetForegroundWindow","user32.dll","GetForegroundWindow"
	extern.Declare micLong,"AttachThreadInput","user32.dll","AttachThreadInput", micLong, micLong,micLong
	extern.Declare micLong,"GetWindowThreadProcessId","user32.dll","GetWindowThreadProcessId", micLong, micLong
	extern.Declare micLong,"GetCurrentThreadId","kernel32.dll","GetCurrentThreadId"
	extern.Declare micLong,"GetCursor","user32.dll","GetCursor"

	iWindowHandle = extern.GetForegroundWindow()

	iProcessID = extern.GetWindowThreadProcessId(iWindowHandle, NULL)
	iThreadID = extern.GetCurrentThreadId()
	extern.AttachThreadInput iProcessID,iThreadID,True

	Fn_CommonUtil_GetCursorState = Eval("extern.GetCursor")

	extern.AttachThreadInput iProcessID,iThreadID,False
End function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_CursorReadyStatusOperation
'
'Function Description	 :	Function used to perfrom operation on cursor ready status
'
'Function Parameters	 :  1.iCursorState	: Cursor state
'						   	2.iIterations	: Number of iterations
'                           3.sCondition	: Condition to check
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn = Fn_CommonUtil_CursorReadyStatusOperation("","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  25-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CommonUtil_CursorReadyStatusOperation(iCursorState, iIterations, sCondition)
 	'Declaring variables
	Dim iCount,iCounter
	
	'Initially set function return value as False
	Fn_CommonUtil_CursorReadyStatusOperation = False
	
	'Set the cursor state
	If iCursorState = "" Then
	   iCursorState = "65539"
	End If
	
	'Set the number of iterations
	If iIterations = "" Then
		iIterations = 1
	End If
	
	'Set the condition to check
	If sCondition = "" Then
		sCondition = "not equal"
	End If

	Select Case sCondition
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case for cursor ready status sync
		Case "not equal"
			For iCount = 1 to iIterations
				For iCounter = 1 To 25
					If CStr(Fn_CommonUtil_GetCursorState()) = CStr(iCursorState) Then
						Fn_CommonUtil_CursorReadyStatusOperation = True
						Exit For
					Else
						Wait GBL_MICRO_TIMEOUT
					End If
				Next
				If Fn_CommonUtil_CursorReadyStatusOperation Then
					Exit for
				End If
			Next
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_CommonUtil_WindowsApplicationOperations
'
'Function Description	:	Function used to perform operations on running processes
'
'Function Parameters	:   1.sAction		: Action name 
'						    2.sApplication	: Application Name
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_CommonUtil_WindowsApplicationOperations("IsRunning", "EXCEL.EXE")
'Function Usage		     :	Call Fn_CommonUtil_WindowsApplicationOperations("Terminate", "EXCEL.EXE")
'Function Usage		     :	Call Fn_CommonUtil_WindowsApplicationOperations("TerminateAll", "EXCEL.EXE")
'Function Usage		     :	Call Fn_CommonUtil_WindowsApplicationOperations("TerminateAllExcel", "")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  25-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CommonUtil_WindowsApplicationOperations(sAction, sApplication)
	'Declaring Variables
	Dim objAllProcess,objProcess

	'Initially set function return value as False
	Fn_CommonUtil_WindowsApplicationOperations = False
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to check if the given process is running or not running state
		Case "isrunning"
			'Creating object of Process
			Set objAllProcess = getobject("winmgmts:") 
			'Get all the processes running in your PC
			For Each objProcess In objAllProcess.InstancesOf("Win32_process")
				If (Instr (Ucase(objProcess.Name),uCase(sApplication)) = 1) Then 
					Fn_CommonUtil_WindowsApplicationOperations = True
					Exit for
				End If
			Next
			If Fn_CommonUtil_WindowsApplicationOperations = True Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"PASS >> [ Fn_CommonUtil_WindowsApplicationOperations ] Application [ " & sApplication & " ] is running.")	
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"PASS >> [ Fn_CommonUtil_WindowsApplicationOperations ] Application [ " & sApplication & " ] is not running.")	
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to close all excels
		Case "terminateallexcel"
			SystemUtil.CloseProcessByName("Excel.exe")
			Fn_CommonUtil_WindowsApplicationOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to terminate process(s)
		Case "terminate", "terminateall"
			'Creating object of process
			Set objAllProcess = getobject("winmgmts:") 
			'Get all the processes running in your PC
			For Each objProcess In objAllProcess.InstancesOf("Win32_process")
				'Made all uppercase to remove ambiguity. Replace TASKMGR.EXE with your application name in CAPS.
				If (Instr(Ucase(objProcess.Name),uCase(sApplication)) = 1) Then 
					'You can replace this with Reporter.ReportEvent
					Call Fn_Setup_ReporterFilter("DisableAll")
					objProcess.Terminate
					Call Fn_Setup_ReporterFilter("EnableAll")
					wait 5
					Fn_CommonUtil_WindowsApplicationOperations = True
					If lcase(sAction) <> "terminateall" Then
						Exit for
					End If
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ Fn_CommonUtil_WindowsApplicationOperations ] :  [No valid case] : No valid case was passed for function [ Fn_CommonUtil_WindowsApplicationOperations ].")	
			Exit function
	End Select
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_CommonUtil_WindowsApplicationOperations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ Fn_CommonUtil_WindowsApplicationOperations ] : fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If	
	'Release Object
	Set objAllProcess = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_CommonUtil_GenerateRandomNumber
'
'Function Description	 :	Function used to generate random number 
'
'Function Parameters	 :  1.iLength : Random number lenght
'
'Function Return Value	 : 	Random number
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :  bReturn = Fn_CommonUtil_GenerateRandomNumber(3)
'Function Usage		     :  bReturn = Fn_CommonUtil_GenerateRandomNumber(7)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  09-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CommonUtil_GenerateRandomNumber(iLength)
	'Declaring variables
	Dim iRandomNumber,iStartNumber,iCounter
	Randomize
	
	iStartNumber="9"
	For iCounter=1 To iLength-1
		iStartNumber=Cstr(iStartNumber)+"0"
	Next
	iRandomNumber = Int((iStartNumber * Rnd) + 1)
	
	If Len(Cstr(iRandomNumber)) < iLength Then
		Fn_CommonUtil_GenerateRandomNumber = "6" + Cstr(iRandomNumber)
	Else
		Fn_CommonUtil_GenerateRandomNumber = Cstr(iRandomNumber)
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_GenerateRandomString
'
'Function Description	 :	Function used to Generate Random String of given length
'
'Function Parameters	 :  1.iLength		: Random String Length
'							2.sLetterCase	: Random String Letter case
'
'Function Return Value	 : 	Random String
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn = Fn_CommonUtil_GenerateRandomString(6,"")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  09-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_CommonUtil_GenerateRandomString(iLength,sLetterCase)
	'Declaring variables
	Dim sRandomString
	Dim iCounter
	
	Const sMainString= "abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyz"
	
	For iCounter = 1 to iLength
	  sRandomString=sRandomString & Mid(sMainString,RandomNumber(1,Len(sMainString)),1)
	Next
	
	If sLetterCase="Lower" Then
		Fn_CommonUtil_GenerateRandomString=Lcase(sRandomString)
	Else
		Fn_CommonUtil_GenerateRandomString=UCase(sRandomString)
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_MouseWheelRotationOperations
'
'Function Description	 :	Function used to scroll mouse wheel up/down
'
'Function Parameters	 :  1.iNumberOfRotation	: Number of times to rotate(-ve value for down, viceversa)
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn = Fn_CommonUtil_MouseWheelRotationOperations(6)
'Function Usage		     :	bReturn = Fn_CommonUtil_MouseWheelRotationOperations(-3)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  09-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_CommonUtil_MouseWheelRotationOperations(iNumberOfRotation)
	'Declaring variables
    Const MOUSEEVENTF_WHEEL = 2048 	'@const long 	| MOUSEEVENTF_WHEEL | middle button up
    Const POSWHEEL_DELTA = 120 		'@const long 	| POSWHEEL_DELTA 	| movement of 1 mousewheel click Down<nl>
    Const NEGWHEEL_DELTA = -120 	'@const long	| NEGWHEEL_DELTA 	| movement of 1 mousewheel click Up<nl>
    Dim iCounter
    dim sMovementPosition
	
    Extern.Declare micVoid,"mouse_event","user32.dll","mouse_event",micLong,micLong,micLong,micLong,micLong
    
    If iNumberOfRotation > 0 then
		sMovementPosition="Down"
    End if
	
    For iCounter = 1 to Abs(iNumberOfRotation)
		If sMovementPosition="Down" then
			Extern.mouse_event MOUSEEVENTF_WHEEL,0,0,POSWHEEL_DELTA,0
		Else
			Extern.mouse_event MOUSEEVENTF_WHEEL,0,0,NEGWHEEL_DELTA,0
		End if
    Next
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_StringArrayOperations
'
'Function Description	 :	Function used to perform operations on string array
'
'Function Parameters	 :  1.sAction		: Action to perform on array
'							2.aStringData	: String array
'							3.sOrder		: Array order
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn=Fn_CommonUtil_StringArrayOperations("Sort",aStringData,"Ascending")
'Function Usage		     :	bReturn=Fn_CommonUtil_StringArrayOperations("VerifyOrder",aStringData,"Ascending")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  09-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CommonUtil_StringArrayOperations(sAction,ByRef aStringData,sOrder)
	'Declaring variables
	Dim iCount,iCounter
	Dim sTempValue
	
	Select Case (sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "", "Sort"
			If IsArray(aStringData)=True Then
				If sOrder="Descending" Then
					for iCount = UBound(aStringData) - 1 To 0 Step -1
						For iCounter= 0 to iCount
							If aStringData(iCounter)<aStringData(iCounter+1) then 
								sTempValue=aStringData(iCounter+1) 
								aStringData(iCounter+1)=aStringData(iCounter)
								aStringData(iCounter)=sTempValue 
							End if 
						  Next 
					  Next 
				Else
					 for iCount = UBound(aStringData) - 1 To 0 Step -1
						For iCounter= 0 to iCount
							  If aStringData(iCounter)>aStringData(iCounter+1) then 
								  sTempValue=aStringData(iCounter+1) 
								  aStringData(iCounter+1)=aStringData(iCounter)
								  aStringData(iCounter)=sTempValue 
							  End if 
						  Next 
					  Next 
				End If
				Fn_CommonUtil_StringArrayOperations=True
			Else
				Fn_CommonUtil_StringArrayOperations=False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_CommonUtil_StringArrayOperations ] : Fail as given array is not standard format array due to which fail to sort array in [ " & Cstr(sOrder) & " ] oder")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyOrder"
			Fn_CommonUtil_StringArrayOperations = True
			If IsArray(aStringData)=True Then
				If sOrder="Descending" Then
					For iCounter= 0 to UBound(aStringData)-1
						If aStringData(iCounter) >aStringData(iCounter+1) OR aStringData(iCounter) =aStringData(iCounter+1) then 
							'Do Nothing
						Else
							Fn_CommonUtil_StringArrayOperations = False
							Exit For
						End if 
					Next 
				Else
					For iCounter= 0 to UBound(aStringData)-1
						If aStringData(iCounter) < aStringData(iCounter+1) OR aStringData(iCounter) =aStringData(iCounter+1) then 
							'Do Nothing
						Else
							Fn_CommonUtil_StringArrayOperations = False
							Exit For
						End If 
					Next 
				End If
			Else
				Fn_CommonUtil_StringArrayOperations=False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_CommonUtil_StringArrayOperations ] : Fail as given array is not standard format array due to which fail to sort array in [ " & Cstr(sOrder) & " ] oder")
			End If
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_LocalMachineOperations
'
'Function Description	 :	Function used to perform operations on local computer
'
'Function Parameters	 :  1.sAction		: Action to perform
'							2.sValue		: Value
'
'Function Return Value	 : 	False\Computer Name\User name
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn=Fn_CommonUtil_LocalMachineOperations("getcurrentloginusername","")
'Function Usage		     :	bReturn=Fn_CommonUtil_StringArrayOperations("getcomputername","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  09-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CommonUtil_LocalMachineOperations(sAction,sValue)
	'Variable Declaration
	Dim sNamingContext,sUserDN,sUserName
	Dim objRootDSE,objADSysInfo,objUser,objShell
	
	Fn_CommonUtil_LocalMachineOperations=False
	
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - 
		'Case to get current login user to the local machine
		Case "getcurrentloginusername"
			If Environment("UserName")="" Then
				Set objRootDSE = GetObject("LDAP://RootDSE")
				If Err.Number = 0 Then 
					sNamingContext = objRootDSE.Get("defaultNamingContext")  
				Else 
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_CommonUtil_LocalMachineOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
					Exit Function
				End If
				Set objADSysInfo = CreateObject("ADSystemInfo")
				sUserDN = objADSysInfo.username 
				Set objUser = Getobject("LDAP://" & sUserDN)
				sUserName = objUser.Get("givenName")  & " " & objUser.Get("sn")
				Set objUser = Nothing
				Set objADSysInfo = Nothing
				Set objRootDSE = Nothing
			Else
				sUserName=Environment("UserName")
			End If
			If sUserName<>"" Then
				Fn_CommonUtil_LocalMachineOperations=sUserName
			End If
		' - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - 
		'Case to get computer name
		Case "getcomputername"
			Set objShell = CreateObject( "WScript.Shell" )
			If objShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )<>"" Then
				Fn_CommonUtil_LocalMachineOperations= objShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
			End If
			Set objShell =Nothing
	End Select	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_AddDivider
'
'Function Description	 :	Function used to add divider
'
'Function Parameters	 :  1.sDivider		: Divider type
'							2.iDividerCount	: Divider count
'
'Function Return Value	 : 	Divider
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn=Fn_CommonUtil_AddDivider("Tab","3")
'Function Usage		     :	bReturn=Fn_CommonUtil_AddDivider("NewLine","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  10-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_CommonUtil_AddDivider(sDivider,iDividerCount)
	'Declaring variables
	Dim sTempDivider
	Dim iCounter
	If iDividerCount="" Then
		iDividerCount=0
	End If
	Select Case Lcase(sDivider)
		Case "tab"
			For iCounter = 0 to iDividerCount
				sTempDivider = sTempDivider & vbTab
			Next
			Fn_CommonUtil_AddDivider=sTempDivider
		Case "newline"
			For iCounter = 0 to iDividerCount				
				sTempDivider = sTempDivider & vblf
			Next
			Fn_CommonUtil_AddDivider=sTempDivider
	End Select	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CommonUtil_ArrayStringContains
'
'Function Description	 :	Function used to verify specific value contains in string array
'
'Function Parameters	 :  1.sString		: String Array
'							2.sValue		: String value
'							3.sSeparator	: Array seperator
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn=Fn_CommonUtil_ArrayStringContains("Ab;dc;zx;wt","zx",";")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  13-Sep-2016	    |	 1.0		|	Minal N			 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CommonUtil_ArrayStringContains(sString,sValue,sSeparator)
	'Declaring variables
	Dim iCounter
	Dim aValues
	Fn_CommonUtil_ArrayStringContains = False
	aValues = split(sString, sSeparator)
	For iCounter = 0 to UBound(aValues)
		If aValues(iCounter) = sValue Then
			Fn_CommonUtil_ArrayStringContains = True
			Exit for
		End If
	Next
End Function

Function Fn_CommonUtil_DateOperations(sAction,sDate,sFormat,iNumber)

'Declaring variables
	Dim aDateTime,aDate
	Dim iMonth
	
	Fn_CommonUtil_DateOperations=False
	
	Select Case sAction		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetPostDate"
			Select Case sFormat
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "DD-MMM-YYYY"	
					If iNumber="" Then
						iNumber=0
					ENd If
					aDateTime=Split(Now() + iNumber," ")
					aDate=Split(aDateTime(0),"/")
					iMonth=MonthName(aDate(0),True)
					If Len(aDate(1))=1 Then
						aDate(1)="0" & Cstr(aDate(1))
					End IF
					Fn_CommonUtil_DateOperations=aDate(1) & "-" & iMonth & "-" & aDate(2)
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "DD/MM/YYYY"	
					If iNumber="" Then
						iNumber=0
					ENd If
					aDateTime=Split(Now() + iNumber," ")
					aDate=Split(aDateTime(0),"/")
					iMonth=aDate(0)
					If Len(iMonth) = 1 Then
						iMonth = "0" & iMonth
					End If
					If Len(aDate(1))=1 Then
						aDate(1)="0" & Cstr(aDate(1))
					End IF
					Fn_CommonUtil_DateOperations=aDate(1) & "/" & iMonth & "/" & aDate(2)	
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "YYMMDD"	
					If iNumber="" Then
						iNumber=0
					ENd If
					aDateTime=Split(Now() + iNumber," ")
					aDate=Split(aDateTime(0),"/")
					iMonth=aDate(0)
					If Len(iMonth) = 1 Then
						iMonth = "0" & iMonth
					End If
					If Len(aDate(1))=1 Then
						aDate(1)="0" & Cstr(aDate(1))
					End IF
					Fn_CommonUtil_DateOperations = Right(aDate(2),2) & iMonth & aDate(1)
				End Select
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "GetPastDate"
			Select Case sFormat
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "DD-MMM-YYYY"	
					If iNumber="" Then
						iNumber=0
					ENd If
					aDateTime=Split(Now()," ")
					aDate=Split(aDateTime(0),"/")
					iMonth=MonthName(aDate(0)- iNumber,True)
					If Len(aDate(1))=1 Then
						aDate(1)="0" & Cstr(aDate(1))
					End IF
					Fn_CommonUtil_DateOperations=aDate(1) & "-" & iMonth & "-" & aDate(2)
			End Select			
	End Select
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_PSEFrozenColumnTableRowOperations
'
'Function Description	 :	Function used to perform operations on Frozen Column Table rows
'
'Function Parameters	 :  1.sAction	 		: Action to perform					
'							2.objJavaTabl  		: PSE BOM table object
'							3.sNodeName 		: Tree\Table node path
'
'Function Return Value	 : 	-1 or Row number of node
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	PSE BOM table should be available
'
'Function Usage		     :	bReturn=Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objBOMTable,"518611/A;1-Item_518611 (view)~001270/A;1-ffff")
'Function Usage		     :	bReturn=Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindexext",objBOMTable,"518611/A;1-Item_518611 (view)~001270/A;1-ffff")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  29-Jun-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_RAC_PSEFrozenColumnTableRowOperations(sAction,ByVal objJavaTable,sNodeName)
	'Variable Declaration
	Dim iInstance,iOccurance,iRows,iCounter,iCount,iColumnIndex,iRowIndex
	Dim sNodePath,sNodeName1,sNodePath1,sPath,sPath1,sNodePath2,sTopNode
	Dim aNodePath,aRowNode,aNodeName,aNodePathArray
	Dim bFlag
	Dim objComponent
	
	'Initially set function return value as -1
	Fn_RAC_PSEFrozenColumnTableRowOperations = -1
	
	Select Case LCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getnodeindex"
			iColumnIndex = 0
			bFlag = False
			iRows = cInt(objJavaTable.GetROProperty ("rows"))
			aNodeName = split(sNodeName,"~")
			iRowIndex = 0
			sPath = ""
			For iCounter=0 to UBound(aNodeName)
				aRowNode = Split(trim((aNodeName(iCounter))),"@")
				If sPath = "" Then
					sPath =  trim(aRowNode(0))
				Else
					sPath = sPath & "~" & trim(aRowNode(0))
				End If
			Next
			For iCounter=0 to UBound(aNodeName)
				If iRowIndex = iRows  Then
					Exit for
				End If
				aRowNode = split(trim((aNodeName(iCounter))),"@")
				iInstance = 0
				bFlag = False
				Do While iRowIndex < iRows
					If uBound(aRowNode) > 0 Then
						'instance number exist in name
						'initialize instance number
						'ith row matches with aRowNode(0) then
						sNodePath = objJavaTable.object.getCellRenderer(0,0).getPathForRow(iRowIndex).toString()
						sNodePath = Right(sNodePath, (Len(sNodePath)-Instr(1, sNodePath, ",", 1)))					
						sNodePath = trim(Left(sNodePath, Len(sNodePath)-1))
						sNodePath1 = Replace(LCase(sNodePath)," (view)","") 
						sTopNode=Replace(LCase(aRowNode(0))," (view)","")
							
						If trim(sNodePath) = trim(aRowNode(0)) or trim(sNodePath1) = trim(sTopNode) then							
							If instr(sNodePath2, "@BOM::") > 0 Then
								sNodePath2 = trim(replace(sNodePath2,"""",""))
								aNodePath = split(sNodePath2,",")
								sNodePath2 = ""
								For iCount = 0 to uBound(aNodePath)
									aNodePath(iCount) = Left(aNodePath(iCount), instr(aNodePath(iCount),"@")-1)
									If sNodePath2 = "" Then
										sNodePath2 = trim(aNodePath(iCount))
									else
										sNodePath2 = sNodePath2 & ", " & trim(aNodePath(iCount))
									End If
								Next
							End If
								
							sNodePath2 = trim(replace(sNodePath2,", ","~"))
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
							'Added Code to match node removing (view) word
							sNodePath1=replace(LCase(sNodePath2)," (view)","")
							sPath1=replace(LCase(sPath)," (view)","")
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -																									
							If instr(sPath, sNodePath2 ) > 0 Then
								iInstance = iInstance +1
								If iInstance = cInt(aRowNode(1)) Then 
									If UBound(aNodeName) = iCounter Then
										bFlag = True
									End If
									Exit do
								End If												
							ElseIf instr(sPath1, sNodePath1 ) > 0 Then
								iInstance = iInstance +1
								If iInstance = cInt(aRowNode(1)) Then 
									If UBound(aNodeName) = iCounter Then
										bFlag = True
									End If
									Exit do
								End If
							End If
						End if
					Else
						'if row matches with aRowNode(0) then
						If objJavaTable.object.getPathForRow(iRowIndex).getLastPathComponent().getClass().toString() <> "class com.teamcenter.rac.treetable.HiddenSiblingNode" Then
							sNodePath = objJavaTable.object.getValueAt(iRowIndex, iColumnIndex).toString()
						Else
							sNodePath = ""
						End If				
						sNodePath1 = Replace(LCase(sNodePath)," (view)","") 
						sTopNode=Replace(LCase(aRowNode(0))," (view)","") 

						If trim(sNodePath) = trim(aRowNode(0)) or trim(sNodePath1) =sTopNode  then
							sNodePath2 =objJavaTable.Object.getCellRenderer(0,0).getPathForRow(iRowIndex).toString()
							sNodePath2 = Right(sNodePath2, (Len(sNodePath2)-Instr(1, sNodePath2, ",", 1)))					
							'Added condition to select BOM line node in Markup mode
							If instr(sNodePath2, "TreeTableNode#")>0 Then
								aNodePathArray = Split(sNodePath2,",")
								sNodePath2 = aNodePathArray(0)
								sNodePath2 = Trim(sNodePath2) & ", " & sNodePath
							Else
								sNodePath2 = trim(Left(sNodePath2, Len(sNodePath2)-1))
							End If
							
							If instr(sNodePath2, "@BOM::") > 0 Then
								sNodePath2 = trim(replace(sNodePath2,"""",""))
								aNodePath = Split(sNodePath2,",")
								sNodePath2 = ""
								For iCount = 0 to uBound(aNodePath)
									aNodePath(iCount) = Left(aNodePath(iCount), instr(aNodePath(iCount),"@")-1)
									If sNodePath2 = "" Then
										sNodePath2 = trim(aNodePath(iCount))
									else
										sNodePath2 = sNodePath2 & ", " & trim(aNodePath(iCount))
									End If
								Next
							End If
						
							sNodePath2 = trim(replace(sNodePath2,", ","~"))
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
							'Added Code to match node removing (view) word
							sNodePath1=replace(LCase(sNodePath2)," (view)","")
							sPath1=replace(LCase(sPath)," (view)","")
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
							If instr(sPath, sNodePath2 ) > 0 Then
								If UBound(aNodeName) = iCounter Then
									bFlag = True
								End If
								Exit do
								'exit loop
							ElseIf instr(sPath1, sNodePath1 ) > 0 Then
								If UBound(aNodeName) = iCounter Then
									bFlag = True
								End If
								Exit do
							End if
						End if
					End If
					iRowIndex = iRowIndex + 1
				Loop
			Next
			If bFlag=False Then
				Fn_RAC_PSEFrozenColumnTableRowOperations = -1
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_PSEFrozenColumnTableRowOperations ] : Fail to get row index of node [ " & Cstr(sNodeName) & " ] as node is not exist in table")
			Else
				Fn_RAC_PSEFrozenColumnTableRowOperations = iRowIndex
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<PASS>:  [ Fn_RAC_PSEFrozenColumnTableRowOperations ] : Row index of node [ " & Cstr(sNodeName) & " ] is [ " & CStr(iRowIndex) & " ]")
			End If
			'Releasing table object	
			Set objJavaTable = Nothing		
	End Select
End Function