Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Function Name									|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001.	Fn_UI_Object_Operations							|	Sandeep.Navghane@sqs.com	|	02-Sep-2013	|	This function is Use to check Existance of given Object.
'002. 	Fn_UI_JavaButton_Operations						|	Ganesh.Bhosale@sqs.com		|	03-Feb-2016	|	This Function used to perform operations on JavaButton
'003. 	Fn_UI_JavaRadioButton_Operations				|	Ganesh.Bhosale@sqs.com		|	03-Feb-2016	|	This Function to perform operations on JavaRadioButton component.
'004. 	Fn_UI_JavaEdit_Operations						|	Ganesh.Bhosale@sqs.com		|	03-Feb-2016	|	This function to perform operations on JavaEdit component.
'005. 	Fn_UI_JavaList_Operations						|	Ganesh.Bhosale@sqs.com		|	03-Feb-2016	|	This function to perform operations on JavaList component.
'006. 	Fn_UI_JavaCheckBox_Operations					|	Ganesh.Bhosale@sqs.com		|	04-Feb-2016	|	This function to perform operations on JavaCheckbox component.
'007.   Fn_UI_JavaStaticText_Operations					|	Ganesh.Bhosale@sqs.com		|	04-Feb-2016	|	This function to perform operations on JavaStaticText component.
'008.   Fn_UI_Twistie_Operations						|	Ganesh.Bhosale@sqs.com		|	08-Feb-2016	|	This function to perform operations on Twistie component.
'009.   Fn_UI_JavaTab_Operations						|	Ganesh.Bhosale@sqs.com		|	08-Feb-2016	|	This function to perform operations on JavaTab component.
'010.   Fn_UI_JavaTable_Operations						|	Ganesh.Bhosale@sqs.com		|	08-Feb-2016	|	This function to perform operations on Java Table component.
'011.	Fn_UI_JavaTree_Operations						|	Ganesh.Bhosale@sqs.com		|	08-Feb-2016	|	This function to perform operations on Java Tree component.
'012.	Fn_UI_Object_GetChildObjects					|	vrushali.sahare@sqs.com		|	24-Feb-2016	|	This function used to get child objects of specified component descriptively.
'013.	Fn_UI_JavaMenu_Operations						|	vrushali.sahare@sqs.com		|	25-Feb-2016	|	This function to perform operations on JavaMenu component
'014.	Fn_UI_JavaObject_Operations						|	vrushali.sahare@sqs.com		|	25-Feb-2016	|	This function to perform operations on Java Object component
'015.	Fn_UI_JavaToolbar_Operations					|	Sandeep.Navghane@sqs.com	|	08-Jun-2016	|	This function to perform operations on Java toolbar component
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_Object_Operations
'
'Function Description	 :	This function is Use to check Existance of given Object.
'
'Function Parameters	 :  1. sFunctionName	: Valid Function Name
'						    2. sAction			: Action to be performend
'						    3. objReferencePath : Valid Java Object Hierarchy Path	
'				    		4. iTimeOut			: Time out time in seconds.
'							5. sPropertyName	: Valid Property Name
'							6. sPropertyValue	: Property value
'
'Function Return Value	 : 	True \ False \ Object
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  bReturn = Fn_UI_Object_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Create", objDateControl,"", "", "") - Depricated
'Function Usage		     :  bReturn = Fn_UI_Object_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Enabled", objDateControl, GBL_MAX_TIMEOUT, "", "")
'Function Usage		     :  bReturn = Fn_UI_Object_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Exist", objDateControl,"", "", "")
'Function Usage		     :  bReturn = Fn_UI_Object_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","settoexistcheck", objDateControl,"", "title", "Date")
'Function Usage		     :  bReturn = Fn_UI_Object_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","getroproperty", objDateControl,"", "abs_x", "")
'Function Usage		     :  bReturn = Fn_UI_Object_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","settoproperty", objDateControl,"", "title", "Date")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale			    |  02-Sep-2013	    |	 1.0		|	Sandeep Navghane 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End

Public Function Fn_UI_Object_Operations(sFunctionName, sAction, objReferencePath, iTimeOut,sPropertyName,sPropertyValue)
	Err.Clear
	'Declaring variables
	Dim bFlag
	
	GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_Object_Operations : [ " &  objReferencePath.toString & " ] : Action = " & sAction & " : "
	Fn_UI_Object_Operations = False
	
	If iTimeOut = "" Then
		iTimeOut = cInt(GBL_DEFAULT_TIMEOUT)
	Else
		iTimeOut = cInt(iTimeOut)
	End If
	
	Select Case lCase(sAction)
		Case "create"
			If Fn_UI_Object_Operations(sFunctionName, "Enabled", objReferencePath, iTimeOut, "", "") Then
				Set Fn_UI_Object_Operations = objReferencePath
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "New Object created for " & objReferencePath.toString & " in Function ")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "Object " & objReferencePath.toString & " is not enable of Function ")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "exist"
			Fn_UI_Object_Operations = objReferencePath.Exist(iTimeOut)
			If Fn_UI_Object_Operations Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & " [ " & objReferencePath.tostring & " ] object is exist.")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & " [ " & objReferencePath.tostring & " ] object does not exist.")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "enabled"
			If Fn_UI_Object_Operations(sFunctionName, "Exist", objReferencePath, iTimeOut, "", "") Then
				If objReferencePath.GetROProperty("enabled") = "1"  OR objReferencePath.GetROProperty("enabled") = True Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " [ " & objReferencePath.tostring & " ] object is exists and enabled.")
					Fn_UI_Object_Operations = True
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & " [ " & objReferencePath.tostring & " ]object is not enabled.")
				End If
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & " [ " & objReferencePath.tostring & " ] does not exist.")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "settoexistcheck"
				objReferencePath.SetTOProperty sPropertyName,sPropertyValue
				bFlag= Fn_UI_Object_Operations("Fn_UI_Object_Operations", "Exist", objReferencePath , "", "", "")
				If  bFlag = True Then
					Fn_UI_Object_Operations = True
					'log on success
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully Set TO property [ " & sPropertyName & " ] and  verified [ " & objReferencePath.tostring & " ] object is exist.")
				End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getroproperty"

Fn_UI_Object_Operations=objReferencePath.GetROProperty(sPropertyName)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully fetched  [ " & sPropertyName & " ] property value.")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "settoproperty"
				objReferencePath.SetTOProperty sPropertyName,sPropertyValue
				Fn_UI_Object_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully Set TO property [ " & sPropertyName & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : Invalid Case.")
	End Select
	' log for error
	If Err.Number <> 0 Then	
		Fn_UI_Object_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & " ] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaButton_Operations
'
'Function Description	 :	This function is used to perform operations on JavaButton object
'
'Function Parameters	 :  1.sFunctionName	: Function name from which this function is called.
'							2.sAction		: Action to be performed on JavaButton
'							3.objJavaDialog	: Parent UI Component or JavaRadioButton object
'							4.sJavaButton	: Valid JavaButton name
'
'Function Return Value	 : 	True \ False 
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  bReturn = Fn_UI_JavaButton_Operations("Fn_TeamcenterLogin", "Click", JavaWindow("Teamcenter Login"),"Login")
'Function Usage		     :  bReturn = Fn_UI_JavaButton_Operations("Fn_TeamcenterLogin", "Click", JavaWindow("Teamcenter Login").JavaButton("Login"),"")
'Function Usage		     :  bReturn = Fn_UI_JavaButton_Operations("Fn_TeamcenterLogin", "DeviceReplay.Click", JavaWindow("Teamcenter Login").JavaButton("Login"),"")
'Function Usage		     :  bReturn = Fn_UI_JavaButton_Operations("Fn_TeamcenterLogin", "object.Click", JavaWindow("Teamcenter Login").JavaButton("Login"),"")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale			    |  03-Feb-2016	    |	 1.0		|	Sandeep Navghane 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaButton_Operations(sFunctionName, sAction, objJavaDialog, sJavaButton)
	Dim objJavaButton, objDeviceReplay
	Fn_UI_JavaButton_Operations = False
	'Object Creation
	If sJavaButton <> "" Then
		Set objJavaButton = objJavaDialog.JavaButton(sJavaButton)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaButton_Operations  : [ " &  objJavaDialog.toString & " ] : [ " &  objJavaButton.toString & " ] : Action = " & sAction & " : "
	Else
		Set objJavaButton = objJavaDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaButton_Operations  : [ " &  objJavaButton.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaButton object exists
	If Fn_UI_Object_Operations("Fn_UI_JavaButton_Operations", "Enabled", objJavaButton,"", "", "") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objJavaButton.toString & " ] object Does not exist.")
		Set objJavaButton = Nothing 
		Exit Function
	End If
	
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "click"
			'clcik on JavaButton
			objJavaButton.Click	
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully clicked on JavaButton.")
			Fn_UI_JavaButton_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "object.click"
			'to click on java button
			If objJavaButton.GetROProperty("enabled") = 1 Then
			   objJavaButton.Object.click
			   Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully clicked on JavaButton.")
			   Fn_UI_JavaButton_Operations = True 
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "devicereplay.click"
			'to click on java button using devicereplay method
			If sJavaButton <> "" Then
				objJavaDialog.Activate
			End If
			'create DeviceReplay object
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			'move mouce to JavaButton
			objDeviceReplay.MouseMove (objJavaButton.GetROProperty("abs_x") + 5), (objJavaButton.GetROProperty("abs_y") + 5)
			'click on Java button
			objDeviceReplay.MouseClick  (objJavaButton.GetROProperty("abs_x") + 5), (objJavaButton.GetROProperty("abs_y") + 5), 0
		    Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully clicked on JavaButton.")
			Fn_UI_JavaButton_Operations = True
			'Clear memory of JavaButton object.
			Set objDeviceReplay = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "[ " & GBL_FUNCTIONLOG & " ] : FAIL : Invalid Case.")
			Set objJavaButton = Nothing 
			Exit Function
	End Select
	' log for error
	If Err.Number <> 0 Then	
		Fn_UI_JavaButton_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Clear memory of JavaButton object.
	Set objJavaButton = Nothing 
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaRadioButton_Operations
'
'Function Description	 :	This function to perform operations on JavaRadioButton component.
'
'Function Parameters	 :  1. sFunctionName	: Function name from which this function is called.
'							2. sAction			: Action to be performed
'							3. objJavaDialog 	: Parent UI Component or JavaRadioButton object
'							4. sRadioButtonName	: JavaRadioButton Control name
'							5. sValue			: value to be selected
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  bReturn = Fn_UI_JavaRadioButton_Operations("Fn_SISW_RAC_PMM_UserContextSettings", "Set", objDialog, "Version", "ON")
'Function Usage		     :  bReturn = Fn_UI_JavaRadioButton_Operations("Fn_SISW_RAC_PMM_UserContextSettings", "Set", objDialog.JavaRadioButton("Precise"), "" , "OFF")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale			    |  16-Jan-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaRadioButton_Operations(sFunctionName, sAction, objJavaDialog, sRadioButtonName, sValue)
	Dim objRadioButton
	
	'Initialize Function value to False
	Fn_UI_JavaRadioButton_Operations = False
	'Object Creation
	If sRadioButtonName <> "" Then
		Set objRadioButton = objJavaDialog.JavaRadioButton(sRadioButtonName)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaRadioButton_Operations  : [ " &  objJavaDialog.toString & " ] : [ " &  objRadioButton.toString & " ] : Action = " & sAction & " : "
	Else
		Set objRadioButton = objJavaDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaRadioButton_Operations  : [ " &  objRadioButton.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaRadioButton object exists
	If Fn_UI_Object_Operations("Fn_UI_JavaRadioButton_Operations", "Exist", objRadioButton,"", "", "") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objRadioButton.toString & " ] object Does not exist.")
		Fn_UI_JavaRadioButton_Operations = False
		Set objRadioButton = Nothing 
		Exit Function
	End If
	
	Select case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "set"
			objRadioButton.Set sValue
			Fn_UI_JavaRadioButton_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully set JavaRadioButton [ " & sValue & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "[ " &  GBL_FUNCTIONLOG & " ] : FAIL : Invalid Case.")
			Set objRadioButton = Nothing 
			Exit Function
	End Select
	If Err.Number <> 0 Then	
		Fn_UI_JavaRadioButton_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Clear memory of JavaRadioButton object.
	Set objRadioButton = Nothing 
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaEdit_Operations
'
'Function Description	 :	This function to perform operations on JavaEdit component.
'
'Function Parameters	 :  1. sFunctionName	: Function name from which this function is called.
'							2. sAction			: Action to be performed
'							3. objJavaDialog 	: Parent UI Component or JavaRadioButton object
'							4. sJavaEdit		: JavaEdit Control name
'							5. sText			: text to be set in edit box
'
'Function Return Value	 : 	True \ False \ text in edit box
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  bReturn = Fn_UI_JavaEdit_Operations("Fn_SISW_PMM_UserContextSettings", "Set",  objDialog, "EWO", "EWO_Name" )
'Function Usage		     :  bReturn = Fn_UI_JavaEdit_Operations("Fn_SISW_PMM_UserContextSettings", "Type",  objDialog.JavaEdit("EWO"),"", "EWO_Name" )
'Function Usage		     :  bReturn = Fn_UI_JavaEdit_Operations("Fn_SISW_PMM_UserContextSettings", "activate",  objDialog.JavaEdit("EWO"),"", "" )
'Function Usage		     :  bReturn = Fn_UI_JavaEdit_Operations("Fn_SISW_PMM_UserContextSettings", "setsecure",  objDialog.JavaEdit("EWO"),"", "EWO_Name" )
'Function Usage		     :  bReturn = Fn_UI_JavaEdit_Operations("Fn_SISW_PMM_UserContextSettings", "settext",  objDialog.JavaEdit("EWO"),"", "EWO_Name" )
'Function Usage		     :  bReturn = Fn_UI_JavaEdit_Operations("Fn_SISW_PMM_UserContextSettings", "gettext",  objDialog.JavaEdit("EWO"),"", "" )
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale		   		|  03-Feb-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaEdit_Operations(sFunctionName, sAction,  ByVal objJavaDialog, sJavaEdit, sText)
	Dim objJavaEdit
	Dim objDeviceReplay
	
	Fn_UI_JavaEdit_Operations = False
	'Set an Edit Object on variable
	If sJavaEdit <> "" Then
		Set objJavaEdit= objJavaDialog.JavaEdit(sJavaEdit)
		GBL_FUNCTIONLOG = sFunctionName & "> Fn_UI_JavaEdit_Operations : [ " &  objJavaDialog.toString & " ] : [ " & objJavaEdit.toString & " ] : Action = " & sAction & " : "
	Else
		Set objJavaEdit= objJavaDialog
		GBL_FUNCTIONLOG = sFunctionName & "> Fn_UI_JavaEdit_Operations : [ " & objJavaEdit.toString & " ] : Action = " & sAction & " : "
	End If

	If Fn_UI_Object_Operations("Fn_UI_JavaEdit_Operations", "Exist", objJavaEdit,"", "", "") = False Then
		Fn_UI_JavaEdit_Operations= False
		Set objJavaEdit = Nothing 
		Exit Function
	End If
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "set"
			'Setting the editbox
			If Fn_UI_Object_Operations("Fn_UI_JavaEdit_Operations", "Enabled", objJavaEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL :[ " &  objJavaEdit.toString & " ] object Does not exist.")
				Set objJavaEdit = Nothing 
				Exit Function
			End If
			' set value in edit box
			objJavaEdit.Set sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Text " & sText  &" is Set/ Entered in JavaEditBox.")
			Fn_UI_JavaEdit_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "type"
			If Fn_UI_Object_Operations("Fn_UI_JavaEdit_Operations", "Enabled", objJavaEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objJavaEdit.toString & " ] does is not enabled.")
				Set objJavaEdit = Nothing 
				Exit Function
			End If
			'type the value in editbox
			objJavaEdit.Type sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Text " & sText  &" is Set/ Entered in JavaEditBox.")
			Fn_UI_JavaEdit_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "activate"
			If Fn_UI_Object_Operations("Fn_UI_JavaEdit_Operations", "Enabled", objJavaEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objJavaEdit.toString & " ] object does not exist.")
				Set objJavaEdit = Nothing 
				Exit Function
			End If
			'Activate the editbox
			objJavaEdit.Activate
			Fn_UI_JavaEdit_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "gettext"
			' get value from edit box
			Fn_UI_JavaEdit_Operations = objJavaEdit.getROProperty("value")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "setsecure"
			'Setting the edit box
			If Fn_UI_Object_Operations("Fn_UI_JavaEdit_Operations", "Enabled", objJavaEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objJavaEdit.toString & " ] object Does not exist.")
				Set objJavaEdit = Nothing 
				Exit Function
			End If
			' set secure value(encoded text) in edit box
			objJavaEdit.SetSecure sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Text " & sText  &"is Set/ Entered in JavaEditBox.")
			Fn_UI_JavaEdit_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "settext"
			'Setting the editbox
			If Fn_UI_Object_Operations("Fn_UI_JavaEdit_Operations", "Enabled", objJavaEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objJavaEdit.toString & " ] object Does not exist.")
				Set objJavaEdit = Nothing 
				Exit Function
			End If
			'set value in edit box using settest method. which is used when set and type methods are not working on edit box.
			objJavaEdit.Object.setText sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Text " & sText  &"is Set/ Entered in JavaEditBox.")
			Fn_UI_JavaEdit_Operations= True		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "sendstring"
			'Setting the editbox
			If Fn_UI_Object_Operations("Fn_UI_JavaEdit_Operations", "Enabled", objJavaEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL :[ " &  objJavaEdit.toString & " ] object Does not exist.")
				Set objJavaEdit = Nothing 
				Exit Function
			End If
			' set value in edit box
			'create DeviceReplay object
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			'move mouce to JavaButton
			objDeviceReplay.MouseMove (objJavaEdit.GetROProperty("abs_x") + 5), (objJavaEdit.GetROProperty("abs_y") + 5)
			'click on Java button
			objDeviceReplay.MouseClick  (objJavaEdit.GetROProperty("abs_x") + 5), (objJavaEdit.GetROProperty("abs_y") + 5), 0
			objDeviceReplay.SendString sText
			Set objDeviceReplay = Nothing
		
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Text " & sText  &" is Set/ Entered in JavaEditBox.")
			Fn_UI_JavaEdit_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : Invalid Case.")
			Set objJavaEdit = Nothing 
			Exit Function
	End Select
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Set objJavaEdit = Nothing 
	If Err.Number <> 0 Then	
		Fn_UI_JavaEdit_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaList_Operations
'
'Function Description	 :	This function to perform operations on Javalist component.
'
'Function Parameters	 :  1.sFunctionName 	: Valid Function name 
'							2. sAction			: Action to be performed
'							3. objJavaDialog	: Parent dialog \ Javalist Object
'							4. sJavaList   		: JavaList Name
'							5. sValues   		: Value to be selected \ verified
'							6. sColumns   		: Column Name 
'							7. sInstanceHandler : Instance Handler
'
'Function Return Value	 : 	True \ False \ content of list \ selected value
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage			 :	Fn_UI_JavaList_Operations("Fn_TeamcenterLogin", "Select", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Function Usage			 :	Fn_UI_JavaList_Operations("Fn_TeamcenterLogin", "type", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Function Usage			 :	Fn_UI_JavaList_Operations("Fn_TeamcenterLogin", "Click", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Function Usage			 :	Fn_UI_JavaList_Operations("Fn_TeamcenterLogin", "Activate", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Function Usage			 :	Fn_UI_JavaList_Operations("Fn_TeamcenterLogin", "Exist", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Function Usage			 :	Fn_UI_JavaList_Operations("Fn_TeamcenterLogin", "ExtendSelect", JavaWindow("Teamcenter Login"),"Login","AutoTest", "", "")
'Function Usage			 :	Fn_UI_JavaList_Operations("Fn_TeamcenterLogin", "GetText", JavaWindow("Teamcenter Login"),"Login","", "", "")
'Function Usage			 :	Fn_UI_JavaList_Operations("Fn_TeamcenterLogin", "GetContents", JavaWindow("Teamcenter Login"),"Login","", "", "")
'Function Usage			 :	Fn_UI_JavaList_Operations("Fn_TeamcenterLogin", "VerifyContents", JavaWindow("Teamcenter Login"),"Login","", "", "") 
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale 			    |  03-Feb-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaList_Operations(sFunctionName, sAction, objJavaDialog, sJavaList, sValues, sColumns, sInstanceHandler)
	'Declare variables
	Dim objJavaList
	Dim aSelectList, aContents, aValues
	Dim iCounter, iEelementCount, iInstanceCnt, iCount
	Dim sContents
	Dim bFlag
	
	'Initially set function return value as False
	Fn_UI_JavaList_Operations = False
	'Set the instance handler
	If sInstanceHandler = "" Then
		sInstanceHandler = "@"
	End If
	'Object Creation
	If sJavaList <> "" Then
		Set objJavaList = objJavaDialog.JavaList(sJavaList)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaList_Operations  : [ " &  objJavaDialog.toString & " ] : [ " &  objJavaList.toString & " ] : Action = " & sAction & " : "
	Else
		Set objJavaList = objJavaDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaList_Operations  : [ " &  objJavaList.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaList object exists
	If Fn_UI_Object_Operations("Fn_UI_JavaList_Operations", "Exist", objJavaList,"", "", "") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objJavaList.toString & " ] List does not exist.")
		Set objJavaList = Nothing 
		Exit Function
	End If
	
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getcontents"
			' get Total Items in list
			iEelementCount = objJavaList.GetROProperty("items count")
			' get all items from list
			For iCounter = 0 To iEelementCount - 1
				If iCounter = 0 Then
					Fn_UI_JavaList_Operations = Trim(cstr(objJavaList.GetItem(iCounter)))
				Else
					Fn_UI_JavaList_Operations = Fn_UI_JavaList_Operations & "~" & Trim(cstr(objJavaList.GetItem(iCounter)))
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "gettext"
			' get text from list
			Fn_UI_JavaList_Operations = objJavaList.GetROProperty("value")
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "Successfully fetched currently selected value [" & Fn_UI_JavaList_Operations & "] from [ " & objJavaList.tostring & "].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "click"
			' click list Item from list
			objJavaList.Click sValues, sColumns
			Fn_UI_JavaList_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "Successfully clicked on [" & sValues & "] list item from [ " & objJavaList.tostring & "].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "activate"
			' Activate list Item from list
			objJavaList.Activate sValues
			Fn_UI_JavaList_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "Successfully Activated [" & sValues & "] list item from [ " & objJavaList.tostring & "].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "exist"
			' get total items count from list
			iEelementCount = objJavaList.GetROProperty("items count")
			iInstanceCnt = 1
			aSelectList = split(sValues, sInstanceHandler)
			If uBound(aSelectList) > 0 Then
				iInstanceCnt = aSelectList(1)
			End If
			' get total items from list and compare with given value
			aSelectList(0) = trim(aSelectList(0))
			For iCounter = 0 To iEelementCount - 1
				If objJavaList.GetItem(iCounter) <> "" Then
					If Trim(cstr(objJavaList.GetItem(iCounter))) = Trim(aSelectList(0)) Then
						iF iInstanceCnt = 1 Then
							bFlag = True
							Fn_UI_JavaList_Operations = True
							Exit For
						End If
						iInstanceCnt = iInstanceCnt - 1 
					End If
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "extendselect"
			aSelectList=Split(sValues,"~")
			' multiselect the element from list  
			For iCounter = 0 To Ubound(aSelectList)
				objJavaList.ExtendSelect aSelectList(iCounter)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "Successfully Selected [" & aSelectList(iCounter) & "]")
			Next
			Fn_UI_JavaList_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "type"
			' type value to select item from list
			objJavaList.Type sValues
			Fn_UI_JavaList_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "select"
			' get total items from list
			iEelementCount = objJavaList.GetROProperty("items count")
			iInstanceCnt = 1
			aSelectList = split(sValues, sInstanceHandler)
			If uBound(aSelectList) > 0 Then
				iInstanceCnt = aSelectList(1)
			End If
			' select item from list
			aSelectList(0) = trim(aSelectList(0))
			For iCounter = 0 To iEelementCount - 1
				If objJavaList.GetItem(iCounter) <> "" Then
					If Trim(cstr(objJavaList.GetItem(iCounter))) = Trim(aSelectList(0)) Then
						iF iInstanceCnt = 1 Then
							objJavaList.Select iCounter
							Fn_UI_JavaList_Operations = True
							Exit For
						End If
						iInstanceCnt = iInstanceCnt - 1 
					End If
				End If
			Next
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        Case "verifycontents"
			' Get Contents from list
			sContents = Fn_UI_JavaList_Operations("Fn_UI_JavaList_Operations", "GetContents", objJavaList,"","", "", "")
			If sContents = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "Failed to get Content of list")
				' release object
				Set objJavaList = Nothing 
				Exit Function
			End If
			aContents = Split(sContents,"~",-1,1)
			aValues = split(sValues,"~",-1,1)
			For iCount = 0 to Ubound(aValues)
				bFlag = False
				For iCounter = 0 to Ubound(aContents)
					If  aValues(iCount) = aContents(iCounter) Then
						bFlag = True
						Exit For
					End If
				Next
				If bFlag = False Then
					Fn_UI_JavaList_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "Failed to verify that the value [ "& aValues(iCount)&" ] is in list")
					Exit For
				End If
			Next
			If bFlag = True Then
				Fn_UI_JavaList_Operations = True
			End If
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : Invalid Case.")
			Set objJavaList = Nothing 
			Exit Function
	End Select
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If Err.Number <> 0 Then
		Fn_UI_JavaList_Operations = False 
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_UI_JavaButton_Operations ] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	' release object
	Set objJavaList = Nothing 
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaCheckBox_Operations
'
'Function Description	 :	This function to perform operations on JavaCheckBox component.
'
'Function Parameters	 :  1. sFunctionName	: Caller function's name
'							2. sAction			: Action to be performed
'							3. objJavaDialog 	: Parent UI Component or JavaCheckBox object
'							4. sCheckBoxName 	: JavaCheckBox Control name
'							5. sValue			: value
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage			 :	call Fn_UI_JavaCheckBox_Operations("Fn_TeamcenterLogin", "Set", objDialog, "Version", "ON")
'Function Usage			 :	Call Fn_UI_JavaCheckBox_Operations("Fn_TeamcenterLogin", "Set", objDialog.JavaCheckBox("Precise"), "" , "OFF")
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale 			    |  03-Feb-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaCheckBox_Operations(sFunctionName, sAction, objJavaDialog, sCheckBoxName, sValue)
	Dim objCheckBox
	
	'Initially set function return value as False
	Fn_UI_JavaCheckBox_Operations = False
	
	'Object Creation
	If sCheckBoxName <> "" Then
		Set objCheckBox = objJavaDialog.JavaCheckBox(sCheckBoxName)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaCheckBox_Operations  : [ " &  objJavaDialog.toString & " ] : [ " &  objCheckBox.toString & " ] : Action = " & sAction & " : "
	Else
		Set objCheckBox = objJavaDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaCheckBox_Operations  : [ " &  objCheckBox.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaCheckBox object exists
	If Fn_UI_Object_Operations("Fn_UI_JavaCheckBox_Operations", "Exist", objCheckBox,"", "", "") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objCheckBox.toString & " ] object Does not exist.")
		Fn_UI_JavaCheckBox_Operations = False
			Set objCheckBox = Nothing 
		Exit Function
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
	Select case lCase(sAction)
		'check or uncheck JavaCheckbox
		Case "set"
			objCheckBox.WaitProperty "enabled",1,90000
			objCheckBox.Set uCase(sValue)
			Fn_UI_JavaCheckBox_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully set JavaCheckBox [ " & sValue & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : Invalid Case.")
	End Select
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If Err.Number <> 0 Then	
		Fn_UI_JavaCheckBox_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Clear memory of JavaCheckBox object.
	Set objCheckBox = Nothing 
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaStaticText_Operations
'
'Function Description	 :	This function to perform operations on Java Static Text component.
'
'Function Parameters	 :  1. sFunctionName	: Caller function's name
'							2. sAction			: Action to be performed
'							3. objContainer 	: Parent UI Component or JavaRadioButton object
'							4. sJavaStaticText 	: Java Static Text Name
'							5. iX 				: X coordinate of Java Static Text
'							6. iY 				: Y coordinate of Java Static Text
'							4. sMouseButton 	: Mouse button Name [ LEFT | RIGHT ]
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage			 :	call Fn_UI_JavaStaticText_Operations("","Click",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"),"BottomLink","1","1","")
'Function Usage			 :	call Fn_UI_JavaStaticText_Operations("","Click",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").JavaStaticText("BottomLink"),"","1","1","")
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale 				|  03-Feb-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaStaticText_Operations(sFunctionName,sAction,objContainer,sJavaStaticText,iX,iY,sMouseButton)
	
	' declaring variables
	Dim objJavaStaticText
	
	Fn_UI_JavaStaticText_Operations = False
	' object creation
	If sJavaStaticText<>"" Then
		Set objJavaStaticText=objContainer.JavaStaticText(sJavaStaticText)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaStaticText_Operations : [ " & objContainer.toString() & " ] : Action = " & sAction & " : Java Static text = " & sJavaStaticText & " : "
	Else
		Set objJavaStaticText=objContainer
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaStaticText_Operations Java Static text : [ " & objContainer.toString() & " ] : Action = " & sAction & " : "	
	End If
	
	'Checking existance of parent window
	If Fn_UI_Object_Operations("Fn_UI_JavaStaticText_Operations", "Exist", objJavaStaticText,"", "", "") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Java Static text [ " & objJavaStaticText.tostring & " ] object does not exist")
		Set objJavaStaticText = Nothing
		Exit Function
	End If
		
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "click"
			'check object enabled or not
			If Fn_UI_Object_Operations("Fn_UI_JavaStaticText_Operations", "Enabled", objJavaStaticText,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Java Static text [ " & objJavaStaticText.tostring & " ] object not Enabled")
				Set objJavaStaticText = Nothing
				Exit Function
			End If
			' click on java static text
			If sMouseButton<>"" Then
				objJavaStaticText.Click iX,iY,sMouseButton
			Else
				objJavaStaticText.Click iX,iY			
			End If
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " PASS: Successfully Click on Java Static text [ " & objJavaStaticText.tostring & " ]")
			Fn_UI_JavaStaticText_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " FAIL: invalid case.")		
	End Select
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If Err.Number <> 0 Then	
		Fn_UI_JavaStaticText_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	Set objJavaStaticText=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_Twistie_Operations
'
'Function Description	 :	This function to perform operations on Twistie control.
'
'Function Parameters	 :  1. sFunctionName	: Caller function's name
'							2. sAction			: Action to be performed on Twistie component
'							3. objContainer 	: Parent UI Component
'							4. sTwistie 		: Twistie Control name in OR
'							5. sTwistieText		: Twistie Control text
'							6. sTwistieStatic 	: Static Text Against Twistie
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage			 :	call Fn_UI_Twistie_Operations("", "Expand", JavaWindow("ServicePlanner"), "Twistie", "Fault","StaticText")
'Function Usage			 :	call Fn_UI_Twistie_Operations("", "Collapse", JavaWindow("ServicePlanner"), "Twistie", "References","StaticText") 
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale 				|  03-Feb-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_Twistie_Operations(sFunctionName, sAction, objContainer, sTwistie, sTwistieText,sTwistieStatic)
	'declaring variables
	Dim objTwistie, objTwistieText
	
	'create objects
	Set objTwistie = objContainer.JavaObject(sTwistie)
	Set objTwistieText = objContainer.JavaStaticText(sTwistieStatic)
	
	Fn_UI_Twistie_Operations = False
	
	' set label property of static text
	objTwistieText.SetTOProperty "label", sTwistieText
	GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_Twistie_Operations : [ " & sTwistieText & " Twistie Control ] : Action = " & sAction & " : "
	If objTwistie.Exist(GBL_MIN_TIMEOUT)= False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " & sTwistieText & "] object does not exist.")
		Set objTwistie = Nothing 
		Set objTwistieText = Nothing
		Exit Function
	End If

	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "expand"
			'  verify Twistie object is expanded or not
			If cBool(objTwistie.Object.isExpanded()) = False Then
				' click on Twistie object to expand
				objTwistie.Click 1, 1,"LEFT"
				wait GBL_MIN_MICRO_TIMEOUT
				'  verify Twistie object is expanded or not
				If cBool(objTwistie.Object.isExpanded()) = True Then
					Fn_UI_Twistie_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Successfully expanded [ " & sTwistieText & " ]." )
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : Failed to expand [ " & sTwistieText & " ]." )
				End If
			Else
				Fn_UI_Twistie_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Successfully expanded[ " & sTwistieText & " ]." )
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "collapse"	
			'  verify Twistie object is expanded or not
			If cBool(objTwistie.Object.isExpanded()) = True Then
				' click on Twistie object to collpse
				objTwistie.Click 1, 1,"LEFT"
				wait GBL_MIN_MICRO_TIMEOUT
				'  verify Twistie object is expanded or not
				If cBool(objTwistie.Object.isExpanded()) = False Then
					Fn_UI_Twistie_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Successfully collapsed [ " & sTwistieText & " ]." )
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : Failed to collapse [ " & sTwistieText & " ]." )
				End If
			Else
				Fn_UI_Twistie_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Successfully collapsed [ " & sTwistieText & " ]." )
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " FAIL: invalid case.")		
	End Select
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If Err.Number <> 0 Then	
		Fn_UI_Twistie_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	' relaease objects
	Set objTwistie = Nothing 
	Set objTwistieText = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaTab_Operations
'
'Function Description	 :	This function to perform operations on JavaTab component.
'
'Function Parameters	 :  1. sFunctionName	: Caller function's name
'							2. sAction			: Action to be performed
'							3. objContainer 	: Parent UI Component or JavaCheckBox object
'							4. sTabObjectName 	: JavaTab Control name in OR
'							5. sItem			: Tab Text to be selectd or verified
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage			 :	bReturn = Fn_UI_JavaTab_Operations("", "Select", JavaWindow("ProductMasterManager"), "AuthorizationTab",  "EW11397")
'Function Usage			 :	bReturn = Fn_UI_JavaTab_Operations("", "Click", JavaWindow("ProductMasterManager"), "AuthorizationTab",  "EW11397")
'Function Usage			 :	bReturn = Fn_UI_JavaTab_Operations("", "Exist", JavaWindow("ProductMasterManager").JavaTab("AuthorizationTab"), "", "EW11397")
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale 				|  03-Feb-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaTab_Operations(sFunctionName, sAction,objContainer,sTabObjectName,sItem)
	Dim objTab , objItems
	Dim sBounds
	Dim iItemCount, iCounter
	Dim aBounds

	'Object Creation
	Fn_UI_JavaTab_Operations = False
	If sTabObjectName <> "" Then
		Set objTab = objContainer.JavaTab(sTabObjectName)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaTab_Operations  : [ " &  objContainer.toString & " ] : [ " &  sTabObjectName & " ] : Action = " & sAction & " : "
	Else
		Set objTab = objContainer
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaTab_Operations  : [ " &  objTab.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify JavaTab object exists
	If Fn_UI_Object_Operations("Fn_UI_JavaTab_Operations","Exist", objTab, GBL_MIN_TIMEOUT, "", "") = False Then	
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objTab.toString & " ] object does not exist.")
		Set objTab = Nothing 
		Exit Function
	End If
	
	Select case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "select"
			'Verify JavaTab Item exists
			If Fn_UI_JavaTab_Operations(sFunctionName, "Exist", objContainer, sTabObjectName, sItem) = True Then
				On error resume next
				'select JavaTab
				objTab.Select sItem
				If Err.Number <> 0 Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : Failed to find JavaTab [ " & sItem & " ].")
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : Error Description - " & Err.Description )
                    On Error GoTo 0
				Else
					Fn_UI_JavaTab_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully select JavaTab [ " & sItem & " ].")
				End If				
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : Failed to find JavaTab [ " & sItem & " ].")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "exist"
				Select Case objTab.Object.getClass().toString()
					Case "class javax.swing.JTabbedPane"
						'get total no. of tab items
						iItemCount = cInt(objTab.Object.getTabCount())
						For iCounter = 0 to iItemCount - 1
							If sItem = objTab.Object.getTitleAt(iCounter) Then
								Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully verified JavaTab [ " & sItem & " ].")
								Fn_UI_JavaTab_Operations = True
								Exit for
							End If
						Next
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					Case Else
						'create object Tabs
						Set objItems = objTab.Object.getItems()
						'This condition put to handle BMIDE Tab
						If instr(1,lCase(objTab.Object.toString()),"wrong thread") then
							Fn_UI_JavaTab_Operations = True
							Exit function
						End if
						'get total no. of tab items
						iItemCount = cInt(objTab.Object.getItemCount())
						For iCounter = 0 to iItemCount - 1
							If sItem = objItems.mic_arr_get(iCounter).getText() Then
								Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully verified JavaTab [ " & sItem & " ].")
								Fn_UI_JavaTab_Operations = True
								Exit for
							End If
						Next
						Set objItems = Nothing
				End Select
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
         Case "click"
			'Set objItems = objTab.Object.getItems()
			'This condition put to handle BMIDE Tab
			If instr(1,lCase(objTab.Object.toString()),"wrong thread") then
				Fn_UI_JavaTab_Operations = True
				Set objItems = Nothing
				Exit function
			End if
			''get total no. of tab items
			iItemCount = cInt(objTab.Object.getItemCount())
			For iCounter = 0 to iItemCount - 1
				If sItem = objItems.mic_arr_get(iCounter).getText() Then
					' get boundary of TabItem
					sBounds=objTab.Object.getItem(iCounter).getBounds().toString()
					sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
					aBounds = split(sBounds,",")
					' click on Tab
					objTab.Click aBounds(0)+ 5, aBounds(1) +5
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully Click on JavaTab [ " & sItem & " ].")
					Fn_UI_JavaTab_Operations = True
					Exit for
				End If
			Next
			Set objItems = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " FAIL: invalid case.")		
	End Select
	If Err.Number <> 0 Then	
		Fn_UI_JavaTab_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Clear memory of JavaTab object.
	Set objTab = Nothing 
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaTable_Operations
'
'Function Description	 :	This function to perform operations on Java Table component.
'
'Function Parameters	 :  1. sFunctionName: Caller function's name
'							2. sAction		: Action to be performed
'							3. objContainer : Parent UI Component or Java Table
'							4. sJavaTable 	: Java Table
'							5. iRow 		: Row number
'							6. sColumn 		: Column name or Number
'							7. sValue 		: value to be set or type or verified
'							8. sMouseButton : Mouse button to click
'							9. sModifier 	: Modifier
'
'Function Return Value	 : 	True \ False\ Value
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","ClickCell",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","0","0","","LEFT","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","DoubleClickCell",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","0","0","","LEFT","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","SelectCell",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","0","1","","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","SelectRow",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","0","","","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","GetCellData",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","0","Vehicle Object","","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","SetCellData",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","0","Vehicle Object","Test","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","Type",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","","","0078","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","GetRowCount",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","","","","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","IsColumnExist",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","","Vehicle Object","","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","ExtendRow",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"VehicleProgramTable","0","","","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTable_Operations("","ClickLink",JavaWindow("FMDefaultWindow").JavaWindow("VehicleProgramDocumentSet"),"QuickLinksTable","","0","Favorites","","")
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale				|  03-Feb-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaTable_Operations(sFunctionName,sAction,objContainer,sJavaTable,iRow,sColumn,sValue,sMouseButton,sModifier)
	'variable declaration
	Dim aColumn
	Dim iCounter,iCount, iRowCount
	Dim sData
	Dim bFlag
	Dim objJavaTable
	'object creation
	If sJavaTable<>"" Then
		Set objJavaTable=objContainer.JavaTable(sJavaTable)	
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaTable_Operations  : [ " &  objContainer.toString & " ] : [ " &  sJavaTable & " ] : Action = " & sAction & " : "
	Else
		Set objJavaTable=objContainer
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaTable_Operations  : [ " & objJavaTable.toString() & " ] : Action = " & sAction & " : "	
	End If
	
	Fn_UI_JavaTable_Operations = False
	
	'Checking existance of Java Table
	If Fn_UI_Object_Operations("Fn_UI_JavaTable_Operations", "Enabled", objJavaTable,"", "", "") = False Then
		Set objJavaTable=Nothing
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Java Table not exist")
		Exit Function
	End If
		
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "clickcell"
			' get no. of rows in Table
			If Cint(objJavaTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Row : [ " & iRow &" ] not found in table")
			Else
				' click on cell with specified row and column
				If sMouseButton<>"" and sModifier<>"" Then
					objJavaTable.ClickCell iRow,sColumn,sMouseButton,sModifier
				ElseIf sMouseButton<>"" Then
					objJavaTable.ClickCell iRow,sColumn,sMouseButton
				Else
					objJavaTable.ClickCell iRow,sColumn
				End If
				If Err.number < 0 Then
					Fn_UI_JavaTable_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Click on cell at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn)&" ]")
				Else
					Fn_UI_JavaTable_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Click on cell at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn) &"]")
				End If
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "doubleclickcell"
			' get no. of rows in Table
			If Cint(objJavaTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Row : [ " & iRow &" ] not found in table")
			Else
			' Double Click on cell with specified row and column
				If sMouseButton<>"" and sModifier<>"" Then
					objJavaTable.DoubleClickCell iRow,sColumn,sMouseButton,sModifier
				ElseIf sMouseButton<>"" Then
					objJavaTable.DoubleClickCell iRow,sColumn,sMouseButton
				Else
					objJavaTable.DoubleClickCell iRow,sColumn
				End If
				If Err.number < 0 Then
					Fn_UI_JavaTable_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Click on cell at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn)&" ]")
				Else
					Fn_UI_JavaTable_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Click on cell at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn) &"]")
				End If
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "selectcell"
			' get no. of rows in Table
			If Cint(objJavaTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Row : [ " & Cstr(iRow)&" ] not found in table")
			Else
				'Select Cell with specified row and column
				objJavaTable.SelectCell iRow,sColumn
				If Err.number < 0 Then
					Fn_UI_JavaTable_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Select cell at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn)&" ]")
				Else
					Fn_UI_JavaTable_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Select cell at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn)&" ]")
				End If
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "selectrowext"
			' get no. of rows in Table
			iCount=Fn_UI_JavaTable_Operations(sFunctionName,"getcolumnindex",objContainer,sJavaTable,"",sColumn,"","","")
			For iCounter = 0 To Cint(objJavaTable.GetROProperty("rows"))-1
				If Trim(objJavaTable.GetCellData(iCounter,iCount))= Trim(sValue) Then
					objJavaTable.SelectRow iCounter
					Fn_UI_JavaTable_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Select row : [ " & Cstr(iCounter)&" ]")
					Exit For
				End If	
			Next
			
			If Fn_UI_JavaTable_Operations = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Select row of value [ " & Cstr(sValue) & " ] under column [ " & Cstr(sColumn) & " ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "selectrow"
			' get no. of rows in Table
			If Cint(objJavaTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Row : [ " & Cstr(iRow)&" ] not found in table")
			Else
				'Select row with specified row 
				objJavaTable.SelectRow iRow
				If Err.number < 0 Then
					Fn_UI_JavaTable_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Select row : [ " & Cstr(iRow)&" ]")
				Else
					Fn_UI_JavaTable_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Select row : [ " & Cstr(iRow)&" ]")
				End If
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "extendrow"
			' get no. of rows in Table
			If Cint(objJavaTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Row : [ " & Cstr(iRow)&" ] not found in table")
			Else
				' extend rows
				objJavaTable.ExtendRow iRow
				If Err.number < 0 Then
					Fn_UI_JavaTable_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Extend row : [ " & Cstr(iRow)&" ]")
				Else
					Fn_UI_JavaTable_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Extend row : [ " & Cstr(iRow)&" ]")
				End If
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "getcelldata"	
			' get no. of rows in Table
			If Cint(objJavaTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Row : [ " & Cstr(iRow)&" ] not found in table")
			Else
				' get value from cell
				sColumn = Cint(Fn_UI_JavaTable_Operations("","getcolumnindex",objJavaTable,"","",sColumn,"","",""))
				sData=objJavaTable.GetCellData(iRow,sColumn)
				If Err.number < 0 Then
					Fn_UI_JavaTable_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Get cell data at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn)&" ]")
				Else
					Fn_UI_JavaTable_Operations = sData
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Get cell data at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn)&" ]")
				End If
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "setcelldata"
			' get no. of rows in Table
			If Cint(objJavaTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Row : [ " & Cstr(iRow)&" ] not found in table")
			Else
				' set value in table cell
				objJavaTable.SetCellData iRow,sColumn,sValue
				If Err.number < 0 Then
					Fn_UI_JavaTable_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Set cell data at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn)&" ]")
				Else
					Fn_UI_JavaTable_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Set cell data at row : [ " & Cstr(iRow)&" ] and column [ "& Cstr(sColumn)&" ]")
				End If
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "type"
			'type value in table
			objJavaTable.Type(sValue)
			If Err.number < 0 Then
				Fn_UI_JavaTable_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Type [ " & Cstr(sValue)&" ]")
			Else
				Fn_UI_JavaTable_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Type [ " & Cstr(sValue)&" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "getrowcount"
			' get no. of rows in Table
			iRowCount=Cint(objJavaTable.GetROProperty("rows"))
			If Err.number < 0 Then
				Fn_UI_JavaTable_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to get row count ")
			Else
				Fn_UI_JavaTable_Operations = CStr(iRowCount)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully get row count [ " & CStr(iRowCount)&" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 				
		Case "getcolumnindex"
			bFlag=False
			' loop through column count to get column Names
			For iCount = 0 To Cint(objJavaTable.GetROProperty("cols"))-1
				If trim(objJavaTable.getColumnName(iCount))=Trim(sColumn) Then
					bFlag=True
					Fn_UI_JavaTable_Operations = iCount					
					Exit for
				End If
			Next
			if bFlag=False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL:  Column [ " & Cstr(sColumn)&" ] does not found in table")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Column [ " & Cstr(sColumn)&" ] Exist in table at position [ " & Cstr(iCounter) & " ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 				
		Case "iscolumnexist"
			aColumn=Split(sColumn,"~")
			For iCounter = 0 To Ubound(aColumn)
				bFlag=False
				' loop through column count to get column Names
				For iCount = 0 To Cint(objJavaTable.GetROProperty("cols"))-1
					If trim(objJavaTable.getColumnName(iCount))=Trim(aColumn(iCounter)) Then
						bFlag=True
						Exit for
					End If
				Next
				if bFlag=False Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL:  Column [ " & Cstr(sColumn)&" ] not found in table")
					Exit for
				End If				
			Next	
			if bFlag=True Then
				Fn_UI_JavaTable_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Column [ " & Cstr(sColumn)&" ] Exist in table")
			End If		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "clicklink"
			bFlag = False
			For iCounter = 0 To Cint(Fn_UI_Object_Operations("Fn_UI_JavaTable_Operations", "getroproperty", objJavaTable, "","rows",""))-1			
				If Trim(objJavaTable.GetCellData(iCounter,sColumn)) = Trim(sValue) Then
					bFlag = Fn_UI_JavaTable_Operations("Fn_UI_JavaTable_Operations","ClickCell",objJavaTable,"",iCounter,sColumn,"","","")
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS: Successfully clicked on Link [ " & Cstr(sValue)&" ]")
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				End If
			Next
			If bFlag = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL:  Link [ " & Cstr(sValue)&" ] not found in table")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "selectallrow"
			' get no. of rows in Table
			For iCounter = 0 To Cint(objJavaTable.GetROProperty("rows"))-1
				If iCounter = 0 Then
					objJavaTable.SelectRow iCounter
				Else
					objJavaTable.ExtendRow iCounter
				End If
			Next
			
			If Err.number < 0 Then
				Fn_UI_JavaTable_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to select all row ]")
			Else
				Fn_UI_JavaTable_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Select all row ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: invalid case.")		
	End Select
	
	If Err.Number <> 0 Then	
		Fn_UI_JavaTable_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Clear memory of JavaTable object.
	Set objJavaTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaTree_Operations
'
'Function Description	 :	This function to perform operations on Java tree component.
'
'Function Parameters	 :  1. sFunctionName: Caller function's name
'							2. sAction		: Action to be performed
'							3. objContainer : Parent UI Component or Java Tree
'							4. sJavaTree 	: Java Tree name
'							5. sNode 		: Tree Node path
'							6. sWinMenu 	: Win menu name
'							7. sPopupMenu 	: Popup menu to select
'
'Function Return Value	 : 	True \ False\ Value
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage			 :	bReturn = Fn_UI_JavaTree_Operations("","Expand",JavaWindow("MyTeamcenter"),"NavTree","Home~Newstuff~158801","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTree_Operations("","Collapse",JavaWindow("MyTeamcenter"),"NavTree","Home~Newstuff~158801","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTree_Operations("","Select",JavaWindow("MyTeamcenter"),"NavTree","Home~Newstuff~158801","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTree_Operations("","Deselect",JavaWindow("MyTeamcenter"),"NavTree","Home~Newstuff~158801","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTree_Operations("","Multiselect",JavaWindow("MyTeamcenter"),"NavTree","Home~Newstuff~158801^Home~Newstuff~12345/A","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTree_Operations("","OpenContextMenu",JavaWindow("MyTeamcenter"),"NavTree","Home~Newstuff~158801","","")
'Function Usage			 :	bReturn = Fn_UI_JavaTree_Operations("","Doubleclick",JavaWindow("MyTeamcenter"),"NavTree","Home~Newstuff~158801","","")
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale				|  03-Feb-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaTree_Operations(sFunctionName,sAction,objContainer,sJavaTree,sNode,sWinMenu,sPopupMenu)
	' declaration of variables
	Dim aNode
	Dim iCounter
	Dim objJavaTree
	
	' object creation
	If sJavaTree<>"" Then
		Set objJavaTree=objContainer.JavaTree(sJavaTree)	
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaTree_Operations  : [ " &  objContainer.toString & " ] : [ " &  sJavaTree & " ] : Action = " & sAction & " : "
	Else
		Set objJavaTree=objContainer
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaTree_Operations  : [ " & objJavaTree.toString() & " ] : Action = " & sAction & " : "	
	End If
	
	Fn_UI_JavaTree_Operations = False
	
	'Checking existance of Java Tree
	If Fn_UI_Object_Operations("Fn_UI_JavaTree_Operations", "Enabled", objJavaTree,"", "", "") = False Then
		Set objJavaTree=Nothing
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Java Tree not exist")
		Exit Function
	End If
		
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "expand"	
			'expand node
			objJavaTree.Expand sNode		
			If Err.number < 0 Then
				Fn_UI_JavaTree_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Expand node : " & sNode)
			Else
				Fn_UI_JavaTree_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Expand node : " & sNode)
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "collapse"	
			'collapse node
			objJavaTree.Collapse sNode		
			If Err.number < 0 Then
				Fn_UI_JavaTree_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Collapse node : " & sNode)
			Else
				Fn_UI_JavaTree_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Collapse node : " & sNode)
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "select"
			'select node
			objJavaTree.Select sNode		
			If Err.number < 0 Then
				Fn_UI_JavaTree_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Select node : " & sNode)
			Else
				Fn_UI_JavaTree_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Select node : " & sNode)
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "deselect"	
			'deselect tree node
			objJavaTree.Deselect sNode		
			If Err.number < 0 Then
				Fn_UI_JavaTree_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Deselect node : " & sNode)
			Else
				Fn_UI_JavaTree_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Deselect node : " & sNode)
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "multiselect"	
			'multiselect tree nodes
			aNode=Split(sNode,"^")
			objJavaTree.Select aNode(0)		
			If Err.number < 0 Then
				Fn_UI_JavaTree_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Multiselect node : " & sNode)
			Else
				For iCounter = 1 To Ubound(aNode)
					objJavaTree.ExtendSelect aNode(iCounter)
					wait GBL_MICRO_TIMEOUT
				Next
				Fn_UI_JavaTree_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Multiselect node : " & sNode)
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "opencontextmenu"	
			' open context menu on specified node
			objJavaTree.OpenContextMenu(sNode)
			If Err.number < 0 Then
				Fn_UI_JavaTree_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Select node : " & sNode)
			Else
				Fn_UI_JavaTree_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Select node : " & sNode)
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "doubleclick","activate"
			' activate or doubleclick on tree node
			objJavaTree.Activate sNode		
			If Err.number < 0 Then
				Fn_UI_JavaTree_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Doubleclick/Activate node : " & sNode)
			Else
				Fn_UI_JavaTree_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Doubleclick/Activate node : " & sNode)
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: invalid case.")		
	End Select
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	If Err.Number <> 0 Then	
		Fn_UI_JavaTree_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Releasing Object
	Set objJavaTree=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_UI_Object_GetChildObjects
'
'Function Description	:	Function used to get child objects of specified component descriptively
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.objComponent: Reference path of object
'							3.sProperties: Property Name of object
'							4.sValues: Property Value of object
'
'Function Return Value	 : 	Nothing \ Array of Objects
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Set objDialog =  Fn_UI_Object_GetChildObjects("", objDialog, "Class Name~tagname", "JavaObject~ImageHyperlink")
'Function Usage		     :	Set objDialog =  Fn_UI_Object_GetChildObjects("", objDialog, "Class Name$RegularExpression~tagname", "Java.*~ImageHyperlink")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  24-Feb-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_Object_GetChildObjects(sFunctionName, objComponent, sProperties, sValues)
	'Declaring Variables
	Dim objSelectType, intNoOfObjects
	Dim arrProperties, arrValues, arrCounter
	
	'Initially set function return value as Nothing
	Set Fn_UI_Object_GetChildObjects = Nothing
	
	'Creating descriptive object of select type
	Set objSelectType = Description.Create()
	
	'Split the properties and values
	arrProperties = split(sProperties, "~")
	arrValues = split(sValues, "~") 
	
	GBL_FUNCTIONLOG = " [ " +  objComponent.toString + " ] : "

	For arrCounter = 0 to UBound(arrProperties)
		If instr(arrProperties(arrCounter),"$RegularExpression") > 0 Then
			objSelectType(replace(arrProperties(arrCounter), "$RegularExpression", "")).RegularExpression = True
			objSelectType(replace(arrProperties(arrCounter), "$RegularExpression", "")).value = arrValues(arrCounter)
		Else
			objSelectType(arrProperties(arrCounter)).value = arrValues(arrCounter)
		End If
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_Object_GetChildObjects >> PASS : "& GBL_FUNCTIONLOG &"Setting Properties [" & arrProperties(arrCounter) & " = " & arrValues(arrCounter) & " ].")		
	Next
	If Fn_UI_Object_Operations("Fn_UI_Object_GetChildObjects","Exist", objComponent, GBL_MIN_TIMEOUT,"","") Then
		Set intNoOfObjects = objComponent.ChildObjects(objSelectType)
		If intNoOfObjects.count <> 0 Then
			Set Fn_UI_Object_GetChildObjects = intNoOfObjects
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_Object_GetChildObjects >> PASS : "& GBL_FUNCTIONLOG &"Successfully found [" & intNoOfObjects.count & "] child objects.")		
		Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_Object_GetChildObjects >> FAIL : "& GBL_FUNCTIONLOG &"No child objects found.")		
		End If
	End If
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Set Fn_UI_Object_GetChildObjects = Nothing
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_Object_GetChildObjects >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release Objects
	Set intNoOfObjects = Nothing
	Set objSelectType = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_UI_JavaMenu_Operations
'
'Function Description	:	Function used to perform operations on JavaMenu component
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action to be performed
'							3.objContainer: Parent UI Component or JavaRadioButton object
'							4.sMenuPath: Menu path
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_UI_JavaMenu_Operations("","Select",JavaWindow("DefaultWindow"),"File:New:Folder...")
'Function Usage		     :	Call Fn_UI_JavaMenu_Operations("","Exist",JavaWindow("DefaultWindow"),"File:New:Folder...")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  25-Feb-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaMenu_Operations(sFunctionName,sAction,objContainer,sMenuPath)
	'Declaring Variables
	Dim arrMenu
	Dim iCounter
	Dim bExist
	
	'Initially set function return value as False
	Fn_UI_JavaMenu_Operations = False
		 
	GBL_FUNCTIONLOG =" [ " &  objContainer.toString & " ] : Menu = [ " +  sMenuPath + " ] : Action = " & sAction & " : "
		
	'Checking existance of parent window
	If Fn_UI_Object_Operations("Fn_UI_JavaMenu_Operations", "Enabled", objContainer,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Java Menu parent component is not enabled")
		Exit Function
	End If
	
	'Split the menu
	arrMenu = Split(sMenuPath,":")
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select the java menu
		Case "select"
			iCounter = ubound(arrMenu)						
			Select Case iCounter
				Case "0"				
					objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0").Select
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected Java Menu [ " & sMenuPath & " ].")
					Fn_UI_JavaMenu_Operations = True
				Case "1"				
					objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0").JavaMenu("label:="& arrMenu(1)&"","index:=0").Select
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected Java Menu [ " & sMenuPath & " ].")
					Fn_UI_JavaMenu_Operations = True
				Case "2"
					objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0").JavaMenu("label:="& arrMenu(1)&"","index:=0").JavaMenu("label:="& arrMenu(2)&"","index:=0").Select
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected Java Menu [ " & sMenuPath & " ].")
					Fn_UI_JavaMenu_Operations = True
				 Case "3"
					objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0").JavaMenu("label:="& arrMenu(1)&"","index:=0").JavaMenu("label:="& arrMenu(2)&"","index:=0").JavaMenu("label:="& arrMenu(3)&"","index:=0").Select
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected Java Menu [ " & sMenuPath & " ].")
					Fn_UI_JavaMenu_Operations = True
				Case "4"
					objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0").JavaMenu("label:="& arrMenu(1)&"","index:=0").JavaMenu("label:="& arrMenu(2)&"","index:=0").JavaMenu("label:="& arrMenu(3)&"","index:=0").JavaMenu("label:="& arrMenu(4)&"","index:=0").Select
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected Java Menu [ " & sMenuPath & " ].")
					Fn_UI_JavaMenu_Operations = True
				Case Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> FAIL : [No valid case] : No valid case was passed to select Java Menu [ sMenuPath ] for function [Fn_UI_JavaMenu_Operations]")
			End Select
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verfy the java menu exist
		Case "exist"
			iCounter = ubound(arrMenu)
			bExist = False			
			Select Case iCounter
				Case "0"				
					bExist = objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0","path:=MenuItem;Shell;").Exist(GBL_DEFAULT_MIN_TIMEOUT)
				Case "1"				
					bExist = objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0").JavaMenu("label:="& arrMenu(1)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)
				Case "2"
					bExist = objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0").JavaMenu("label:="& arrMenu(1)&"","index:=0").JavaMenu("label:="& arrMenu(2)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)
				 Case "3"
					bExist = objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0").JavaMenu("label:="& arrMenu(1)&"","index:=0").JavaMenu("label:="& arrMenu(2)&"","index:=0").JavaMenu("label:="& arrMenu(3)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)
				Case "4"
					bExist = objContainer.JavaMenu("label:="& arrMenu(0)&"","index:=0").JavaMenu("label:="& arrMenu(1)&"","index:=0").JavaMenu("label:="& arrMenu(2)&"","index:=0").JavaMenu("label:="& arrMenu(3)&"","index:=0").JavaMenu("label:="& arrMenu(4)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)
				Case Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> FAIL : [No valid case] : No valid case was passed to verify Java Menu [ sMenuPath ] for function [Fn_UI_JavaMenu_Operations]")
			End Select			
			If bExist Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> PASS : "& GBL_FUNCTIONLOG &"Java Menu [ " & sMenuPath & " ] exist.")
				Fn_UI_JavaMenu_Operations = True
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Java Menu [ " & sMenuPath & " ] does not exist.")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_UI_JavaMenu_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_UI_JavaMenu_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaMenu_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")		
	End If
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_UI_JavaObject_Operations
'
'Function Description	:	Function used to perform operations on Java Object component
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action to be performed
'							3.objContainer: Parent UI Component or Java object
'							4.sJavaObject: Name of Java Object
'							5.iX: X coordinate value
'							6.iY: Y coordinate value
'							7.sMouseButton: Mouse button Name [ LEFT | RIGHT ]
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_UI_JavaObject_Operations("","Click",JavaWindow("DefaultWindow"),"RACTabFolderWidget","10","10","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  25-Feb-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaObject_Operations(sFunctionName,sAction,objContainer,sJavaObject,iX,iY,sMouseButton)	
	'Declaring Variables
	Dim objJavaObject
	
	'Initially set function return value as False
	Fn_UI_JavaObject_Operations = False
	
	'Creating/Setting Object of Java Object
	If sJavaObject <> "" Then
		Set objJavaObject = objContainer.JavaObject(sJavaObject)	
		GBL_FUNCTIONLOG = " [ " &  objContainer.toString & " ] : [ " +  sJavaObject + " ] : Action = " & sAction & " : "
	Else
		Set objJavaObject = objContainer
		GBL_FUNCTIONLOG = " [ " &  objJavaObject.toString & " ] : Action = " & sAction & " : "
	End If	
	
	'Checking existance of Java Object
	If Fn_UI_Object_Operations("Fn_UI_JavaObject_Operations", "Exist", objJavaObject,"","","") = False Then
		'Release Object
		Set objJavaObject = Nothing
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Java object does not exist")
		Exit Function
	End If
		
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click on the java object
		Case "click"
			If Fn_UI_Object_Operations("Fn_UI_JavaObject_Operations", "Enabled", objJavaObject,"","","") = False Then
				'Release Object
				Set objJavaObject = Nothing
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Java object is not enabled")
				Exit Function
			End If		
			'Perform click			
			If sMouseButton <> "" Then
				objJavaObject.Click iX,iY,sMouseButton
			Else
				objJavaObject.Click iX,iY			
			End If
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully clicked on Java Object.")
			Fn_UI_JavaObject_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaObject_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_UI_JavaObject_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_UI_JavaObject_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_JavaObject_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")		
	End If
	
	'Release Object
	Set objJavaObject = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_UI_JavaToolbar_Operations
'
'Function Description	 :	This function to perform operations on Java toolbar component
'
'Function Parameters	 :  1. sFunctionName 		: Caller function's Name
'							2. sAction				: Action to be performed
'							3. objJavaDialog		: Parent UI Component or Java Toolbar
'							4. sToolBarName 		: Toolbar's logical name from OR.
'							5. sToobarButtonName 	: Toolbar button Name
'							6. sValue 				: For future purpose
'							7. sMenu 				: popup / dropdown menu
'							8. iIndex 				: instance number
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage			 :	bReturn = Fn_UI_JavaToolbar_Operations("", "Click", JavaWindow("ProductMasterManager"), "", "Clear", "", "", 2)
'Function Usage			 :	bReturn = Fn_UI_JavaToolbar_Operations("", "DropdownMenuSelect", JavaWindow("ProductMasterManager"), "", "Part Type", "", "Standard Part", "")
'Function Usage			 :	bReturn = Fn_UI_JavaToolbar_Operations("", "DropdownMenuSelect", JavaWindow("ProductMasterManager"), "TabFolderWidgetToolBar", "Part Type", "", "Standard Part", "")
'Function Usage			 :	bReturn = Fn_UI_JavaToolbar_Operations("", "OpenDropdownMenu", JavaWindow("ProductMasterManager"), "TabFolderWidgetToolBar", "Part Type", "", "", "")
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane			|  08-Jun-2016	    |	 1.0		|	Kundan Kudale	  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_UI_JavaToolbar_Operations(sFunctionName, sAction, objJavaDialog, sToolBarName, sToobarButtonName, sValue, sMenu, iIndex)
	Err.Clear
	'Declaring variables
	Dim objJavaToolbar,objObjects
	Dim iCounter
	Dim sContents
	Dim bFlag
	
	'Initially set function return value as False
	Fn_UI_JavaToolbar_Operations = False
	
	GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_JavaToolbar_Operations : [ " & objJavaDialog.toString() & " ] : Action = " & sAction & " : Toolbar Button = " & sToobarButtonName & " : "
	
	If sToolBarName = "" Then
		If iIndex <> "" Then
			iIndex = cInt(iIndex)
		Else
			iIndex = 1
		End If
		bFlag = False
		Set objObjects = Fn_UI_Object_GetChildObjects( "Fn_UI_JavaToolbar_Operations", objJavaDialog, "Class Name~enabled", "JavaToolbar~1")
		If typename(objObjects ) <> "Nothing" Then
			For iCounter = 0 to objObjects.Count - 1
                sContents = objObjects(iCounter).GetContent()
				If  Fn_CommonUtil_ArrayStringContains(sContents, sToobarButtonName, ";") Then
					If iIndex = 1 Then
						Set objJavaToolbar = objObjects(iCounter)
						bFlag = True
						Exit for
					Else
						iIndex = iIndex - 1
					End If
				End If
			Next
		End If
		If bFlag = False Then
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Failed to find toolbar [ " & sToobarButtonName & " ].")
			Exit function
		End If
		Set objObjects =Nothing
	Else
		Set objJavaToolbar = objJavaDialog.JavaToolbar(sToolBarName)
	End If
	
	If Fn_UI_Object_Operations("Fn_UI_JavaToolbar_Operations", "Exist", objJavaToolbar,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Failed to find toolbar button [ " & sToobarButtonName & " ].")
		Exit Function
	End If
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to click on toolbar button
		Case "Click"
			objJavaToolbar.Press sToobarButtonName
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully clicked on [ " & sToobarButtonName & " ].")
			Fn_UI_JavaToolbar_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to select popup menu from toolbar button
		Case "DropdownMenuSelect"
			objJavaToolbar.ShowDropdown sToobarButtonName
			wait GBL_MICRO_TIMEOUT
			Fn_UI_JavaToolbar_Operations = Fn_UI_JavaMenu_Operations("Fn_UI_JavaToolbar_Operations","Select", objJavaDialog,sMenu)
			If Fn_UI_JavaToolbar_Operations Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully clicked on [ " & sToobarButtonName & " ] and [ " & sMenu & " ] selected.")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Failed to click on [ " & sToobarButtonName & " ] and select [ " & sMenu & " ].")
			End If				
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to open toolbar button dropdown box
		Case "OpenDropdownMenu"
			objJavaToolbar.ShowDropdown sToobarButtonName
			wait GBL_MICRO_TIMEOUT
			Fn_UI_JavaToolbar_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: invalid case.")		
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_UI_Object_GetChildObjects >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Clear memory of JavaToolbar object
	Set objJavaToolbar=Nothing
End Function
