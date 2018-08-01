Option Explicit

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Function Name								|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. Fn_SAP_UI_SAPObject_Operations					|	sandeep.navghane@sqs.com	|	05-Jul-2016	|	Function used to perform operations on SAP objects
'002. Fn_SAP_UI_SAPGuiEdit_Operations					|	sandeep.navghane@sqs.com	|	05-Jul-2016	|	Function used to perform operations on SAPGuiEdit component
'003. Fn_SAP_UI_SAPGuiButton_Operations					|	sandeep.navghane@sqs.com	|	05-Jul-2016	|	Function is used to perform operations on SAPGuiButton object
'004. Fn_SAP_UI_SAPGuiOKCode_Operations					|	sandeep.navghane@sqs.com	|	05-Jul-2016	|	Function used to perform operations on SAPGuiOKCode component
'005. Fn_SAP_UI_SAPGuiTable_Operations					|	sandeep.navghane@sqs.com	|	05-Jul-2016	|	Function used to perform operations on SAPGuiTable component
'006. Fn_SAP_UI_SAPGuiTabStrip_Operations				|	sandeep.navghane@sqs.com	|	05-Jul-2016	|	Function used to perform operations on SAPGuiTabStrip component
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_SAP_UI_SAPObject_Operations
'
'Function Description	:	Function used to perform operations on SAP objects.
'
'Function Parameters	:   1.sFunctionName		: Name of function
'						    2.sAction			: Action name
'							3.objReferencePath	: Reference path of Object
'							4.iTimeOut			: Time out time in seconds
'							5.sPropertyName		: Valid Property Name
'							6.sPropertyValue	: Property value
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_SAP_UI_SAPObject_Operations("Fn_DateControl_SetDate","create", objDateControl,"","","")
'Function Usage		     :	Call Fn_SAP_UI_SAPObject_Operations("Fn_DateControl_SetDate","enabled", objDateControl, MAX_TIMEOUT,"","")
'Function Usage		     :	Call Fn_SAP_UI_SAPObject_Operations("Fn_DateControl_SetDate","exist", objDateControl,"","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  05-Jul-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_SAP_UI_SAPObject_Operations(sFunctionName, sAction, objReferencePath, iTimeOut, sPropertyName, sPropertyValue)
	Err.Clear	
	'Initially set function return value as False
	Fn_SAP_UI_SAPObject_Operations = False
	
	GBL_FUNCTIONLOG = sFunctionName & " > Fn_UI_Object_Operations : [ " &  objReferencePath.toString & " ] : Action = " & sAction & " : "
	
	'Set the default time out
	If iTimeOut = "" Then
		iTimeOut = GBL_DEFAULT_TIMEOUT
	Else
		iTimeOut = cInt(iTimeOut)
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to create a SAPObject
		Case "create"
			'Verify if the reference path of object is enabled
			If Fn_SAP_UI_SAPObject_Operations(sFunctionName, "Enabled", objReferencePath, iTimeOut,"","") Then
				'Returning SAPObject
				Set Fn_SAP_UI_SAPObject_Operations = objReferencePath
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully created new object for [ " & objReferencePath.ToString & " ]")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &" SAPObject is not enabled")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify existance of SAPObject
		Case "exist"
			Fn_SAP_UI_SAPObject_Operations = objReferencePath.Exist(iTimeOut)
			If Fn_SAP_UI_SAPObject_Operations Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully verified object [ "& objReferencePath.ToString &" ] exist")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &"SAPObject does not exist")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify SAPObject is enabled
		Case "enabled"
			'Verify if the reference path of object is exist
			If Fn_SAP_UI_SAPObject_Operations(sFunctionName, "Exist", objReferencePath, iTimeOut,"","") Then
				If objReferencePath.GetROProperty("enabled") = "1"  OR objReferencePath.GetROProperty("enabled") = True Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully verified object [ "& objReferencePath.ToString &" ] exist and enabled")
					Fn_SAP_UI_SAPObject_Operations = True
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &"SAPObject is not enabled")
				End If
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &"SAPObject does not exist")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the property value
		Case "getroproperty"
				Fn_SAP_UI_SAPObject_Operations = objReferencePath.GetROProperty(sPropertyName)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFunctionName&" >> Fn_SAP_UI_SAPObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Sucessfully fetched  property value [ " & sPropertyName & " ] of object [ "& objReferencePath.ToString &" ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set the property value
		Case "settoproperty"
				objReferencePath.SetTOProperty sPropertyName,sPropertyValue
				Fn_SAP_UI_SAPObject_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Sucessfully set  property [ " & sPropertyName & " ] of object [ "& objReferencePath.ToString &" ] with value [ " & sPropertyName & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_SAP_UI_SAPObject_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_SAP_UI_SAPObject_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_SAP_UI_SAPObject_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_SAP_UI_SAPGuiEdit_Operations
'
'Function Description	 :	This function to perform operations on SAPGuiEdit component.
'
'Function Parameters	 :  1. sFunctionName	: Function name from which this function is called.
'							2. sAction			: Action to be performed
'							3. objSAPDialog 	: Parent UI Component or SAPGuiEdit object
'							4. sSAPGuiEdit		: SAPGuiEdit Control name
'							5. sText			: text to be set in edit box
'
'Function Return Value	 : 	True \ False \ text in edit box
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiEdit_Operations("Fn_SISW_PMM_UserContextSettings", "Set",  objDialog, "EWO", "EWO_Name" )
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiEdit_Operations("Fn_SISW_PMM_UserContextSettings", "Type",  objDialog.SAPGuiEdit("EWO"),"", "EWO_Name" )
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiEdit_Operations("Fn_SISW_PMM_UserContextSettings", "activate",  objDialog.SAPGuiEdit("EWO"),"", "" )
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiEdit_Operations("Fn_SISW_PMM_UserContextSettings", "setsecure",  objDialog.SAPGuiEdit("EWO"),"", "EWO_Name" )
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiEdit_Operations("Fn_SISW_PMM_UserContextSettings", "settext",  objDialog.SAPGuiEdit("EWO"),"", "EWO_Name" )
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiEdit_Operations("Fn_SISW_PMM_UserContextSettings", "gettext",  objDialog.SAPGuiEdit("EWO"),"", "" )
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  05-Jul-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_SAP_UI_SAPGuiEdit_Operations(sFunctionName, sAction, objSAPDialog, sSAPGuiEdit, sText)
	'Declaring Variables
	Dim objSAPGuiEdit
	
	'Initially set function return value as False	
	Fn_SAP_UI_SAPGuiEdit_Operations = False
	
	'Set an Edit Object on variable
	If sSAPGuiEdit <> "" Then
		Set objSAPGuiEdit= objSAPDialog.SAPGuiEdit(sSAPGuiEdit)
		GBL_FUNCTIONLOG = sFunctionName & "> Fn_SAP_UI_SAPGuiEdit_Operations : [ " &  objSAPDialog.toString & " ] : [ " & objSAPGuiEdit.toString & " ] : Action = " & sAction & " : "
	Else
		Set objSAPGuiEdit= objSAPDialog
		GBL_FUNCTIONLOG = sFunctionName & "> Fn_SAP_UI_SAPGuiEdit_Operations : [ " & objSAPGuiEdit.toString & " ] : Action = " & sAction & " : "
	End If

	If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiEdit_Operations", "Exist", objSAPGuiEdit,"", "", "") = False Then
		Fn_SAP_UI_SAPGuiEdit_Operations = False
		Set objSAPGuiEdit = Nothing 
		Exit Function
	End If
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "set"
			'Setting the editbox
			If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiEdit_Operations", "Enabled", objSAPGuiEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL :[ " &  objSAPGuiEdit.toString & " ] object Does not exist.")
				Set objSAPGuiEdit = Nothing 
				Exit Function
			End If
			' set value in edit box
			objSAPGuiEdit.Set sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Text " & sText  &" is Set/ Entered in SAPGuiEditBox.")
			Fn_SAP_UI_SAPGuiEdit_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "type"
			If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiEdit_Operations", "Enabled", objSAPGuiEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objSAPGuiEdit.toString & " ] does is not enabled.")
				Set objSAPGuiEdit = Nothing 
				Exit Function
			End If
			'type the value in editbox
			objSAPGuiEdit.Type sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Text " & sText  &" is Set/ Entered in SAPGuiEditBox.")
			Fn_SAP_UI_SAPGuiEdit_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "activate"
			If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiEdit_Operations", "Enabled", objSAPGuiEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objSAPGuiEdit.toString & " ] object does not exist.")
				Set objSAPGuiEdit = Nothing 
				Exit Function
			End If
			'Activate the editbox
			objSAPGuiEdit.Activate
			Fn_SAP_UI_SAPGuiEdit_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "gettext"
			' get value from edit box
			Fn_SAP_UI_SAPGuiEdit_Operations = objSAPGuiEdit.getROProperty("value")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "setsecure"
			'Setting the edit box
			If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiEdit_Operations", "Enabled", objSAPGuiEdit,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objSAPGuiEdit.toString & " ] object Does not exist.")
				Set objSAPGuiEdit = Nothing 
				Exit Function
			End If
			' set secure value(encoded text) in edit box
			objSAPGuiEdit.SetSecure sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Text " & sText  &"is Set/ Entered in SAPGuiEditBox.")
			Fn_SAP_UI_SAPGuiEdit_Operations= True	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : Invalid Case.")
			Set objSAPGuiEdit = Nothing 
			Exit Function
	End Select
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Set objSAPGuiEdit = Nothing 
	If Err.Number <> 0 Then	
		Fn_SAP_UI_SAPGuiEdit_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_SAP_UI_SAPGuiButton_Operations
'
'Function Description	 :	This function is used to perform operations on SAPGuiButton object
'
'Function Parameters	 :  1.sFunctionName	: Function name from which this function is called.
'							2.sAction		: Action to be performed on SAPGuiButton
'							3.objSAPDialog	: Parent UI Component or SAPGuiButton object
'							4.sSAPGuiButton	: Valid SAPGuiButton name
'
'Function Return Value	 : 	True \ False 
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiButton_Operations("Fn_TeamcenterLogin", "Click",objSAPDialog,"Login")
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiButton_Operations("Fn_TeamcenterLogin", "Click", objSAPDialog.SAPGuiButton("Login"),"")
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiButton_Operations("Fn_TeamcenterLogin", "DeviceReplay.Click", objSAPDialog.SAPGuiButton("Login"),"")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale			    |  03-Feb-2016	    |	 1.0		|	Sandeep Navghane 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_SAP_UI_SAPGuiButton_Operations(sFunctionName, sAction, objSAPDialog, sSAPGuiButton)
	'Declaring Variables
	Dim objSAPGuiButton, objDeviceReplay
	
	'Initially set function return value as False
	Fn_SAP_UI_SAPGuiButton_Operations = False
	
	'Object Creation
	If sSAPGuiButton <> "" Then
		Set objSAPGuiButton = objSAPDialog.SAPGuiButton(sSAPGuiButton)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_SAP_UI_SAPGuiButton_Operations  : [ " &  objSAPDialog.toString & " ] : [ " &  objSAPGuiButton.toString & " ] : Action = " & sAction & " : "
	Else
		Set objSAPGuiButton = objSAPDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_SAP_UI_SAPGuiButton_Operations  : [ " &  objSAPGuiButton.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify SAPGuiButton object exists
	If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiButton_Operations", "Enabled", objSAPGuiButton,"", "", "") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objSAPGuiButton.toString & " ] object Does not exist.")
		Set objSAPGuiButton = Nothing 
		Exit Function
	End If
	
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "click"
			'clcik on SAPGuiButton
			objSAPGuiButton.Click	
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully clicked on SAPGuiButton.")
			Fn_SAP_UI_SAPGuiButton_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "devicereplay.click"
			'to click on Sapgui button using devicereplay method
			If sSAPGuiButton <> "" Then
				objSAPDialog.Activate
			End If
			'create DeviceReplay object
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			'move mouce to SAPGuiButton
			objDeviceReplay.MouseMove (objSAPGuiButton.GetROProperty("abs_x") & 5), (objSAPGuiButton.GetROProperty("abs_y") & 5)
			'click on Sapgui button
			objDeviceReplay.MouseClick  (objSAPGuiButton.GetROProperty("abs_x") & 5), (objSAPGuiButton.GetROProperty("abs_y") & 5), 0
		    Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully clicked on SAPGuiButton.")
			Fn_SAP_UI_SAPGuiButton_Operations = True
			'Clear memory of SAPGuiButton object.
			Set objDeviceReplay = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "[ " & GBL_FUNCTIONLOG & " ] : FAIL : Invalid Case.")
			Set objSAPGuiButton = Nothing 
			Exit Function
	End Select
	' log for error
	If Err.Number <> 0 Then	
		Fn_SAP_UI_SAPGuiButton_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Clear memory of SAPGuiButton object.
	Set objSAPGuiButton = Nothing 
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_SAP_UI_SAPGuiOKCode_Operations
'
'Function Description	 :	Function used to perform operations on SAPGuiOKCode component.
'
'Function Parameters	 :  1. sFunctionName	: Function name from which this function is called.
'							2. sAction			: Action to be performed
'							3. objSAPDialog 	: Parent UI Component or SAPGuiOKCode object
'							4. sSAPGuiOKCode	: SAPGuiOKCode Control name
'							5. sText			: text to be set in edit box
'
'Function Return Value	 : 	True \ False \ text in edit box
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiOKCode_Operations("Fn_SISW_PMM_UserContextSettings", "Set",  objDialog, "EWO", "EWO_Name" )
'Function Usage		     :  bReturn = Fn_SAP_UI_SAPGuiOKCode_Operations("Fn_SISW_PMM_UserContextSettings", "gettext",  objDialog.SAPGuiOKCode("EWO"),"", "" )
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  05-Jul-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_SAP_UI_SAPGuiOKCode_Operations(sFunctionName, sAction, objSAPDialog, sSAPGuiOKCode, sText)
	'Declaring Variables
	Dim objSAPGuiOKCode
	
	'Initially set function return value as False	
	Fn_SAP_UI_SAPGuiOKCode_Operations = False
	
	'Set an Edit Object on variable
	If sSAPGuiOKCode <> "" Then
		Set objSAPGuiOKCode= objSAPDialog.SAPGuiOKCode(sSAPGuiOKCode)
		GBL_FUNCTIONLOG = sFunctionName & "> Fn_SAP_UI_SAPGuiOKCode_Operations : [ " &  objSAPDialog.toString & " ] : [ " & objSAPGuiOKCode.toString & " ] : Action = " & sAction & " : "
	Else
		Set objSAPGuiOKCode= objSAPDialog
		GBL_FUNCTIONLOG = sFunctionName & "> Fn_SAP_UI_SAPGuiOKCode_Operations : [ " & objSAPGuiOKCode.toString & " ] : Action = " & sAction & " : "
	End If

	If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiOKCode_Operations", "Exist", objSAPGuiOKCode,"", "", "") = False Then
		Fn_SAP_UI_SAPGuiOKCode_Operations = False
		Set objSAPGuiOKCode = Nothing 
		Exit Function
	End If
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "set"
			'Setting the editbox
			If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiOKCode_Operations", "Enabled", objSAPGuiOKCode,"", "", "") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL :[ " &  objSAPGuiOKCode.toString & " ] object Does not exist.")
				Set objSAPGuiOKCode = Nothing 
				Exit Function
			End If
			' set value in edit box
			objSAPGuiOKCode.Set sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Text " & sText  &" is Set/ Entered in SAPGuiOKCodeBox.")
			Fn_SAP_UI_SAPGuiOKCode_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "gettext"
			' get value from edit box
			Fn_SAP_UI_SAPGuiOKCode_Operations = objSAPGuiOKCode.getROProperty("value")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : Invalid Case.")
			Set objSAPGuiOKCode = Nothing 
			Exit Function
	End Select
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Set objSAPGuiOKCode = Nothing 
	If Err.Number <> 0 Then	
		Fn_SAP_UI_SAPGuiOKCode_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_SAP_UI_SAPGuiTable_Operations
'
'Function Description	 :	Function used to perform operations on SAPGuiTable component.
'
'Function Parameters	 :  1. sFunctionName: Caller function's name
'							2. sAction		: Action to be performed
'							3. objContainer : Parent UI Component or SAPGuiTable
'							4. sSAPGuiTable : SAPGuiTable
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
'Function Usage			 :	bReturn = Fn_SAP_UI_SAPGuiTable_Operations("","ClickCell",objContainer,"VehicleProgramTable","","SAP Coce","Temp","","")
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  05-Jul-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_SAP_UI_SAPGuiTable_Operations(sFunctionName,sAction,objContainer,sSAPGuiTable,iRow,sColumn,sValue,sMouseButton,sModifier)
	'variable declaration
	Dim iCounter,iRowCount
	Dim objSAPGuiTable
	
	'object creation
	If sSAPGuiTable<>"" Then
		Set objSAPGuiTable=objContainer.SAPGuiTable(sSAPGuiTable)	
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_SAP_UI_SAPGuiTable_Operations  : [ " &  objContainer.toString & " ] : [ " &  sSAPGuiTable & " ] : Action = " & sAction & " : "
	Else
		Set objSAPGuiTable=objContainer
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_SAP_UI_SAPGuiTable_Operations  : [ " & objSAPGuiTable.toString() & " ] : Action = " & sAction & " : "	
	End If
	
	Fn_SAP_UI_SAPGuiTable_Operations = False
	
	'Checking existance of SAPGuiTable
	If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiTable_Operations", "Enabled", objSAPGuiTable,"", "", "") = False Then
		Set objSAPGuiTable=Nothing
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: SAPGuiTable does not exist")
		Exit Function
	End If
		
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "getrowindex"
			Fn_SAP_UI_SAPGuiTable_Operations = -1
			For iCounter = 1 to objSAPGuiTable.GetROProperty("rowcount")
				sTemp = objSAPGuiTable.GetCellData(iCounter ,sColumn)
				If Trim(objSAPGuiTable.GetCellData(iCounter ,sColumn)) = Trim(sValue) Then
					Fn_SAP_UI_SAPGuiTable_Operations = iCounter
					Exit for
				End If			
			Next
			If Fn_SAP_UI_SAPGuiTable_Operations = -1 Then				
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : Value [ " & Cstr(sValue) & " ] does not found in table")
			End if
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "selectcell"
			iRowCount=Fn_SAP_UI_SAPGuiTable_Operations("Fn_SAP_UI_SAPGuiTable_Operations","getrowindex",objContainer,sSAPGuiTable,"",sColumn,sValue,"","")
			If iRowCount<>-1 Then
				'Select Cell with specified row and column
				objSAPGuiTable.SelectCell iRowCount,sColumn
				If Err.number < 0 Then
					Fn_SAP_UI_SAPGuiTable_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: Fail to Select cell at row : [ " & Cstr(iRowCount)&" ] and column [ " & Cstr(sColumn) & " ] due to error [ " & Cstr(err.description) & " ]")
				Else
					Fn_SAP_UI_SAPGuiTable_Operations = True
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS: Successfully Select cell at row : [ " & Cstr(iRowCount) & " ] and column [ "& Cstr(sColumn)&" ]")
				End If
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "cellexist"		
			iRowCount = sSAPGuiTable.FindRowByCellContent(sColumn,sValue)
			If iRow = 0 Then
				Fn_SAP_UI_SAPGuiTable_Operations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : Value [ " & Cstr(sValue) & " ] does not found in table")
			Else
				Fn_SAP_UI_SAPGuiTable_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Value [ " & Cstr(sValue) & " ] found in table")
			End If		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL: invalid case.")		
	End Select
	
	If Err.Number <> 0 Then	
		Fn_SAP_UI_SAPGuiTable_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Clear memory of SAPGuiTable object.
	Set objSAPGuiTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_SAP_UI_SAPGuiTabStrip_Operations
'
'Function Description	 :	Function used to perform operations on SAPGuiTabStrip component
'
'Function Parameters	 :  1. sFunctionName	: Caller function's name
'							2. sAction			: Action to be performed
'							3. objContainer 	: Parent UI Component or JavaCheckBox object
'							4. sTabObjectName 	: SAPGuiTabStrip Control name in OR
'							5. sItem			: Tab Text to be selectd or verified
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage			 :	bReturn = Fn_SAP_UI_SAPGuiTabStrip_Operations("", "Select", objContainer, "AuthorizationTab",  "EW11397")
'
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale 				|  03-Feb-2016	    |	 1.0		|	Sandeep Navghane  	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_SAP_UI_SAPGuiTabStrip_Operations(sFunctionName, sAction, objContainer, sTabObjectName, sItem)
	'Variables declaration
	Dim objTab
	
	'Object Creation
	Fn_SAP_UI_SAPGuiTabStrip_Operations = False
	
	If sTabObjectName <> "" Then
		Set objTab = objContainer.SAPGuiTabStrip(sTabObjectName)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_SAP_UI_SAPGuiTabStrip_Operations  : [ " &  objContainer.toString & " ] : [ " &  sTabObjectName & " ] : Action = " & sAction & " : "
	Else
		Set objTab = objContainer
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_SAP_UI_SAPGuiTabStrip_Operations  : [ " &  objTab.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify SAPGuiTabStrip object exists
	If Fn_SAP_UI_SAPObject_Operations("Fn_SAP_UI_SAPGuiTabStrip_Operations","Exist", objTab, GBL_MIN_TIMEOUT, "", "") = False Then	
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " : FAIL : [ " &  objTab.toString & " ] object does not exist.")
		Set objTab = Nothing 
		Exit Function
	End If
	
	Select case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "select"
			objTab.Select sItem
			If Err.Number <> 0 Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : Failed to select SAP Gui Tab Strip item [ " & Cstr(sItem) & " ] as item does not exist")
			Else
				Fn_SAP_UI_SAPGuiTabStrip_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully select SAP Gui Tab Strip item [ " & Cstr(sItem) & " ]")
			End If		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " FAIL: invalid case.")		
	End Select
	
	If Err.Number <> 0 Then	
		Fn_SAP_UI_SAPGuiTabStrip_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ " & GBL_FUNCTIONLOG & "] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Clear memory of SAPGuiTabStrip object.
	Set objTab = Nothing 
End Function
