Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Function Name								|     Developer						|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - - -	| - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. Fn_WEB_UI_WebButton_Operations					|	sandeep.navghane@sqs.com		|	09-May-2016	|	Function used to perform operations on WebButton object
'002. Fn_WEB_UI_WebEdit_Operations						|	sandeep.navghane@sqs.com		|	09-May-2016	|	Function used to perform operations on WebEditBox object
'003. Fn_WEB_UI_WebObject_Operations					|	sandeep.navghane@sqs.com		|	09-May-2016	|	Function used to perform operations on Webobjects
'004. Fn_WEB_UI_WebTable_Operations						|	sandeep.navghane@sqs.com		|	09-May-2016	|	Function used to perform operations on WebTable object
'005. Fn_Web_UI_Link_Operations							|	sandeep.navghane@sqs.com		|	09-May-2016	|	Function used to perform operations on WebLink object
'006. Fn_Web_UI_WebElement_Operations					|	sandeep.navghane@sqs.com		|	09-May-2016	|	Function used to perform operations on WebElement object
'007. Fn_Web_UI_WebList_Operations						|	sandeep.navghane@sqs.com		|	09-May-2016	|	Function used to perform operations on WebList component
'008. Fn_Web_UI_Image_Operations						
'009. Fn_Web_UI_ObjectProperties_Operations				|	sandeep.navghane@sqs.com		|	09-May-2016	|	Function used to perform operations on Web Image
'011. Fn_WEB_UI_WebObject_OperationsExt					|	sandeep.navghane@sqs.com		|	21-Sep-2017	|	Function used to perform operations on Webobjects
'						
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WEB_UI_WebButton_Operations
'
'Function Description	:	Function used to perform operations on WebButton object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWebDialog: Object of Container  in which WebButton is present
'							4.sWebButtonName: Repository Name of WebButton
'							5.iX: X cordinate  Value
'							6.iY: Y cordinate Value
'							7.sMicButton: MouseButton Click Name
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WEB_UI_WebButtonOperations("Fn_Testing", "click", Page("Login"),"Login","","","")
'Function Usage		     :	Call Fn_WEB_UI_WebButtonOperations("Fn_Testing", "click", Page("Login").WebButton("Login"),"","1","1","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane			|  09-May-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WEB_UI_WebButton_Operations(sFunctionName, sAction, objWebDialog, sWebButtonName,iX,iY,sMicButton)
	'Declaring Variables	
	Dim objWebButton,bResult
	'Initially set function return value as False
	Fn_WEB_UI_WebButton_Operations=False
	'Object Creation
	If sWebButtonName<>"" Then
		Set objWebButton = objWebDialog.WebButton(sWebButtonName)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_WEB_UI_WebButton_Operations  : [ " & objWebDialog.toString & " ] : [ " &  objWebButton.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWebButton = objWebDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_WEB_UI_WebButton_Operations  : [ " & objWebButton.toString & " ] : Action = " & sAction & " : "
	End If
	
	'Verify WebButton object exists
	If  Fn_WEB_UI_WebObject_Operations("Fn_WEB_UI_WebButton_Operations", "Enabled", objWebButton , "","","") = False Then
		'Report error/message when WebEditBox object is disable.
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & " : FAIL : [ " & objWebButton.toString & " ] WebButton does not exist")
		Set objWebButton=Nothing
		Exit Function
	End If

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Click"
			If  iX <> "" AND iY <> "" Then
				'Click the mouse button at X,Y Co-ordinates
				objWebButton.Click iX, iY, sMicButton
			Else
				objWebButton.Click
			End If
			Fn_WEB_UI_WebButton_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Successfully clicked WebButton [ "& objWebButton.toString &" ]")
	End Select
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WEB_UI_WebButton_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & "] : Fail To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Release objects	
	Set objWebButton = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WEB_UI_WebEdit_Operations
'
'Function Description	:	Function used to perform operations on WebEditBox object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'							2.sAction: Action Name
'							3.objWebDialog: Object of Container  in which WebEditBox is present
'							4.sWebEdit: Repository name of WebEditBox
'							5.sValue: Text String Value
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WEB_UI_WebEdit_Operations("Fn_Testing","Set",objWebDialog, "Name", "SQS_Part_Name" )
'Function Usage		     :	Call Fn_WEB_UI_WebEdit_Operations("Fn_Testing","SetSecure",objWebDialog.WebEdit("Password"),"","Testing" )
'Function Usage		     :	Call Fn_WEB_UI_WebEdit_Operations("Fn_Testing","SendString",objWebDialog,"Name","SQS_Part_Name" )
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane			|  09-May-2016	    |	 1.0		|		Ganesh Bhosale  | 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Web_UI_WebEdit_Operations(sFunctionName, sAction, objWebDialog, sWebEdit, sValue)
	'Declaring Variables
	Dim objWebEdit,objDeviceReplay

	'Initially set function return value as False
	Fn_Web_UI_WebEdit_Operations=False
	
	If sWebEdit<>"" Then
		Set objWebEdit = objWebDialog.WebEdit(sWebEdit)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_WebEdit_Operations  : [ " &  objWebDialog.toString & " ] : [ " & objWebEdit.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWebEdit = objWebDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_WebEdit_Operations  : [ " & objWebEdit.toString & " ] : Action = " & sAction & " : "
	End If
	
	'Verify WebEditBox object exists
	If  Fn_WEB_UI_WebObject_Operations("Fn_Web_UI_WebEdit_Operations","Exist",objWebEdit ,"","","")= False Then
		'Report error/message when WebEditBox object is not exist
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & " FAIL : [ " & objWebEdit.toString & " ] WebEditBox does not exist")
		Set objWebEdit=Nothing
		Exit Function
	End If
		
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set the value in WebEditBox
		Case "Set","SetSecure"
			'Verify WebEditBox object is enabled
			If Fn_WEB_UI_WebObject_Operations("Fn_WEB_UI_WebEdit_Operations", "Enabled", objWebEdit,"","","") = False Then
				Fn_WEB_UI_WebEdit_Operations= False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : [ " & objWebEdit.toString & " ] WebEditBox is not enabled")
				'Release Object
				Set objWebEdit = Nothing 
				Exit Function
			End If
			If sAction="Set" Then
				objWebEdit.Set sValue
			ElseIf sAction="SetSecure" Then
				objWebEdit.SetSecure sValue
			End If
			Fn_Web_UI_WebEdit_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Successfully set the value [ " & Cstr(sValue) & " ] in WebEditBox [ " & objWebEdit.toString & " ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set the value in WebEditBox using send string method
        Case "SendString"
	        Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objWebEdit.Set ""
			objWebEdit.Object.focus
			objDeviceReplay.SendString sValue
			Wait 0, 300
			objDeviceReplay.PressKey 28
			Fn_Web_UI_WebEdit_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Successfully type the value [ " & Cstr(sValue) & " ] in WebEditBox [ " & objWebEdit.toString & " ]")
			Set objDeviceReplay =Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case "Get"
			Fn_Web_UI_WebEdit_Operations = objWebEdit.GetROProperty("value") 			
		'Case to set the value in WebEditBox
		Case "Click"
			'Verify WebEditBox object is enabled
			If Fn_WEB_UI_WebObject_Operations("Fn_WEB_UI_WebEdit_Operations", "Enabled", objWebEdit,"","","") = False Then
				Fn_WEB_UI_WebEdit_Operations= False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : [ " & objWebEdit.toString & " ] WebEditBox is not enabled")
				'Release Object
				Set objWebEdit = Nothing 
				Exit Function
			End If
				objWebEdit.Click
				Fn_Web_UI_WebEdit_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Successfully Clicked WebEditBox [ " & objWebEdit.toString & " ]")
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : [No valid case] : No valid case was passed for function [Fn_WEB_UI_WebEdit_Operations]")
	End Select
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WEB_UI_WebEdit_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & " ] : Fail To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release objects
	Set objWebEdit = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WEB_UI_WebObject_Operations
'
'Function Description	:	Function used to perform operations on Webobjects.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objReferencePath: Reference path of Object
'							4.iTimeOut: Time out time in seconds
'							5.sPropertyName: Valid Property Name
'							6.sPropertyValue: Property value
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WEB_UI_WebObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","create", objDateControl,"","","")
'Function Usage		     :	Call Fn_WEB_UI_WebObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","enabled", objDateControl, MAX_TIMEOUT,"","")
'Function Usage		     :	Call Fn_WEB_UI_WebObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","exist", objDateControl,"","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane			    |  09-May-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WEB_UI_WebObject_Operations(sFunctionName, sAction, objReferencePath, iTimeOut, sPropertyName, sPropertyValue)
	Err.Clear
	'Initially set function return value as False
	Fn_WEB_UI_WebObject_Operations = False
	
	GBL_FUNCTIONLOG = sFunctionName & " > Fn_WEB_UI_WebObject_Operations : [ " &  objReferencePath.toString & " ] : Action = " & sAction & " : "
	
	'Set the default time out
	If iTimeOut = "" Then
		iTimeOut = GBL_DEFAULT_TIMEOUT
	Else
		iTimeOut = cInt(iTimeOut)
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to create a Webobject
		Case "create"
			'Verify if the reference path of object is enabled
			If Fn_WEB_UI_WebObject_Operations(sFunctionName, "Enabled", objReferencePath, iTimeOut,"","") Then
				'Returning Webobject
				Set Fn_WEB_UI_WebObject_Operations = objReferencePath
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Successfully created new object for [ " & objReferencePath.ToString & " ]")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : WebObject is not enabled")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify existance of Webobject
		Case "exist"
			Fn_WEB_UI_WebObject_Operations = objReferencePath.Exist(iTimeOut)
			If Fn_WEB_UI_WebObject_Operations Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Successfully verified object [ "& objReferencePath.ToString &" ] exist")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : WebObject does not exist")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify Webobject is enabled
		Case "enabled"
			'Verify if the reference path of object is exist
			If Fn_WEB_UI_WebObject_Operations(sFunctionName, "Exist", objReferencePath, iTimeOut,"","") Then
				If objReferencePath.CheckProperty("Disabled",0) Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Successfully verified object [ " & objReferencePath.ToString & " ] exist and enabled")
					Fn_WEB_UI_WebObject_Operations = True
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : WebObject is not enabled")
				End If
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : WebObject does not exist")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the property value
		Case "getroproperty"
				Fn_WEB_UI_WebObject_Operations = objReferencePath.GetROProperty(sPropertyName)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Sucessfully fetched  property value [ " & sPropertyName & " ] of object [ "& objReferencePath.ToString &" ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set the property value
		Case "settoproperty"
				objReferencePath.SetTOProperty sPropertyName,sPropertyValue
				Fn_WEB_UI_WebObject_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Sucessfully set  property [ " & sPropertyName & " ] of object [ "& objReferencePath.ToString &" ] with value [ " & sPropertyName & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : [No valid case] : No valid case was passed for function [Fn_WEB_UI_WebObject_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WEB_UI_WebObject_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & " ] : Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WEB_UI_WebTable_Operations
'
'Function Description	:	Function used to perform operations on WebTable object.
'
'Function Parameters	:  	1. sFunctionName: Name of function
'							2. sAction: Action name
'							3. objWebDialog : Object of Container
'							4. sWebTableName : table name
'							5. sColName : Column name
'							6. iColumnRowNumber : Column row index
'							7. iColumnIndex : Column index
'							8. sNodeName : Table node name
'							9. sObjType :  object type
'							10.iIndex : Item index
'							11.sMethod : method name to perform
'							12.sValue : Value for the method
'							13.iX : x co-ordinate
'							14.iY : y co-ordinate
'							15.sMicButton : Button name
'
'Function Return Value	 : 	True or False Or Child Item object Or Column position Or Row position
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_Web_UI_WebTable_Operations("","getcolumnindex", objContentsPanel, "ContentsTableHeader", "OBJECT","","", "", "", "","","","","","")
'Function Usage		     :	Call Fn_Web_UI_WebTable_Operations("","executeobject", objContentsPanel, "ContentsTable", "","",2, "39.19.10.160 Configure Revision Rules_35309", "WebElement", "","Click","","","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane			    |  09-May-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_Web_UI_WebTable_Operations(sFunctionName,sAction,objWebDialog,sWebTableName,sColName,iColumnRowNumber,iColumnIndex,sNodeName,sObjType,iIndex,sMethod,sValue,iX,iY,sMicButton)
	'Declaring Variables
	Dim objWebTable,objChildItem
	Dim iColumnCount,iCounter,iColumnNumber,iRowNumber	
	Dim bFlag
	Dim sNode
	
	'Initially set function return value as False
	Fn_Web_UI_WebTable_Operations = False
	
	If sWebTableName<>"" Then
		Set objWebTable = objWebDialog.WebTable(sWebTableName)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_WebTable_Operations  : [ " &  objWebDialog.toString & " ] : [ " &  objWebTable.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWebTable = objWebDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_WebTable_Operations  : [ " &  objWebTable.toString & " ] : Action = " & sAction & " : "
	End If
	
	'Verify WebTable object exist
	If Fn_WEB_UI_WebObject_Operations("Fn_WEB_UI_WebTable_Operations", "Exist", objWebTable,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : [ " &  objWebTable.toString & " ] WebTable Object does not exist")
		'Release Object
		Set objWebTable = Nothing 
		Exit Function
	End If
	
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getcolumnindex"
			Fn_Web_UI_WebTable_Operations = -1
			If iColumnRowNumber="" Then
				iColumnRowNumber=1
			End If
			iColumnCount = objWebTable.ColumnCount(iColumnRowNumber)
			For iCounter = 1 to iColumnCount
				If  Trim(lcase(sColName)) = trim(lcase(objWebTable.GetCellData(iColumnRowNumber,iCounter))) Then
					Fn_Web_UI_WebTable_Operations = iCounter
					Exit for
				End If
			Next
			If Fn_Web_UI_WebTable_Operations = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : Column [ " & sColName & " ] is not present in specified table")
			else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Column [ " & sColName & " ] is present at index [ " & Cstr(Fn_Web_UI_WebTable_Operations) & " ] in specified table")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getrownumber"
			Fn_Web_UI_WebTable_Operations = -1
			bFlag = False
			If iColumnIndex="" Then
				iColumnNumber= Fn_Web_UI_WebTable_Operations(sFunctionName,"getcolumnindex", objWebTable, "", sColName,iColumnRowNumber,"", "","","","","","","","")
			Else
				iColumnNumber=iColumnIndex
			End If

			bFlag=False			
			For iCounter = 1 To objWebTable.RowCount 
				If Trim(objWebTable.GetCellData(iCounter, iColumnNumber)) = Trim(sNodeName) Then
					bFlag=True
					Exit For
				End If
			Next
			If bFlag Then
				Fn_Web_UI_WebTable_Operations = iCounter
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Node [ " & sNodeName & " ] is present at index [ " & Fn_Web_UI_WebTable_Operations & " ] in specified table.")
			else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : Node [ " & sNodeName & " ] is not present in specified table.")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getchildobjects"
			If iColumnIndex="" Then
				iColumnNumber = Fn_Web_UI_WebTable_Operations(sFunctionName,"getcolumnindex", objWebTable, "", sColName,iColumnRowNumber,"", "","","","","","","","")
			Else
				iColumnNumber =iColumnIndex
			End If			
			iRowNumber = Fn_Web_UI_WebTable_Operations(sFunctionName,"getrownumber", objWebTable, "", "", "",iColumnNumber,sNodeName,"","","","","","","")
			If iColumnNumber = -1 OR iRowNumber = -1 Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : Node [ " & sNodeName & " ] is not present in specified table.")
			End If
			If iIndex = "" Then
				iIndex = 0
			End If
			Set objChildItem = objWebTable.ChildItem(iRowNumber,iColumnNumber,sObjType,iIndex)
			If Typename(objChildItem)<>"Nothing" Then
				Set Fn_Web_UI_WebTable_Operations=objChildItem
			End If
			Set objChildItem = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "executeobject"
			If iColumnIndex="" Then
				Set objChildItem = Fn_Web_UI_WebTable_Operations(sFunctionName,"getchildobjects", objWebTable, "", sColName,iColumnRowNumber,"", sNodeName,sObjType,iIndex,"","","","","")
			Else
				Set objChildItem = Fn_Web_UI_WebTable_Operations(sFunctionName,"getchildobjects", objWebTable, "", "","",iColumnIndex,sNodeName,sObjType,iIndex,"","","","","")
			End If
			
			If Typename(objChildItem)="Nothing" Then
				Set objChildItem =Nothing
				Exit function
			End If
			If sMethod ="" Then
				sMethod="Click"
			End If
			Select Case lCase(sMethod)
				Case "click"
					If  iX <> "" AND iY <> "" Then
						'Click the mouse button at X,Y Co-ordinates
						objChildItem.Click iX, iY, sMicButton
					Else
						'Click
						objChildItem.Click
					End If
				Case "set"
					objChildItem.Set sValue
			End Select
			Set objChildItem = Nothing
			Fn_Web_UI_WebTable_Operations=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getcelldata"
			Fn_Web_UI_WebTable_Operations=Trim(objWebTable.GetCellData(sNodeName, iColumnIndex))
	End Select
	' log for error
	If Err.Number <> 0 Then	
		Fn_Web_UI_WebTable_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & " ] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If	
	'Release objects
	Set objWebTable = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_Web_UI_Link_Operations
'
'Function Description	:	Function used to perform operations on Web link object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'							2.sAction: Action Name
'							3.objWebDialog: Object of Container in which link is present
'							4.sLinkName: Valid Link Name
'							5.iX: Valid X Co-ordiate value
'							6.iY: Valid Y Co-ordiate value
'							7.sMicButtonToClick: Valid Mouse button to be Clicked
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_Web_UI_Link_Operations("","Click",Browser("TeamcenterWeb"),"Logout","","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane			|  09-May-2016	    |	 1.0		|		Ganesh Bhosale  | 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_Web_UI_Link_Operations(sFunctionName,sAction,objWebDialog,sLinkName,iX,iY,sMicButton)
	'Declaring Variables
	Dim objLink
	'Initially set function return value as False
	Fn_Web_UI_Link_Operations=False
	
	'Object Creation
	If sLinkName<>"" Then
		Set objLink = objWebDialog.Link(sLinkName)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_Link_Operations  : [ " &  objWebDialog.toString & " ] : [ " &  objLink.toString & " ] : Action = " & sAction & " : "
	Else
		Set objLink = objWebDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_Link_Operations  : [ " &  objLink.toString & " ] : Action = " & sAction & " : "
	End If
	
	'Verifying link existance
	If  Fn_WEB_UI_WebObject_Operations("Fn_Web_UI_Link_Operations", "Exist", objLink , "","","") = False Then
		'Report error/message when WebEditBox object is disable.
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : [ " &  objLink.toString & " ] link is not enabled")
		Set objLink = Nothing
		Exit Function
	End If
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click on link
		Case "Click"
            If  iX <> "" AND iY <> "" Then
				'Click the mouse button at X,Y Co-ordinates
				objLink.Click iX, iY, sMicButton
				Fn_Web_UI_Link_Operations = True
				'log on success
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully clicked on Link [ " &  objLink.toString & " ] at Co-ordinates [ " & Cstr(iX) & "," & Cstr(iY) & " ]")
			Else
				objLink.Click
				Fn_Web_UI_Link_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & " PASS : Sucessfully Clicked on Link [ " &  objLink.toString & " ]")
			End If
	End Select
	
	'log for error
	If Err.Number <> 0 Then	
		Fn_Web_UI_Link_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & " ] : Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Release objects	
	Set objLink = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_Web_UI_WebElement_Operations
'
'Function Description	:	Function used to perform operations on WebElement object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'							2.sAction: Action Name
'							3.objWebDialog: Object of Container in which webelement is present
'							4.sLinkName: Valid Link Name
'							5.iX: Valid X Co-ordiate value
'							6.iY: Valid Y Co-ordiate value
'							7.sMicButtonToClick: Valid Mouse button to be Clicked
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_Web_UI_WebElement_Operations("","Click",Browser("TeamcenterIssueManager").Page("IssueManager"),"CommonAttributes","","","")
'Function Usage		     :	Call Fn_Web_UI_WebElement_Operations("","RightClick",Browser("TeamcenterIssueManager").Page("IssueManager"),"CommonAttributes","","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane			|  09-May-2016	    |	 1.0		|		Ganesh Bhosale  | 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_Web_UI_WebElement_Operations(sFunctionName,sAction,objWebDialog,sWebElement,iX,iY,sMicButton)
	'Declaring Variables	
	Dim objWebElement
   	'Initially set function return value as False
	Fn_Web_UI_WebElement_Operations=False
	'Object Creation
	If sWebElement<>"" Then
		Set objWebElement = objWebDialog.WebElement(sWebElement)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_WebElement_Operations  : [ " &  objWebDialog.toString & " ] : [ " &  objWebElement.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWebElement = objWebDialog
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_WebElement_Operations  : [ " &  objWebElement.toString & " ] : Action = " & sAction & " : "
	End If
	
	'Verify WebElement object exists	
	If  Fn_WEB_UI_WebObject_Operations("Fn_Web_UI_WebElement_Operations", "Exist", objWebElement , "","","") = False Then
		'Report error/message when WebEditBox object is disable.
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "FAIL : [ " &  objWebElement.toString & " ] web element is not enabled")
		Set objWebElement = Nothing
		Exit Function
	End If
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click on web element
		Case "Click"
            If  iX <> "" AND iY <> "" Then
				'Click the mouse button at X,Y Co-ordinates
				objWebElement.Click iX, iY,sMicButton
				Fn_Web_UI_WebElement_Operations = True
				'log on success
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully clicked on WebElement [ " &  objWebElement.toString & " ] at Co-ordinates [ " & Cstr(iX) & "," & Cstr(iY) & " ]")
			Else
				objWebElement.Click
				Fn_Web_UI_WebElement_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully Clicked on WebElement [ " &  objWebElement.toString & " ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to right click on web element
		Case "RightClick"
            If  iX <> "" AND iY <> "" Then
				'Click the mouse button at X,Y Co-ordinates
				objWebElement.RightClick iX, iY
				Fn_Web_UI_WebElement_Operations = True
				'log on success
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully right clicked on WebElement [ " &  objWebElement.toString & " ] at Co-ordinates [ " & Cstr(iX) & "," & Cstr(iY) & " ]")
			Else
				objWebElement.RightClick
				Fn_Web_UI_WebElement_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), GBL_FUNCTIONLOG & "PASS : Sucessfully right Clicked on WebElement [ " &  objWebElement.toString & " ]")
			End If
	End Select
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_Web_UI_WebElement_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & "] : Fail To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	'Release objects
	Set objWebElement = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_Web_UI_WebList_Operations
'
'Function Description	 :	This function is used to perform operations on WebList component.
'
'Function Parameters	 :  1. sFunctionName	: Caller function's name
'						    2. sAction			: Action to be performend
'						    3. objContainer 	: Parent UI Component or Weblist	
'				    		4. sWeblist			: Valid Weblist name on which operation is to be done
'							5. sElementToSelect	: Valid value that has to be selected from the List
'
'Function Return Value	 : 	True \ False \ ListItem value
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  bReturn = Fn_Web_UI_WebList_Operations("Fn_Web_PasteAs", "Select", Browser("Teamcenter").Page("MyTeamcenter").WebTable("NewFolder").WebList("Type"), "","Folder")
'Function Usage		     :  bReturn = Fn_Web_UI_WebList_Operations("Fn_Web_PasteAs", "verifycurrentselectedvalue", Browser("Teamcenter").Page("MyTeamcenter").WebTable("NewFolder"), "Type","Folder")
'Function Usage		     :  bReturn = Fn_Web_UI_WebList_Operations("Fn_Web_PasteAs", "verify", Browser("Teamcenter").Page("MyTeamcenter").WebTable("NewFolder"), "Type","Folder")
'Function Usage		     :  bReturn = Fn_Web_UI_WebList_Operations("Fn_Web_PasteAs", "getcontents", Browser("Teamcenter").Page("MyTeamcenter").WebTable("NewFolder"), "Type","")
'Function Usage		     :  bReturn = Fn_Web_UI_WebList_Operations("Fn_Web_PasteAs", "verifycontents", Browser("Teamcenter").Page("MyTeamcenter").WebTable("NewFolder").WebList("Type"), "","Folder~Item")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Ganesh Bhosale		  		|  08-Feb-2016	    |	 1.0		|	Ganesh Bhosale 		| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Web_UI_WebList_Operations(sFunctionName, sAction, objContainer, sWeblist,sElementToSelect)
	' Variable declaration
	Dim objWeblist
	Dim bFlag
	Dim iCounter, iCount, iEelecount
	Dim sContents,sValues
	Dim aValues,aContents
	
	'Initialize Function value to False
	Fn_Web_UI_WebList_Operations=False
	
	'Object Creation
	If sWeblist<>"" Then
		Set objWeblist = objContainer.Weblist(sWeblist)
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_WebList_Operations  : [ " &  objContainer.toString & " ] : [ " &  objWeblist.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWeblist = objContainer
		GBL_FUNCTIONLOG = sFunctionName & " > Fn_Web_UI_WebList_Operations  : [ " &  objWeblist.toString & " ] : Action = " & sAction & " : "
	End If
	
	'Report error/message when WebEditBox object is disable.
	If  Fn_WEB_UI_WebObject_Operations("Fn_Web_UI_WebList_Operations", "Exist", objWeblist ,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "[ " & GBL_FUNCTIONLOG & " ] FAIL : WebList [ "& sWeblist.toString &" ] is not enabled.")
		Set objWeblist=Nothing
		Exit Function
	End If

	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "select"
			'Select the element from List
			objWeblist.Select sElementToSelect
			'Report message of Selected the Element from Weblist succesfully.				
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "[ " & GBL_FUNCTIONLOG & " ] PASS : Sucessfully Selected element [ " & sElementToSelect &" ] under Web List [ " & sWeblist & " ] of Function [ " &sFunctionName & " ]")
			Fn_Web_UI_WebList_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "verifycurrentselectedvalue"
			' to verify current selected value
			If Trim(objWeblist.GetROProperty("value"))=Trim(sElementToSelect) Then
				'Report message of verify current selected Element from Weblist.				
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "[ " & GBL_FUNCTIONLOG & " ] : Sucessfully verified element " & sElementToSelect &" is currently selected value in Web List " & sWeblist & " of Function " & sFunctionName)
				Fn_Web_UI_WebList_Operations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "verify"
			'verify Element is exist in list or not
			For iCounter=1 to objWeblist.GetROProperty("items count")
				If Trim(objWeblist.GetItem(iCounter))=Trim(sElementToSelect) Then
					'Report message of Selected the Element from Weblist succesfully.				
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "[ " & GBL_FUNCTIONLOG & " ] : Sucessfully verified element " & sElementToSelect &" exist in Web List " & sWeblist & " of Function " & sFunctionName)
					Fn_Web_UI_WebList_Operations = True
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getcontents"
			'get Element count in list
			iEelecount = objWeblist.GetROProperty("items count")
			' fetch all the items from list
			For iCounter = 1 To iEelecount
				If iCounter = 1 Then
					Fn_Web_UI_WebList_Operations = Trim(objWeblist.GetItem(iCounter))
				Else
					Fn_Web_UI_WebList_Operations = Fn_Web_UI_WebList_Operations & "~" & Trim(cstr(objWeblist.GetItem(iCounter)))
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getitematindex"
		   'get element at specified index in variable sElementToSelect
		   Fn_Web_UI_WebList_Operations= objWeblist.GetItem(sElementToSelect)
		   'Report message of returning element at specified index succesfully
		   Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "[ " & GBL_FUNCTIONLOG & " ] PASS : Sucessfully retrieved element under Web List [ " & sWeblist & " ] at index [ " & sElementToSelect & " ]")
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        Case "verifycontents"
			' get contents of weblist
			sContents = Fn_Web_UI_WebList_Operations(sFunctionName, "GetContents", objWeblist,"","")
			If sContents = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"[ " & GBL_FUNCTIONLOG & " ] : Failed to get Content of list")
				Exit Function
			End If
			' compare value passed with contents of weblist
			aContents = Split(sContents,"~",-1,1)
			aValues = split(sElementToSelect,"~",-1,1)
			For iCount = 0 to Ubound(aValues)
				bFlag = False
				For iCounter = 0 to Ubound(aContents)
					If  aValues(iCount) = aContents(iCounter) Then
						bFlag = True
						Exit For
					End If
				Next
				If bFlag = False Then
					Fn_Web_UI_WebList_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"[ " & GBL_FUNCTIONLOG & " ] : Failed to verify that the value [ " & aValues(iCount) & " ] is in list")
					Exit For
				End If
			Next
			If bFlag = True Then
				Fn_Web_UI_WebList_Operations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "[ " & GBL_FUNCTIONLOG & " ] FAIL : Invalid Case.")
	End Select
	
	'log for error
	If Err.Number <> 0 Then	
		Fn_Web_UI_WebList_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & " ] Fail to perform [ " & sAction & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	' Release Object memory
	Set objWeblist=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_Web_UI_Image_Operations
'
'Function Description	:	Function used to perform operations on Web Image.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWebPage: Valid Web Page path
'							4.sWebImage: Valid image name
'							5.iXValue: X cordinate  Value
'							6.iYValue: Y cordinate Value
'							7.sMicButton: MouseButton Click Name
'							8.sEvent: Event name to be fired
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Teamcenter web application should be displayed
'
'Function Usage		     :	Call Fn_Web_UI_Image_Operations("","fireevent",Browser("TeamcenterIssueManager").Page("IssueManager").WebTable("ManuTable"),"MyIssues","","","","onmouseover")
'Function Usage		     :	Call Fn_Web_UI_Image_Operations("","click",Browser("TeamcenterIssueManager").Page("IssueManager"),"NavigateOTBTeamcenter","","","","")
'Function Usage		     :	Call Fn_Web_UI_Image_Operations("","click",Browser("TeamcenterIssueManager").Page("IssueManager"),"NavigateOTBTeamcenter","123","154","","")
'Function Usage		     :	Call Fn_Web_UI_Image_Operations("","doubleclick",Browser("TeamcenterIssueManager").Page("IssueManager").WebTable("ManuTable"),"MyIssues","","","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  09-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Web_UI_Image_Operations(sFunctionName, sAction, objWebPage, sWebImage, iXValue, iYValue, sMicButton, sEvent)
	'Declaring variables	
	Dim objWebImage, objDeviceReplay 
	
	'Initially set function return value as False
	Fn_Web_UI_Image_Operations = False
	
	'Creating/Setting Object of WebImage
	If sWebImage <> "" Then
		Set objWebImage = objWebPage.Image(sWebImage)
		GBL_FUNCTIONLOG = " [ " &  objWebPage.toString & " ] : [ " &  objWebImage.toString & " ] : [ Action = " & sAction & " ] : "
	Else
		Set objWebImage = objWebPage
		GBL_FUNCTIONLOG = " [ " +  objWebImage.toString + " ] : [ Action = " & sAction & " ] : "
	End If
	
	'Verify WebImage object exists
	If Fn_WEB_UI_WebObject_Operations("Fn_Web_UI_Image_Operations", "Exist", objWebImage ,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_Web_UI_Image_Operations >> FAIL : "& GBL_FUNCTIONLOG&"WebImage does not exist")
		'Release object
		Set objWebImage = Nothing 
		Exit Function
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to fire an event on the WebImage
		Case "fireevent"
			objWebImage.Object.focus
			objWebImage.FireEvent sEvent
			Fn_Web_UI_Image_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_Web_UI_Image_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully fired event [ "& sEvent &" ] on WebImage [ "& objWebImage.toString&" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click on the WebImage
		Case "click"
            If  iXValue <> "" AND iYValue <> "" Then
				'Click the mouse button at X,Y Co-ordinates
				objWebImage.Click iXValue, iYValue, sMicButton
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_Web_UI_Image_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully clicked at co-ordinates X : [ "& iXValue &" ] and Y : [ "& iYValue &" ] on WebImage [ "& objWebImage.toString &" ]")
			Else
				objWebImage.Click
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_Web_UI_Image_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully clicked on WebImage [ "& objWebImage.toString &" ]")
			End If
			Fn_Web_UI_Image_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to double click on the WebImage
		Case "doubleclick"
			iXValue = objWebImage.getROProperty("abs_x")
			iYValue = objWebImage.getROProperty("abs_y")
            If  iXValue <> "" AND iYValue <> "" Then
				'Creating object of DeviceReplay
				Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
				objDeviceReplay.MouseMove iXValue, iYValue
				wait GBL_MIN_TIMEOUT
				objDeviceReplay.MouseDblClick iXValue, iYValue,0
				Fn_Web_UI_Image_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_Web_UI_Image_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully double clicked at co-ordinates X : [ "& iXValue &" ] and Y : [ "& iYValue &" ] on WebImage [ "& objWebImage.toString &" ]")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_Web_UI_Image_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Co-ordinates of WebImage is not get")
			End If
			'Release Object
			Set objDeviceReplay = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_Web_UI_Image_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_Web_UI_Image_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_Web_UI_Image_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_Web_UI_Image_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If	
	'Release objects
	Set objWebImage = Nothing 
End Function

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Function Name		:	Fn_Web_UI_ObjectProperties_Operations

'Description		:	Function to perform operations on Web object properties.

'Parameters			:	1. sFunctionName
'						2. sAction
'						3. sReferencePath : Valid Reference Path
'						4. sProperty    : Valid property Name
'						5. sPropertyValue	  : Property value

'Return Value		: 	TRUE \ FALSE

'Pre-requisite		:	Teamcenter web application should be displayed

'Examples			:	Call Fn_Web_UI_ObjectProperties_Operations("","SetTOExistCheck",Browser("TeamcenterIssueManager").Page("IssueManager").WebElement("ObjectTab"),"innertext","Common Attributes")
'						
'History			:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name						Date					Rev. No.			Reviewer					Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep Navghane		 	 22-Oct-2013			1.0						Sandeep
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Function Fn_Web_UI_ObjectProperties_Operations(sFunctionName,sAction,sReferencePath,sProperty,sPropertyValue)
	Dim objDialog, sFuncLog,bResult
	Fn_Web_UI_ObjectProperties_Operations = False
	'Object Creation
	If Not IsEmpty(sReferencePath) Then
		Set objDialog = sReferencePath
		sFuncLog = sFunctionName & " > Fn_Web_UI_ObjectProperties_Operations > " & objDialog.toString() & " : "
	End If
		
	Select Case (sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "SetTOExistCheck"
				objDialog.SetTOProperty sProperty,sPropertyValue
				bResult= Fn_WEB_UI_WebObject_Operations("Fn_Web_UI_ObjectProperties_Operations", "Exist", objDialog ,"","","")
				If  bResult = True Then
					Fn_Web_UI_ObjectProperties_Operations = True
					'log on success
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully Set TO property")
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & "FAIL : Fail to Set TO property")
				End If
		Case "GetROProperty"
				Fn_Web_UI_ObjectProperties_Operations=objDialog.GetROProperty(sProperty)
	End Select
	Set objDialog = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WEB_UI_WebObject_OperationsExt
'
'Function Description	:	Function used to perform operations on Webobjects.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objReferencePath: Reference path of Object
'							4.iTimeOut: Time out time in seconds
'							5.sPropertyName: Valid Property Name
'							6.sPropertyValue: Property value
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WEB_UI_WebObject_OperationsExt("Fn_SISW_RAC_UI_DateControl_SetDate","exist", objDateControl,"","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane			    |  09-May-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WEB_UI_WebObject_OperationsExt(sFunctionName, sAction, objReferencePath, iTimeOut, sPropertyName, sPropertyValue)
	Err.Clear
	Dim iCounter
	'Initially set function return value as False
	Fn_WEB_UI_WebObject_OperationsExt = False
	
	GBL_FUNCTIONLOG = sFunctionName & " > Fn_WEB_UI_WebObject_OperationsExt : [ " &  objReferencePath.toString & " ] : Action = " & sAction & " : "
	
	'Set the default time out
	If iTimeOut = "" Then
		iTimeOut = GBL_DEFAULT_TIMEOUT
	Else
		iTimeOut = cInt(iTimeOut)
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify existance of Webobject
		Case "exist"
			For iCounter=1 to Cint(iTimeOut)
				If objReferencePath.Exist(1) Then
					If objReferencePath.GetROProperty("height")>0 Then
						Fn_WEB_UI_WebObject_OperationsExt =True
						Exit For
					Else
						wait 1
					End if
				Else
					wait 1
				End If
			Next
			
			If Fn_WEB_UI_WebObject_OperationsExt Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "PASS : Successfully verified object [ "& objReferencePath.ToString &" ] exist")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : WebObject does not exist")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),GBL_FUNCTIONLOG & "FAIL : [No valid case] : No valid case was passed for function [Fn_WEB_UI_WebObject_OperationsExt]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WEB_UI_WebObject_OperationsExt = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ " & GBL_FUNCTIONLOG & " ] : Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
End Function
