Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Function Name								|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. Fn_WIN_UI_WinButtonOperations						|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinButton object
'002. Fn_WIN_UI_WinEdit_Operations						|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinEditBox object
'003. Fn_WIN_UI_WinObject_Operations					|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on Winobjects
'004. Fn_WIN_UI_WinList_Operations						|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinList object
'005. Fn_WIN_UI_WinRadioButton_Operations				|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinRadioButton object
'006. Fn_WIN_UI_WinComboBox_Operations					|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinComboBox object
'007. Fn_WIN_UI_WinCheckBox_Operations					|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinCheckBox object
'008. Fn_WIN_UI_WinEditor_Operations					|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinEditor object
'009. Fn_WIN_UI_WinTab_Operations						|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinTab object
'010. Fn_WIN_UI_WinTable_Operations						|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinTable object
'011. Fn_WIN_UI_WinToolbar_Operations					|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinToolbar object
'012. Fn_WIN_UI_WinTreeView_Operations					|	vrushali.sahare@sqs.com		|	03-Feb-2016	|	Function used to perform operations on WinTreeView object
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinButtonOperations
'
'Function Description	:	Function used to perform operations on WinButton object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWinDialog: Object of Container  in which WinButton is present
'							4.sWinButtonName: Repository Name of WinButton
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
'Function Usage		     :	Call Fn_WIN_UI_WinButtonOperations("Fn_CatiaLogin", "click", Window("Catia Login").WinButton("Login"),"","1","1","")
'Function Usage		     :	Call Fn_WIN_UI_WinButtonOperations("Fn_CatiaLogin", "devicereplay.click", Window("Catia Login"),"Login","","","")
'Function Usage		     :	Call Fn_WIN_UI_WinButtonOperations("Fn_CatiaLogin", "doubleclick", Window("Catia Login").WinButton("Login"),"","1","1","micLeftBtn")
'Function Usage		     :	Call Fn_WIN_UI_WinButtonOperations("Fn_CatiaLogin", "getvisibletext", Window("Catia Login").WinButton("Login"),"","","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinButtonOperations(sFunctionName, sAction, objWinDialog, sWinButtonName,iX ,iY, sMicButton)
	'Declaring variables
	Dim objWinButton, objDeviceReplay
	
	'Initially set function return value as False
	Fn_WIN_UI_WinButtonOperations = False
	
	'Creating/Setting Object of Win Button
	If sWinButtonName <> "" Then
		Set objWinButton = objWinDialog.WinButton(sWinButtonName)
		GBL_FUNCTIONLOG = " [ " & objWinDialog.toString & " ] : [ " &  objWinButton.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWinButton = objWinDialog
		GBL_FUNCTIONLOG = " [ " & objWinButton.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify WinButton object exists
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinButtonOperations", "Enabled", objWinButton,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinButtonOperations >> FAIL : "& GBL_FUNCTIONLOG &"WinButton does not exist")
		'Release object
		Set objWinButton = Nothing 
		Exit Function
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click on the win button
		Case "click"
			If iX <> "" and iY <> "" Then
				If sMicButton <> "" Then
					objWinButton.Click iX,iY,sMicButton
				Else
					objWinButton.Click iX,iY
				End If
			Else
				objWinButton.Click
			End If	
			Fn_WIN_UI_WinButtonOperations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinButtonOperations >> PASS : "& GBL_FUNCTIONLOG &"Successfully clicked WinButton [ "& objWinButton.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click on the win button with device replay
		Case "devicereplay.click"
			If sWinButtonName <> "" Then
				objWinDialog.Activate
				wait GBL_MICRO_TIMEOUT
			End If
			'Creating object of DeviceReplay
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.MouseMove (objWinButton.GetROProperty("abs_x") + 5), (objWinButton.GetROProperty("abs_y") + 5)
			objDeviceReplay.MouseClick  (objWinButton.GetROProperty("abs_x") + 5), (objWinButton.GetROProperty("abs_y") + 5), 0
			Fn_WIN_UI_WinButtonOperations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinButtonOperations >> PASS : "& GBL_FUNCTIONLOG &"Successfully performed device replay click on WinButton [ "& objWinButton.toString &" ]")
			'Release Object
			Set objDeviceReplay = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click on the win button with device replay
		Case "devicereplay.clickext"
			objWinButton.Highlight
			wait GBL_MIN_TIMEOUT
				
			'Creating object of DeviceReplay
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.MouseMove (objWinButton.GetROProperty("abs_x") + 5), (objWinButton.GetROProperty("abs_y") + 5)
			objDeviceReplay.MouseClick  (objWinButton.GetROProperty("abs_x") + 5), (objWinButton.GetROProperty("abs_y") + 5), 0
			Fn_WIN_UI_WinButtonOperations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinButtonOperations >> PASS : "& GBL_FUNCTIONLOG &"Successfully performed device replay click on WinButton [ "& objWinButton.toString &" ]")
			'Release Object
			Set objDeviceReplay = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to double click on WinButton
		Case "doubleclick"
			If sMicButton <> "" Then
				objWinButton.DblClick iX,iY,sMicButton
			Else
				objWinButton.DblClick iX,iY
			End If
			Fn_WIN_UI_WinButtonOperations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinButtonOperations >> PASS : "& GBL_FUNCTIONLOG &"Successfully double clicked on WinButton [ "& objWinButton.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'### This case is not tested.
		'Case to get the visible text of the win button 
		Case "getvisibletext"
			Fn_WIN_UI_WinButtonOperations = objWinButton.GetVisibleText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinButtonOperations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the visible text [ "& Fn_WIN_UI_WinButtonOperations &" ] of WinButton [ "& objWinButton.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinButtonOperations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinButtonOperations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinButtonOperations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinButtonOperations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release objects
	Set objWinButton = Nothing 
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinEdit_Operations
'
'Function Description	:	Function used to perform operations on WinEditBox object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'							2.sAction: Action Name
'							3.objWinDialog: Object of Container  in which WinEditBox is present
'							4.sWinEdit: Repository name of WinEditBox
'							5.sText: Text String Value
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WIN_UI_WinEdit_Operations("Fn_Catia_PartCreate", "set",  objWinDialog, "Name", "SQS_Part_Name" )
'Function Usage		     :	Call Fn_WIN_UI_WinEdit_Operations("Fn_Catia_PartCreate", "type",  objWinDialog.WinEdit("Name"),"", "SQS_Part_Name" )
'Function Usage		     :	Call Fn_WIN_UI_WinEdit_Operations("Fn_Catia_PartCreate", "gettext",  objWinDialog,"Name", "" )
'Function Usage		     :	Call Fn_WIN_UI_WinEdit_Operations("Fn_Catia_PartCreate", "sendstring",  objWinDialog,"Name", "SQS_Part_Name" )
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale  | 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinEdit_Operations(sFunctionName, sAction, objWinDialog, sWinEdit, sText)
	'Declaring Variables
	Dim objWinEdit
	
	'Initially set function return value as False
	Fn_WIN_UI_WinEdit_Operations = False
	
	'Creating/Setting Object of WinEditBox
	If sWinEdit <> "" Then
		Set objWinEdit= objWinDialog.WinEdit(sWinEdit)
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " & objWinEdit.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWinEdit= objWinDialog
		GBL_FUNCTIONLOG = " [ " & objWinEdit.toString & " ] : Action = " & sAction & " : "
	End If

	'Verify WinEditBox object exists
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinEdit_Operations", "Exist", objWinEdit,"","","") = False Then
		Fn_WIN_UI_WinEdit_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEdit_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinEditBox does not exist")
		'Release Object
		Set objWinEdit = Nothing 
		Exit Function
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set the value in WinEditBox
		Case "set","type"
			'Verify WinEditBox object is enabled
			If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinEdit_Operations", "Enabled", objWinEdit,"","","") = False Then
				Fn_WIN_UI_WinEdit_Operations= False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEdit_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinEditBox is not enabled")
				'Release Object
				Set objWinEdit = Nothing 
				Exit Function
			End If
			If Lcase(sAction) = "set" Then
				'Set the value if WinEditBox is enabled
				objWinEdit.Set sText
			Else
				'Type the value if WinEditBox is enabled
				objWinEdit.Type sText
			End If
			Fn_WIN_UI_WinEdit_Operations= True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEdit_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully " & Cstr(sAction) & " the value [ "& sText &" ] in WinEditBox [ " & objWinEdit.toString & " ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the text value from the WinEditBox
		Case "gettext"
			Fn_WIN_UI_WinEdit_Operations = objWinEdit.getROProperty("text")
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEdit_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the text value [ "& Fn_WIN_UI_WinEdit_Operations &" ] from the WinEditBox [ " & objWinEdit.toString & " ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to send the string value to the WinEditBox
		Case "sendstring"
			If sWinEdit <> "" Then
				objWinEdit.SetFocus
				wait GBL_MICRO_TIMEOUT
			End If 
			'Creating object of DeviceReplay
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			'Sending the string
			objDeviceReplay.SendString sText
			Fn_WIN_UI_WinEdit_Operations= True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEdit_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully send the string value [ "& sText &" ] to the WinEditBox [ " & objWinEdit.toString & " ]")
			'Release Object
			Set objDeviceReplay = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEdit_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinEdit_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinEdit_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEdit_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release objects
	Set objWinEdit = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinObject_Operations
'
'Function Description	:	Function used to perform operations on Winobjects.
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
'Function Usage		     :	Call Fn_WIN_UI_WinObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","create", objDateControl,"")
'Function Usage		     :	Call Fn_WIN_UI_WinObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","enabled", objDateControl, MAX_TIMEOUT)
'Function Usage		     :	Call Fn_WIN_UI_WinObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","exist", objDateControl,"")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinObject_Operations(sFunctionName, sAction, objReferencePath, iTimeOut, sPropertyName, sPropertyValue)
	Err.Clear
	'Initially set function return value as False
	Fn_WIN_UI_WinObject_Operations = False
	
	GBL_FUNCTIONLOG = " [ " &  objReferencePath.toString & " ] : Action = " & sAction & " : "
	
	'Set the default time out
	If iTimeOut = "" Then
		iTimeOut = GBL_DEFAULT_TIMEOUT
	Else
		iTimeOut = cInt(iTimeOut)
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to create a Winobject
		Case "create"
			'Verify if the reference path of object is enabled
			If Fn_WIN_UI_WinObject_Operations(sFunctionName, "Enabled", objReferencePath, iTimeOut) Then
				'Returning Winobject
				Set Fn_WIN_UI_WinObject_Operations = objReferencePath
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully created new object for [ "& objReferencePath.ToString &" ]")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinObject is not enabled")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify existance of Winobject
		Case "exist"
			Fn_WIN_UI_WinObject_Operations = objReferencePath.Exist(iTimeOut)
			If Fn_WIN_UI_WinObject_Operations Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully verified object [ "& objReferencePath.ToString &" ] exist")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinObject does not exist")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify Winobject is enabled
		Case "enabled"
			'Verify if the reference path of object is exist
			If Fn_WIN_UI_WinObject_Operations(sFunctionName, "Exist", objReferencePath, iTimeOut,"","") Then
				If objReferencePath.GetROProperty("enabled") = "1"  OR objReferencePath.GetROProperty("enabled") = True Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully verified object [ "& objReferencePath.ToString &" ] exist and enabled")
					Fn_WIN_UI_WinObject_Operations = True
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinObject is not enabled")
				End If
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinObject does not exist")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the property value
		Case "getroproperty"
				Fn_WIN_UI_WinObject_Operations = objReferencePath.GetROProperty(sPropertyName)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFunctionName&" >> Fn_WIN_UI_WinObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Sucessfully fetched  property value [ " & sPropertyName & " ] of object [ "& objReferencePath.ToString &" ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set the property value
		Case "settoproperty"
				objReferencePath.SetTOProperty sPropertyName,sPropertyValue
				Fn_WIN_UI_WinObject_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> PASS : "& GBL_FUNCTIONLOG &"Sucessfully set  property [ " & sPropertyName & " ] of object [ "& objReferencePath.ToString &" ] with value [ " & sPropertyName & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinObject_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinObject_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinObject_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinList_Operations
'
'Function Description	:	Function used to perform operations on WinList object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWinDialog: Object of Container  in which list is present
'							4.sWinList: Repository Name of WinList
'							5.sValues: Value to be selected / verified
'							6.sColumns: Column of value
'							7.sInstanceHandler: Instance Handler
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "select", Dialog("Teamcenter Login"),"ItemList","AutoTest", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "click", Dialog("Teamcenter Login"),"LoItemListgin","AutoTest", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "activate", Dialog("Teamcenter Login"),"ItemList","AutoTest", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "exist", Dialog("Teamcenter Login"),"ItemList","AutoTest", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "extendselect", Dialog("Teamcenter Login"),"ItemList","AutoTest", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "gettext", Dialog("Teamcenter Login"),"ItemList","", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "getcontents", Dialog("Teamcenter Login"),"ItemList","", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "verifycontents", Dialog("Teamcenter Login"),"ItemList","", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "type", Dialog("Teamcenter Login"),"ItemList","AutoTest", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinList_Operations("Fn_TeamcenterLogin", "getselection", Dialog("Teamcenter Login"),"ItemList","", "", "")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinList_Operations(sFunctionName, sAction, objWinDialog, sWinList, sValues, sColumns, sInstanceHandler)
	'Declaring Variables
	Dim objWinList
	Dim aSelectList, aContents, aValues
	Dim iCounter, iElementCount, iInstanceCount, iCount
	Dim sContents
	Dim bFlag

	'Initially set function return value as False
	Fn_WIN_UI_WinList_Operations = False
	
	'Set the instance handler
	If sInstanceHandler = "" Then
		sInstanceHandler = "@"
	End If
	
	'Creating/Setting Object of WinList
	If sWinList <> "" Then
		Set objWinList = objWinDialog.WinList(sWinList)
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " +  objWinList.toString + " ] : Action = " & sAction & " : "
	Else
		Set objWinList = objWinDialog
		GBL_FUNCTIONLOG = " [ " &  objWinList.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify WinList object exists
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinList_Operations", "Exist", objWinList,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinList Object does not exist")
		'Release Object
		Set objWinList = Nothing 
		Exit Function
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the contents of list
		Case "getcontents"
			iElementCount = objWinList.GetROProperty("items count")
			For iCounter = 0 To iElementCount - 1
				If iCounter = 0 Then
					Fn_WIN_UI_WinList_Operations = Trim(cstr(objWinList.GetItem(iCounter)))
				Else
					Fn_WIN_UI_WinList_Operations = Fn_WIN_UI_WinList_Operations & "~" & Trim(cstr(objWinList.GetItem(iCounter)))
				End If
			Next
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the contents [ "& Fn_WIN_UI_WinList_Operations &" ] of WinList [ "& objWinList.toString & " ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the text value of WinList
		Case "gettext"
			Fn_WIN_UI_WinList_Operations = objWinList.GetROProperty("value")
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the text value [ "& Fn_WIN_UI_WinList_Operations &" ] of WinList [ "& objWinList.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click on the WinList
		Case "click"
			objWinList.Click sValues, sColumns
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully clicked on WinList [ "& objWinList.toString &" ]")
			Fn_WIN_UI_WinList_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to activate the WinList value
		Case "activate"
			objWinList.Activate sValues
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully activated the value [ "& sValues &" ] on WinList [ "& objWinList.toString &" ]")
			Fn_WIN_UI_WinList_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify the WinList value exist
		Case "exist"
			'get total items from list
			iElementCount = objWinList.GetROProperty("items count")
			iInstanceCount = 1
			aSelectList = split(sValues, sInstanceHandler)
			If uBound(aSelectList) > 0 Then
				iInstanceCount = aSelectList(1)
			End If
			aSelectList(0) = trim(aSelectList(0))
			For iCounter = 0 To iElementCount - 1
				If objWinList.GetItem(iCounter) <> "" Then
					If Trim(cstr(objWinList.GetItem(iCounter))) = Trim(aSelectList(0)) Then
						If iInstanceCount = 1 Then
							Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully verified WinList value [ "& cstr(objWinList.GetItem(iCounter))&" ] exist")
							Fn_WIN_UI_WinList_Operations = True
							Exit For
						End If
						iInstanceCount = iInstanceCount - 1 
					End If
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to extend the WinList selection
		Case "extendselect"
			aSelectList=Split(sValues,"~")
			'Select the element from list  
			For iCounter = 0 To Ubound(aSelectList)
				objWinList.ExtendSelect aSelectList(iCounter)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected WinList value [ "& aSelectList(iCounter) &" ]")
			Next
			Fn_WIN_UI_WinList_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to type value in WinList
		Case "type"
			objWinList.Type sValues
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully type WinList value [ "& sValues &" ]")
			Fn_WIN_UI_WinList_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select WinList value
		Case "select"
			'get total items from list
			iElementCount = objWinList.GetROProperty("items count")
			iInstanceCount = 1
			aSelectList = split(sValues, sInstanceHandler)
			If uBound(aSelectList) > 0 Then
				iInstanceCount = aSelectList(1)
			End If
			aSelectList(0) = trim(aSelectList(0))
			For iCounter = 0 To iElementCount - 1
				If objWinList.GetItem(iCounter) <> "" Then
					If Trim(cstr(objWinList.GetItem(iCounter))) = Trim(aSelectList(0)) Then
						If iInstanceCount = 1 Then
							objWinList.Select iCounter
							Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected WinList value [ "& sValues &" ]")
							Fn_WIN_UI_WinList_Operations = True
							Exit For
						End If
						iInstanceCount = iInstanceCount - 1 
					End If
				End If
			Next
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify contents of WinList
        Case "verifycontents"
			sContents = Fn_WIN_UI_WinList_Operations("Fn_WIN_UI_WinList_Operations", "GetContents", objWinList,"","", "", "")
			If sContents = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Content of List is not get")
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
					Fn_WIN_UI_WinList_Operations = False
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Failed to verify the value [ "& aValues(iCount) & " ] exist in list")
					Exit For
				End If
			Next
			If bFlag = True Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully verify the value [ "& aValues(iCount) &" ] exist in list")
				Fn_WIN_UI_WinList_Operations = True
			End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'### This case is not tested.
	'Case to get the selected items in the WinList
	Case "getselection"
		Fn_WIN_UI_WinList_Operations = objWinList.GetSelection
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the selected item(s) [ "& Fn_WIN_UI_WinList_Operations &" ] of WinList [ "& objWinList.toString &" ]")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to handle invalid request
	Case Else
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinList_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinList_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinList_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release object
	Set objWinList = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinRadioButton_Operations
'
'Function Description	:	Function used to perform operations on WinRadioButton object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWinDialog: Object of Container  in which WinRadioButton is present
'							4.sWinRadioButtonName: Repository name of RadioButton
'							5.sValue: Value of Radio Button
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'							
'
'Function Usage		     :	Call Fn_WIN_UI_WinRadioButton_Operations("Fn_SISW_RAC_PMM_UserContextSettings", "set", objWinDialog, "Version", "")
'Function Usage		     :	Call Fn_WIN_UI_WinRadioButton_Operations("Fn_SISW_RAC_PMM_UserContextSettings", "set", objWinDialog.WinRadioButton("Precise"), "" , "")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinRadioButton_Operations(sFunctionName, sAction, objWinDialog, sWinRadioButtonName, sValue)
	'Declaring Variables
	Dim objRadioButton
	
	'Initially set function return value as False
	Fn_WIN_UI_WinRadioButton_Operations = False
	
	'Creating/Setting Object of WinRadioButton
	If sWinRadioButtonName <> "" Then
		Set objRadioButton = objWinDialog.WinRadioButton(sWinRadioButtonName)
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " &  objRadioButton.toString & " ] : Action = " & sAction & " : "
	Else
		Set objRadioButton = objWinDialog
		GBL_FUNCTIONLOG = " [ " &  objRadioButton.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify WinRadioButton object exists
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinRadioButton_Operations", "Exist", objRadioButton,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinRadioButton_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinRadioButton Object does not exist")
		'Release Object
		Set objRadioButton = Nothing 
		Exit Function
	End If
	
	Select case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify contents of WinList
		Case "set"
			objRadioButton.Set
			Fn_WIN_UI_WinRadioButton_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinRadioButton_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully set WinRadioButton [ "& objRadioButton.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" Fn_WIN_UI_WinRadioButton_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinRadioButton_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinRadioButton_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" Fn_WIN_UI_WinRadioButton_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release object
	Set objRadioButton = Nothing 
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinComboBox_Operations
'
'Function Description	:	Function used to perform operations on WinComboBox object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWinDialog: Object of Container  in which WinComboBox is present
'							4.sWinComboBox: Repository Name of WinComboBox
'							5.sValues: Value to be selected / verified
'							6.sColumns: Column value to be selected
'							7.sInstanceHandler: Instance Handler
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WIN_UI_WinComboBox_Operations("Fn_TeamcenterLogin", "getcontents", Dialog("Properties"),"Source","", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinComboBox_Operations("Fn_TeamcenterLogin", "gettext", Dialog("Properties"),"Source","", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinComboBox_Operations("Fn_TeamcenterLogin", "click", Dialog("Properties"),"Source","UnKnown", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinComboBox_Operations("Fn_TeamcenterLogin", "activate", Dialog("Properties"),"Source","UnKnown", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinComboBox_Operations("Fn_TeamcenterLogin", "exist", Dialog("Properties"),"Source","UnKnown", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinComboBox_Operations("Fn_TeamcenterLogin", "extendselect", Dialog("Properties"),"Source","UnKnown~Precise~Rule", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinComboBox_Operations("Fn_TeamcenterLogin", "type", Dialog("Properties"),"Source","UnKnown", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinComboBox_Operations("Fn_TeamcenterLogin", "select", Dialog("Properties"),"Source","UnKnown", "", "")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinComboBox_Operations(sFunctionName, sAction, objWinDialog, sWinComboBox, sValues, sColumns, sInstanceHandler)
	'Declaring Variables
	Dim objWinComboBox
	Dim	aSelectList
	Dim iCounter, iInstanceCount, iEelementCount

	'Initially set function return value as False
	Fn_WIN_UI_WinComboBox_Operations = False
	
	If sInstanceHandler = "" Then
		sInstanceHandler = "@"
	End If
	
	'Creating/Setting Object of WinComboBox
	If sWinComboBox <> "" Then
		Set objWinComboBox = objWinDialog.WinComboBox(sWinComboBox)
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " +  objWinComboBox.toString + " ] : Action = " & sAction & " : "
	Else
		Set objWinComboBox = objWinDialog
		GBL_FUNCTIONLOG = " [ " & objWinComboBox.toString & " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify WinComboBox object exists
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinComboBox_Operations", "Exist", objWinComboBox,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinComboBox Object does not exist")
		'Release Object
		Set objWinComboBox = Nothing 
		Exit Function
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the contents of WinComboBox
		Case "getcontents"
			iEelementCount = objWinComboBox.GetROProperty("items count")
			For iCounter = 0 To iEelementCount - 1
				If iCounter = 0 Then
					Fn_WIN_UI_WinComboBox_Operations = Trim(cstr(objWinComboBox.GetItem(iCounter)))
				Else
					Fn_WIN_UI_WinComboBox_Operations = Fn_WIN_UI_WinComboBox_Operations & "~" & Trim(cstr(objWinComboBox.GetItem(iCounter)))
				End If
			Next
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the contents [ "& Fn_WIN_UI_WinComboBox_Operations &" ] from WinComboBox [ "& objWinComboBox.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the text value of WinComboBox
		Case "gettext"
			Fn_WIN_UI_WinComboBox_Operations = objWinComboBox.GetROProperty("value")
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the text value [ "& Fn_WIN_UI_WinComboBox_Operations &" ] of WinComboBox [ "& objWinComboBox.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click on WinComboBox
		Case "click"
			objWinComboBox.Click sValues, sColumns
			Fn_WIN_UI_WinComboBox_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully clicked the value [ "& sValues &":"& sColumns &" ] of WinComboBox [ "& objWinComboBox.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to activate the WinComboBox
		Case "activate"
			objWinComboBox.Activate sValues
			Fn_WIN_UI_WinComboBox_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully activated the value [ "& sValues &" ] of WinComboBox [ "& objWinComboBox.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify contents of WinComboBox
		Case "exist"
			'get total items from list
			iEelementCount = objWinComboBox.GetROProperty("items count")
			iInstanceCount = 1
			aSelectList = split(sValues, sInstanceHandler)
			If uBound(aSelectList) > 0 Then
				iInstanceCount = aSelectList(1)
			End If
			aSelectList(0) = trim(aSelectList(0))
			For iCounter = 0 To iEelementCount - 1
				If objWinComboBox.GetItem(iCounter) <> "" Then
					If Trim(cstr(objWinComboBox.GetItem(iCounter))) = Trim(aSelectList(0)) Then
						If iInstanceCount = 1 Then
							Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully verified WinComboBox value [ "& cstr(objWinComboBox.GetItem(iCounter))&" ] exist")
							Fn_WIN_UI_WinComboBox_Operations = True
							Exit For
						End If
						iInstanceCount = iInstanceCount - 1 
					End If
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to extend the WinComboBox selection
		Case "extendselect"
			aSelectList = Split(sValues,"~")
			'Select the element from list  
			For iCounter = 0 To Ubound(aSelectList)
				objWinComboBox.ExtendSelect aSelectList(iCounter)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected WinComboBox value [ "& aSelectList(iCounter)&" ]")
			Next
			Fn_WIN_UI_WinComboBox_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to type value in WinComboBox
		Case "type"
			objWinComboBox.Type sValues
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully type WinComboBox value [ "& sValues &" ]")
			Fn_WIN_UI_WinComboBox_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify contents of WinList
		Case "select"
			'get total items from list
			iEelementCount = objWinComboBox.GetROProperty("items count")
			iInstanceCount = 1
			aSelectList = split(sValues, sInstanceHandler)
			If uBound(aSelectList) > 0 Then
				iInstanceCount = aSelectList(1)
			End If
			aSelectList(0) = trim(aSelectList(0))
			For iCounter = 0 To iEelementCount - 1
				If objWinComboBox.GetItem(iCounter) <> "" Then
					If Trim(cstr(objWinComboBox.GetItem(iCounter))) = Trim(aSelectList(0)) Then
						iF iInstanceCount = 1 Then
							objWinComboBox.Select iCounter
							Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected WinComboBox value [ "& cstr(objWinComboBox.GetItem(iCounter))&" ]")
							Fn_WIN_UI_WinComboBox_Operations = True
							Exit For
						End If
						iInstanceCount = iInstanceCount - 1 
					End If
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinComboBox_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinComboBox_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinComboBox_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release object
	Set objWinComboBox = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinCheckBox_Operations
'
'Function Description	:	Function used to perform operations on WinCheckBox object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWinDialog: Object of Container  in which WinCheckBox is present
'							4.sWinCheckBoxName: Repository Name of CheckBox
'							5.sValue: ON/OFF
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WIN_UI_WinCheckBox_Operations("", "set", objWinDialog, "Version", "ON")
'Function Usage		     :	Call Fn_WIN_UI_WinCheckBox_Operations("", "set", objWinDialog.WinCheckBox("Precise"), "" , "OFF")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinCheckBox_Operations(sFunctionName, sAction, objWinDialog, sWinCheckBoxName, sValue)
	'Declaring Variables
	Dim objWinCheckBox
	
	'Initially set function return value as False
	Fn_WIN_UI_WinCheckBox_Operations = False
	
	'Creating/Setting Object of WinCheckBox
	If sWinCheckBoxName <> "" Then
		Set objWinCheckBox = objWinDialog.WinCheckBox(sWinCheckBoxName)
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " +  objWinCheckBox.toString + " ] : Action = " & sAction & " : "
	Else
		Set objWinCheckBox = objWinDialog
		GBL_FUNCTIONLOG = " [ " +  objWinCheckBox.toString + " ] : Action = " & sAction & " : "
	End If
	
	'Verify WinCheckBox object exists
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinCheckBox_Operations", "Exist", objWinCheckBox,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinCheckBox_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinCheckBox Object does not exist")
		'Release Object
		Set objWinCheckBox = Nothing 
		Exit Function
	End If
	
	Select case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set the value of WinCheckBox
		Case "set"		
			objWinCheckBox.Set uCase(sValue)
			Fn_WIN_UI_WinCheckBox_Operations = True
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinCheckBox_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully set WinCheckBox value [ "& sValue &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),,sFunctionName &" >> Fn_WIN_UI_WinCheckBox_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinCheckBox_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinCheckBox_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinCheckBox_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release object
	Set objWinCheckBox = Nothing 
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinEditor_Operations
'
'Function Description	:	Function used to perform operations on Win MultiLine Editor object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWinDialog: Object of Container  in which WinEditor is present
'							4.sWinEditor: Repository Name of WinEditor
'							5.sText: Value to be set/verified
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WIN_UI_WinEditor_Operations("Fn_Catia_PartCreate", "set",  objWinDialog, "Name", "SQS_Part_Name" )
'Function Usage		     :	Call Fn_WIN_UI_WinEditor_Operations("Fn_Catia_PartCreate", "type",  objWinDialog.WinEdit("Name"), "", "SQS_Part_Name" )
'Function Usage		     :	Call Fn_WIN_UI_WinEditor_Operations("Fn_Catia_PartCreate", "gettext",  objWinDialog.WinEdit("Name"),"", "" )
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinEditor_Operations(sFunctionName, sAction, objWinDialog, sWinEditor, sText)
	'Declaring Variables
	Dim objWinEditor
	
	'Initially set function return value as False
	Fn_WIN_UI_WinEditor_Operations = False
	
	'Creating/Setting Object of WinEditor
	If sWinEditor <> "" Then
		Set objWinEditor= objWinDialog.WinEditor(sWinEditor)
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " & objWinEditor.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWinEditor= objWinDialog
		GBL_FUNCTIONLOG = " [ " & objWinEdit.toString & " ] : Action = " & sAction & " : "
	End If
	
	'Verify WinEditor object exists
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinEditor_Operations", "Exist", objWinEditor,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEditor_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinEditor Object does not exist")
		'Release Object
		Set objWinEditor = Nothing 
		Exit Function
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set type the value in WinEditor
		Case "type"
			If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinEditor_Operations", "Enabled", objWinEditor,"","","") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEditor_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinEditor Object is not enabled")
				'Release Object
				Set objWinEditor = Nothing 
				Exit Function
			End If
			'Setting the WinEditor
			objWinEditor.Type sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEditor_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully type the value [ "& sText &" ] in WinEditor [ "& objWinEditor.toString &" ]")
			Fn_WIN_UI_WinEditor_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the value of WinEditor
		Case "gettext"
			Fn_WIN_UI_WinEditor_Operations = objWinEditor.getROProperty("text")
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEditor_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the value [ "& Fn_WIN_UI_WinEditor_Operations &" ] from WinEditor [ "& objWinEditor.toString &" ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEditor_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinEditor_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinEditor_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinEditor_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release object
	Set objWinEditor = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinTab_Operations
'
'Function Description	:	Function used to perform operations on WinTab object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'       					3.objWinDialog: Parent Hierarchy of WinTab
'       					4.sWinTab: Name of WinTab object
'       					5.sTabName: Tab to be select
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WIN_UI_WinTab_Operations("Fn_Catia_PartCreate", "select",  objWinDialog, "MainTab", "Export Image" )
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinTab_Operations(sFunctionName, sAction, objWinDialog, sWinTab, sTabName)
	'Declaring Variables
	Dim objWinTab
	
	'Initially set function return value as False
	Fn_WIN_UI_WinTab_Operations = False
	
	'Creating/Setting Object of WinTab
	If sWinTab <> "" Then
		Set objWinTab = objWinDialog.WinTab(sWinTab)
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " & objWinTab.toString & " ] : Action = " & sAction & " : "
	Else
		Set objWinTab= objWinDialog
		GBL_FUNCTIONLOG = " [ " & objWinTab.toString & " ] : Action = " & sAction & " : "
	End If
	
	'Verify WinTab object exist
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinTab_Operations", "Exist", objWinTab,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTab_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinTab Object does not exist")
		'Release Object
		Set objWinTab = Nothing 
		Exit Function
	End If
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select the WinTab
		Case "select"
			If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinTab_Operations", "Enabled", objWinTab,"","","") = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTab_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinTab Object is not enabled")
				'Release Object
				Set objWinTab = Nothing 
				Exit Function
			End If
            'Selecting the WinTab
			objWinTab.Select sTabName
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTab_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected the WinTab [ "& sTabName &" ]")
			Fn_WIN_UI_WinTab_Operations= True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTab_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinTab_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinTab_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTab_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release object
	Set objWinTab = Nothing
End Function

'								### This function is not tested###
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinTable_Operations
'
'Function Description	:	Function used to perform operations on WinTable object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWinDialog: Object of Container  in which WinButton is present
'							4.sWinTable: Repository Name of WinTable
'							5.iRow: Row number value
'							6.sColumn: Column name or number value
'							7.sValue: Value to be set/verified
'							8.sMicButton: Mouse button click name
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WIN_UI_WinTable_Operations("", "activatecell", objWinDialog, VehicleProgramTable, 1, "Name", "", "LEFT")
'Function Usage		     :	Call Fn_WIN_UI_WinTable_Operations("", "getcelldata", objWinDialog.Wintable("VehicleProgramTable"), "", 1, "Name", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinTable_Operations("", "getcolumnname", objWinDialog, VehicleProgramTable, "", "3", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinTable_Operations("", "selectcell", objWinDialog, VehicleProgramTable, 1, "Name", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinTable_Operations("", "selectcolumn", objWinDialog, VehicleProgramTable, "", "Name", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinTable_Operations("", "selectrow", objWinDialog, VehicleProgramTable, 1, "", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinTable_Operations("", "setcelldata", objWinDialog, VehicleProgramTable, 1, "Name", "User", "")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinTable_Operations(sFunctionName, sAction, objWinDialog, sWinTable, iRow, sColumn, sValue, sMicButton)
	'Declaring Variables
	Dim objWinTable
	
	'Initially set function return value as False
	Fn_WIN_UI_WinTable_Operations = False
	
	'Creating/Setting Object of WinTable
	If sWinTable <> "" Then
		Set objWinTable = objWinDialog.WinTable(sWinTable)	
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " +  sWinTable + " ] : Action = " & sAction & " : "
	Else
		Set objWinTable = objWinDialog
		GBL_FUNCTIONLOG = " [ " & objWinTable.toString() & " ] : Action = " & sAction & " : "	
	End If	
	
	'Verify WinTab object exist
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinTable_Operations", "Exist", objWinTable,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTable_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinTable Object does not exist")
		'Release Object
		Set objWinTable = Nothing 
		Exit Function
	End If
		
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to activate a cell of the WinTable
		Case "activatecell"
			If Cint(objWinTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTable_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Row [ "& iRow &" ] does not exist in Wintable")
			Else
				If sMicButton <> "" Then
					objWinTable.ActivateCell iRow,sColumn,sMicButton
				Else
					objWinTable.ActivateCell iRow,sColumn
				End If
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTable_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully activated the cell of WinTable [ " &  objWinTable.toString & " ] at row [ "& iRow &" ] and Column [ "& sColumn &" ]")
				Fn_WIN_UI_WinTable_Operations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the contents of specified cell of the WinTable
		Case "getcelldata"
			If Cint(objWinTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTable_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Row [ "& iRow &" ] does not exist in Wintable")
			Else
				Fn_WIN_UI_WinTable_Operations = objWinTable.GetCellData(iRow,sColumn)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTable_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the cell data [ "& Fn_WIN_UI_WinTable_Operations &" ] of WinTable [ " &  objWinTable.toString & " ] at row [ "& iRow &" ] and Column [ "& sColumn &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the column name of the WinTable
		Case "getcolumnname"
			If objWinTable.ColumnCount > -1 Then
				Fn_WIN_UI_WinTable_Operations = objWinTable.GetColumnName(sColumn)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTable_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the column name [ "& Fn_WIN_UI_WinTable_Operations &" ] of WinTable [ " &  objWinTable.toString & " ] at Column [ "& sColumn &" ]")
			Else	
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTable_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Column [ "& sColumn &" ] does not exist in Wintable")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select the cell of the WinTable
		Case "selectcell"
			If Cint(objWinTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Row [ "&iRow&" ] does not exist in Wintable")
			Else
				objWinTable.SelectCell iRow,sColumn
				Fn_WIN_UI_WinTable_Operations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected the cell of WinTable [ " &  objWinTable.toString & " ] at row [ "& iRow &" ] and Column [ "& sColumn &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select the column of the WinTable
		Case "selectcolumn"
			If objWinTable.ColumnCount > -1 Then
				objWinTable.SelectColumn(sColumn)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected the column [ "& sColumn &" ] of WinTable [ " &  objWinTable.toString & " ]")
				Fn_WIN_UI_WinTable_Operations = True
			Else	
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> FAIL : "&GBL_FUNCTIONLOG&"Column [ "& sColumn &" ] does not exist in Wintable")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select the row of the WinTable
		Case "selectrow"
			If Cint(objWinTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Row [ "& iRow &" ] does not exist in Wintable")
			Else
				objWinTable.SelectRow iRow
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected the Row [ "& iRow &" ] of WinTable [ " &  objWinTable.toString & " ]")
				Fn_WIN_UI_WinTable_Operations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set the contents of a cell of the WinTable
		Case "setcelldata"
			If Cint(objWinTable.GetROProperty("rows")) < Cint(iRow) Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> FAIL : "& GBL_FUNCTIONLOG &"Row [ "& iRow &" ] does not exist in Wintable")
			Else
				objWinTable.SetCellData iRow,sColumn,sValue
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully set the cell data [ "& sValue &" ] at the Row [ "& iRow &" ] and Column [ "& sColumn &" ] of WinTable [ " &  objWinTable.toString & " ]")
				Fn_WIN_UI_WinTable_Operations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinTable_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinTable_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName&" >> Fn_WIN_UI_WinTable_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release object
	Set objWinTable = Nothing
End Function

'								### This function is not tested###
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinToolbar_Operations
'
'Function Description	:	Function used to perform operations on WinToolbar object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWinDialog: Object of Container  in which WinButton is present
'							4.sWinToolbar: Repository Name of WinToolbar
'							5.sWinToolbarButtonName: Win Toolbar Button Name
'							6.sValue: Value to be set/verified
'							7.iIndex: Instance number
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WIN_UI_WinToolbar_Operations("", "click", objWinDialog, "FolderWidgetToolbar", "Open", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinToolbar_Operations("", "getcontents", objWinDialog.WinToolbar("FolderWidgetToolbar"), "", "", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinToolbar_Operations("", "showdropdownmenu", objWinDialog, "FolderWidgetToolbar", "Open", "", "")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinToolbar_Operations(sFunctionName, sAction, objWinDialog, sWinToolbar, sWinToolbarButtonName, sValue, iIndex)
	'Declaring Variables
	Dim objWinToolbar
	
	'Initially set function return value as False
	Fn_WIN_UI_WinToolbar_Operations = False
	
	
	'Creating/Setting Object of WinToolbar
	If sWinToolbar <> "" Then
		Set objWinToolbar = objWinDialog.WinToolbar(sWinToolbar)	
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " &  sWinToolbar & " ] : Action = " & sAction & " : "
	Else
		Set objWinToolbar = objWinDialog
		GBL_FUNCTIONLOG = " [ " & objWinToolbar.toString() & " ] : Action = " & sAction & " : "	
	End If	
	
	'Verify WinToolbar object exist
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinToolbar_Operations", "Exist", objWinToolbar,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinToolbar_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinToolbar Object does not exist")
		'Release Object
		Set objWinToolbar = Nothing 
		Exit Function
	End If
		
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to click an object of the WinToolbar
		Case "click"
			objWinToolbar.Press sWinToolbarButtonName
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinToolbar_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully clicked the Toolbar button [ "& sWinToolbarButtonName &" ] of WinToolbar [ " &  objWinToolbar.toString & " ]")
			Fn_WIN_UI_WinToolbar_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the names of all the buttons in the WinToolbar
		Case "getcontents"
			Fn_WIN_UI_WinToolbar_Operations = objWinToolbar.GetContent
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinToolbar_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully get the Toolbar buttons [ "& Fn_WIN_UI_WinToolbar_Operations &" ] of WinToolbar [ " &  objWinToolbar.toString & " ]")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to open the dropdown menu associated with the toolbar button of the WinToolbar
		Case "showdropdownmenu"
			wait GBL_MICRO_TIMEOUT
			objWinToolbar.ShowDropdown sWinToolbarButtonName
			wait GBL_MICRO_TIMEOUT
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinToolbar_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully open the dropdown menu of Toolbar button [ "& sWinToolbarButtonName &" ] of WinToolbar [ " &  objWinToolbar.toString & " ]")
			Fn_WIN_UI_WinToolbar_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinToolbar_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinToolbar_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinToolbar_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinToolbar_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release object
	Set objWinToolbar = Nothing
End Function

'								### This function is not tested###
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_WIN_UI_WinTreeView_Operations
'
'Function Description	:	Function used to perform operations on WinTreeView object.
'
'Function Parameters	:   1.sFunctionName: Name of function
'						    2.sAction: Action name
'							3.objWinDialog: Object of Container  in which WinButton is present
'							4.sWinTreeView: Repository Name of WinTreeView
'							5.sNodeName: Name of the node
'							6.sWinMenu: WinMenu to be select
'							7.sPopupMenu: Popup Menu to be select
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_WIN_UI_WinTreeView_Operations("", "expand", objWinDialog, "NavTree", "Home~123-Item1", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinTreeView_Operations("", "collapse", objWinDialog, "NavTree", "Home~123-Item1", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinTreeView_Operations("", "select", objWinDialog, "NavTree", "Home~123-Item1", "", "")
'Function Usage		     :	Call Fn_WIN_UI_WinTreeView_Operations("", "doubleclick", objWinDialog, "NavTree", "Home~123-Item1", "", "")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  03-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_WIN_UI_WinTreeView_Operations(sFunctionName, sAction, objWinDialog, sWinTreeView, sNodeName, sWinMenu, sPopupMenu)
	'Declaring Variables
	Dim objWinTree
	
	'Initially set function return value as False
	Fn_WIN_UI_WinTreeView_Operations = False
	
	'Creating/Setting Object of WinTreeView
	If sWinTreeView <> "" Then
		Set objWinTree = objWinDialog.WinTreeView(sWinTreeView)	
		GBL_FUNCTIONLOG = " [ " &  objWinDialog.toString & " ] : [ " & sWinTreeView & " ] : Action = " & sAction & " : "
	Else
		Set objWinTree = objWinDialog
		GBL_FUNCTIONLOG = " [ " & objWinTree.toString() & " ] : Action = " & sAction & " : "	
	End If	
	
	'Verify WinTreeView object exist
	If Fn_WIN_UI_WinObject_Operations("Fn_WIN_UI_WinTreeView_Operations", "Exist", objWinTree,"","","") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTreeView_Operations >> FAIL : "& GBL_FUNCTIONLOG &"WinTreeView Object does not exist")
		'Release Object
		Set objWinTree = Nothing 
		Exit Function
	End If
		
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to expand the node in the WinTreeView
		Case "expand"	
			objWinTree.Expand sNodeName		
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTreeView_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully expanded the node [ "& sNodeName &" ] of WinTreeView [ " &  objWinTree.toString & " ]")
			Fn_WIN_UI_WinTreeView_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to collapse the node in the WinTreeView
		Case "collapse"	
			objWinTree.Collapse sNodeName		
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTreeView_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully Collapsed the node [ "& sNodeName &" ] of WinTreeView [ " &  objWinTree.toString & " ]")
			Fn_WIN_UI_WinTreeView_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select the node in the WinTreeView
		Case "select"	
			objWinTree.Select sNodeName		
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTreeView_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully selected the node [ "& sNodeName &" ] of WinTreeView [ " &  objWinTree.toString & " ]")
			Fn_WIN_UI_WinTreeView_Operations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to doubleclick the node in the WinTreeView
		Case "doubleclick"	
			objWinTree.Activate sNodeName		
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTreeView_Operations >> PASS : "& GBL_FUNCTIONLOG &"Successfully doubleclicked the node [ "& sNodeName &" ] of WinTreeView [ " &  objWinTree.toString & " ]")
			Fn_WIN_UI_WinTreeView_Operations = True
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'More cases to add for Popup menu and other Win menu operations
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTreeView_Operations >> FAIL : [No valid case] : No valid case was passed for function [Fn_WIN_UI_WinTreeView_Operations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_WIN_UI_WinTreeView_Operations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFunctionName &" >> Fn_WIN_UI_WinTreeView_Operations >> FAIL : To perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release object
	Set objWinTree = Nothing
End Function