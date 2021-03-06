'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'01. Fn_UI_SwfButton_Operations
'02. Fn_UI_SwfEdit_Operations
'03. Fn_UI_SwfObject_Operations
'04. Fn_UI_GenSwfObject_Operations
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Function Name		 	:	Fn_UI_SwfButton_Operations
'
'Description		    :  	This function is used to perform operations on SwfButton object.
'
'Parameters			    :	1. sFunctionName : Valid Function name, 
'							2. sAction
'							3. objSwfDialog
'							4. sSwfButton   : Valid Button name
'							5. iX   : X cordinate
'							6. iY   : Y cordinate
'							7. sMicButton   : Button type
'
'Return Value		    :  	True \ False
'
'Examples		     	:	Call Fn_UI_SwfButton_Operations("Fn_CatiaLogin", "Click", Window("Catia Login"),"Login","","","")
'Examples		     	:	Call Fn_UI_SwfButton_Operations("Fn_CatiaLogin", "Click", Window("Catia Login").SwfButton("Login"),"","1","1","")
'Examples		     	:	Call Fn_UI_SwfButton_Operations("Fn_CatiaLogin", "Click", Window("Catia Login").SwfButton("Login"),"","1","1","micLeftBtn")
'
'History:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep Navghane		 	| 10-May-2015	|	1.0			|	Sandeep Navghane	 		| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Function Fn_UI_SwfButton_Operations(sFunctionName, sAction, objSwfDialog, sSwfButton,iX,iY,sMicButton)
	Dim objSwfButton, sFuncLog, objDeviceReplay
	Fn_UI_SwfButton_Operations = False
	'Object Creation
	If sSwfButton <> "" Then
		Set objSwfButton = objSwfDialog.SwfButton(sSwfButton)
		sFuncLog = sFunctionName + " > Fn_UI_SwfButton_Operations  : [ " &  objSwfDialog.toString & " ] : [ " +  objSwfButton.toString + " ] : Action = " & sAction & " : "
	Else
		Set objSwfButton = objSwfDialog
		sFuncLog = sFunctionName + " > Fn_UI_SwfButton_Operations  : [ " +  objSwfButton.toString + " ] : Action = " & sAction & " : "
	ENd IF
	
	'Verify SwfButton object exists
	If Fn_UI_SwfObject_Operations("Fn_UI_SwfButton_Operations", "Enabled", objSwfButton,"") = False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Set objSwfButton = Nothing 
		Exit Function
	End If
	
	Select Case sAction
		Case "Click"
			If iX<>"" and iY<>"" Then
				If sMicButton<>"" Then
					objSwfButton.Click iX,iY,sMicButton
				Else
					objSwfButton.Click iX,iY
				End If
			Else
				objSwfButton.Click
			End If	
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully clicked on SwfButton.")
			Fn_UI_SwfButton_Operations = True
		Case "DeviceReplay.Click"
			If sSwfButton <> "" Then
				objSwfDialog.Activate
				wait GBL_MICRO_TIMEOUT
			End If
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			objDeviceReplay.MouseMove (objSwfButton.GetROProperty("abs_x") + 5), (objSwfButton.GetROProperty("abs_y") + 5)
			objDeviceReplay.MouseClick  (objSwfButton.GetROProperty("abs_x") + 5), (objSwfButton.GetROProperty("abs_y") + 5), 0
			Fn_UI_SwfButton_Operations = True
		Case "DoubleClick"
			If sMicButton<>"" Then
				objSwfButton.DblClick iX,iY,sMicButton
			Else
				objSwfButton.DblClick iX,iY
			End If
					
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & "PASS : Sucessfully clicked on SwfButton.")
			Fn_UI_SwfButton_Operations = True
		Case Else
	End Select
	'Clear memory of SwfButton object.
	Set objDeviceReplay = Nothing
	Set objSwfButton = Nothing 
End Function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Function Name		 	:	Fn_UI_SwfEdit_Operations
'
'Description		    :  	This function is use to perform operations on SwfEditBox.

'Parameters			    :	1. sFunctionName : Valid Function name, 
'							2. sAction : Action name
'							3. objSwfDialog : Parent dialog object
'							4. sSwfEdit : Swf edit box name
'							5. sText : Text as value

'Return Value		    :  	True \ false
'
'Examples		     	:	Call Fn_UI_SwfEdit_Operations("Fn_Catia_PartCreate", "Set",  objSwfDialog, "Name", "SQS_Part_Name" )
'Examples		     	:	Call Fn_UI_SwfEdit_Operations("Fn_Catia_PartCreate", "Type",  objSwfDialog.SwfEdit("Name"),"", "SQS_Part_Name" )

'History:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep Navghane		 	| 10-May-2015	|	1.0			|	Sandeep Navghane	 		| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Function Fn_UI_SwfEdit_Operations(sFunctionName, sAction, objSwfDialog, sSwfEdit, sText)
	Dim objSwfEdit, sFuncLog
	Fn_UI_SwfEdit_Operations = False
	'Set an Edit Object on variable
	If sSwfEdit <> "" Then
		Set objSwfEdit= objSwfDialog.SwfEdit(sSwfEdit)
		sFuncLog = sFunctionName + "> Fn_UI_SwfEdit_Operations : [ " &  objSwfDialog.toString & " ] : [ " & objSwfEdit.toString & " ] : Action = " & sAction & " : "
	Else
		Set objSwfEdit= objSwfDialog
		sFuncLog = sFunctionName + "> Fn_UI_SwfEdit_Operations : [ " & objSwfEdit.toString & " ] : Action = " & sAction & " : "
	End If

	If Fn_UI_SwfObject_Operations("Fn_UI_SwfEdit_Operations", "Exist", objSwfEdit,"") = False Then
		Fn_UI_SwfEdit_Operations= False
		Set objSwfEdit = Nothing 
		Exit Function
	End If
	Select Case sAction
		Case "Set"
			'Setting the editbox
			If Fn_UI_SwfObject_Operations("Fn_UI_SwfEdit_Operations", "Enabled", objSwfEdit,"") = False Then
				Fn_UI_SwfEdit_Operations= False
				Set objSwfEdit = Nothing 
				Exit Function
			End If
			objSwfEdit.Set sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & "PASS : Text " &sText  &"is Set/ Entered in SwfEditBox.")
			Fn_UI_SwfEdit_Operations= True

		Case "Type"
			If Fn_UI_SwfObject_Operations("Fn_UI_SwfEdit_Operations", "Enabled", objSwfEdit,"") = False Then
				Fn_UI_SwfEdit_Operations= False
				Set objSwfEdit = Nothing 
				Exit Function
			End If
			'Setting the editbox
			objSwfEdit.Type sText
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & "PASS : Text " &sText  &" is Set/ Entered in SwfEditBox.")
			Fn_UI_SwfEdit_Operations= True
		Case "GetText"
			Fn_UI_SwfEdit_Operations = objSwfEdit.getROProperty("text")
		Case "SendString"
			If sSwfEdit <> "" Then
				objSwfEdit.SetFocus
				wait GBL_MICRO_TIMEOUT
			End If 
			Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
			'Setting the editbox
			objDeviceReplay.SendString sText
	End Select
	Set objSwfEdit = Nothing
	Set objDeviceReplay = Nothing
End Function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Function Name		:	Fn_UI_SwfObject_Operations

'Description		:	This function is Use to perform operations on Swf Object.

'Parameters		    :	1. sFunctionName: Valid Function Name
'						    		2. sAction: Action name to perform
'						    		3. objReferencePath: Valid Swf Dialog	/Window
'				    				4. iTimeOut: Time out time in seconds.
'
'Return Value		: 	TRUE \ FALSE
'
'Examples		     	:	Fn_UI_SwfObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Create", objDateControl,"")
'Examples		     	:	Fn_UI_SwfObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Enabled", objDateControl, MAX_TIMEOUT)
'Examples		     	:	Fn_UI_SwfObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Exist", objDateControl,"")
'
'History			:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep Navghane		 	| 10-May-2015	|	1.0			|	Sandeep Navghane 		| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Function Fn_UI_SwfObject_Operations(sFunctionName, sAction, objReferencePath, iTimeOut)
	Dim sFuncLog
	sFuncLog = sFunctionName + " > Fn_UI_SwfObject_Operations : [ " &  objReferencePath.toString & " ] : Action = " & sAction & " : "
	Fn_UI_SwfObject_Operations = False
	
	If iTimeOut = "" Then
		iTimeOut = GBL_DEFAULT_TIMEOUT
	Else
		iTimeOut = cInt(iTimeOut)
	End If
	
	Select Case sAction
		Case "Create"
			If Fn_UI_SwfObject_Operations(sFunctionName, "Enabled", objReferencePath, iTimeOut) Then
				Set Fn_UI_SwfObject_Operations = objReferencePath
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "New Object created for " & objReferencePath.toString & " in Function ")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "Object " & objReferencePath.toString & " is not enable of Function ")
			End If
		Case "Exist"
			Fn_UI_SwfObject_Operations = objReferencePath.Exist(iTimeOut)
			If Fn_UI_SwfObject_Operations Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & " object is exist.")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "object does not exist.")
			End If
		Case "Enabled"
			If Fn_UI_SwfObject_Operations(sFunctionName, "Exist", objReferencePath, iTimeOut) Then
				If objReferencePath.GetROProperty("enabled") = "1"  OR objReferencePath.GetROProperty("enabled") = True Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & "object is exists and enabled.")
					Fn_UI_SwfObject_Operations = True
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "object is not enabled.")
				End If
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "does not exist.")
			End If
	End Select
End Function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Function Name		:	Fn_UI_GenSwfObject_Operations

'Description		:	This function to perform operations on Generic Swf objects

'Parameters		    :	1. sFunctionName:
'								2. sAction		:
'								3. objContainer : Prent UI Component or JavaTable object
'								4. sSwfObject 	: Twistie Control name
'								5. sSwfObjectText	: Twistie Control text
'								6. iX : X coordinate
'								7. iX : Y coordinate
'								7. sMicButton : mouse button name
'
'Return Value		: 	TRUE \ FALSE
'
'Examples			:	Fn_UI_GenSwfObject_Operations("", "GetTextLocationAndClick", Window("SolidWorksDefaultWindow"),"SwimManager", "Open","","","","")
'
'History			:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep Navghane		 	| 10-May-2015	|	1.0			|	Sandeep Navghane 		| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Function Fn_UI_GenSwfObject_Operations(sFunctionName, sAction, objContainer, sSwfObject, sSwfObjectText,iX,iY,sValue,sMicButton)
	Dim objSwfObject, objSwfObjectText
	Dim iLeft,iTop,iRight,iBottom
	Dim sFuncLog
	Dim bFlag
	
	Set objSwfObject = objContainer.SwfObject(sSwfObject)
	
	Fn_UI_GenSwfObject_Operations = False
	
	sFuncLog = sFunctionName + " > Fn_UI_GenSwfObject_Operations : [ " &  objSwfObject.toString & " ] : Action = " & sAction & " : "
	If objSwfObject.Exist(5)= False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Call Fn_UI_ExitFromUI(sFuncLog)
		Set objSwfObject = Nothing 
		Exit Function
	End If

	Select Case sAction
		Case "GetTextLocationAndClick"
			objSwfObject.highlight
			wait GBL_MICRO_TIMEOUT
			
			bFlag=objSwfObject.GetTextLocation(sSwfObjectText,iLeft,iTop,iRight,iBottom,True)
			If bFlag=True  Then
				objSwfObject.Click(iLeft+iRight)/2 ,(iTop+iBottom)/2			
				If Err.Number < 0 Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  Fail to click on Swf Object [ " & Cstr(sSwfObjectText) & " ] with Error Number [" & Cstr(Err.Number) & "] Error Description  [" & Cstr(Err.Description) & "]" )
					Fn_UI_GenSwfObject_Operations=False
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<PASS>:  Successfully click on Swf Object [ " & Cstr(sSwfObjectText) & " ]")
					Fn_UI_GenSwfObject_Operations=True
				End If
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  Fail to click on Swf Object [ " & Cstr(sSwfObjectText) & " ] as object does not exist" )
				Fn_UI_GenSwfObject_Operations=False
			End If	
	End Select
	Set objSwfObject = Nothing 
End Function