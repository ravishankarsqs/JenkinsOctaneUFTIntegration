'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'01. Fn_UI_WpfObject_Operations
'02. Fn_UI_GenWpfObject_Operations
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Function Name		:	Fn_UI_WpfObject_Operations

'Description		:	This function is Use to perform operations on Swf Object.

'Parameters		    :	1. sFunctionName: Valid Function Name
'						    		2. sAction: Action name to perform
'						    		3. objReferencePath: Valid Swf Dialog	/Window
'				    				4. iTimeOut: Time out time in seconds.
'
'Return Value		: 	TRUE \ FALSE
'
'Examples		     	:	Fn_UI_WpfObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Create", objDateControl,"")
'Examples		     	:	Fn_UI_WpfObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Enabled", objDateControl, MAX_TIMEOUT)
'Examples		     	:	Fn_UI_WpfObject_Operations("Fn_SISW_RAC_UI_DateControl_SetDate","Exist", objDateControl,"")
'
'History			:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Ashok Kakade			 	| 21-10-2015	|	1.0			|	Sandeep Navghane 		| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Function Fn_UI_WpfObject_Operations(sFunctionName, sAction, objReferencePath, iTimeOut)
	Dim sFuncLog
	sFuncLog = sFunctionName + " > Fn_UI_WpfObject_Operations : [ " &  objReferencePath.toString & " ] : Action = " & sAction & " : "
	Fn_UI_WpfObject_Operations = False
	
	If iTimeOut = "" Then
		iTimeOut = GBL_DEFAULT_TIMEOUT
	Else
		iTimeOut = cInt(iTimeOut)
	End If
	
	Select Case sAction
		Case "Create"
			If Fn_UI_WpfObject_Operations(sFunctionName, "Enabled", objReferencePath, iTimeOut) Then
				Set Fn_UI_WpfObject_Operations = objReferencePath
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "New Object created for " & objReferencePath.toString & " in Function ")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "Object " & objReferencePath.toString & " is not enable of Function ")
				Call Fn_UI_ExitFromSwfUI(sFuncLog & "Object " & objReferencePath.toString & " is not enable of Function ") 
			End If
		Case "Exist"
			Fn_UI_WpfObject_Operations = objReferencePath.Exist(iTimeOut)
			If Fn_UI_WpfObject_Operations Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & " object is exist.")
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "object does not exist.")
			End If
		Case "Enabled"
			If Fn_UI_WpfObject_Operations(sFunctionName, "Exist", objReferencePath, iTimeOut) Then
				If objReferencePath.GetROProperty("enabled") = "1"  OR objReferencePath.GetROProperty("enabled") = True Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & "object is exists and enabled.")
					Fn_UI_WpfObject_Operations = True
				Else
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "object is not enabled.")
				End If
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),sFuncLog & "does not exist.")
			End If
	End Select
End Function
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Function Name		:	Fn_UI_GenWpfObject_Operations

'Description		:	This function to perform operations on Generic Wpf objects

'Parameters		    :	1. sFunctionName: Function name
'						2. sAction		: Action name
'						3. objContainer : Prent UI Component or JavaTable object
'						4. sWpfObject 	: Twistie Control name
'						5. sWpfObjectText	: Twistie Control text
'						6. iX : X coordinate
'						7. iX : Y coordinate
'						8. sMicButton : mouse button name
'
'Return Value		: 	TRUE \ FALSE
'
'Examples			:	Fn_UI_GenWpfObject_Operations("", "Click", Window("SolidWorksDefaultWindow"),"SwimManager", "Open","","","","")
'
'History			:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep Navghane		 	| 10-May-2015	|	1.0			|	Sandeep Navghane 		| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Function Fn_UI_GenWpfObject_Operations(sFunctionName, sAction, objContainer, sWpfObject, sWpfObjectText,iX,iY,sValue,sMicButton)
	Dim objWpfObject, objWpfObjectText
	Dim iLeft,iTop,iRight,iBottom
	Dim sFuncLog
	Dim bFlag
	
	Set objWpfObject = objContainer.WpfObject(sWpfObject)
	
	Fn_UI_GenWpfObject_Operations = False
	
	sFuncLog = sFunctionName + " > Fn_UI_GenWpfObject_Operations : [ " &  objWpfObject.toString & " ] : Action = " & sAction & " : "
	If objWpfObject.Exist(5)= False Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), sFuncLog & " : FAIL : Does not exist.")
		Set objWpfObject = Nothing 
		Exit Function
	End If

	Select Case sAction
		Case "Click"
			If iX<>"" and iY<>"" Then			
				objWpfObject.Click iX,iY
			ElseIf 	iX<>"" and iY<>"" and sMicButton<>"" Then
				objWpfObject.Click iX,iY,sMicButton
			Else
				objWpfObject.Click
			End If	
			If Err.Number < 0 Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  Fail to click on Wpf Object [ " & Cstr(sWpfObjectText) & " ] with Error Number [" & Cstr(Err.Number) & "] Error Description  [" & Cstr(Err.Description) & "]" )
				Fn_UI_GenWpfObject_Operations=False
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<PASS>:  Successfully click on Wpf Object [ " & Cstr(sWpfObjectText) & " ]")
				Fn_UI_GenWpfObject_Operations=True
			End If	
	End Select
	Set objWpfObject = Nothing 
End Function
