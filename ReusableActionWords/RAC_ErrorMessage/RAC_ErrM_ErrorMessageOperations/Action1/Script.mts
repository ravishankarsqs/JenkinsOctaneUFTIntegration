'! @Name 			RAC_ErrM_ErrorMessageOperations
'! @Details 		This Action word used to perform operations on error messgae
'! @InputParam1. 	sErrorAction 			: Error Action
'! @InputParam2. 	dictErrorMessageInfo 	: External dictionary parameter to pass error information
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com 
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			14 Dec 2015
'! @Version 		1.0
'! @Example 		dictErrorMessageInfo("Action")="DetailMessageVerify"
'! @Example 		dictErrorMessageInfo("Perspective")="myteamcenter"
'! @Example 		dictErrorMessageInfo("DialogTitle")="Apple Translators"
'! @Example 		dictErrorMessageInfo("ErrorMessage")="BusinessRulesForHandler"
'! @Example 		dictErrorMessageInfo("ErrorMessageXMLName")="RAC_ErrorMessage_ERM"
'! @Example 		dictErrorMessageInfo("Button")="OK"
'! @Example 		LoadAndRunAction "RAC_ErrorMessage\RAC_ErrM_ErrorMessageOperations","RAC_ErrM_ErrorMessageOperations",OneIteration,"ComonErrorDialog"
'! @Example 		dictErrorMessageInfo("Action")="Basic"
'! @Example 		dictErrorMessageInfo("Perspective")="myteamcenter"
'! @Example 		dictErrorMessageInfo("DialogTitle")="Error"
'! @Example 		dictErrorMessageInfo("ErrorMessage")="InvalidID"
'! @Example 		dictErrorMessageInfo("ErrorMessageXMLName")="RAC_ErrorMessage_ERM"
'! @Example 		dictErrorMessageInfo("Button")="OK"
'! @Example 		LoadAndRunAction "RAC_ErrorMessage\RAC_ErrM_ErrorMessageOperations","RAC_ErrM_ErrorMessageOperations",OneIteration,"ComonErrorWindow"

Option Explicit
Err.Clear

'Declaring varaibles
Dim sErrorAction

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sErrorAction = Parameter("sErrorAction")


Select Case sErrorAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to perform operations on common java error dialogs
	Case "ComonErrorDialog"
		Call Fn_ErrM_ComonErrorDialogOperations(dictErrorMessageInfo("Action"),dictErrorMessageInfo("Perspective"),dictErrorMessageInfo("InvokeOption"),dictErrorMessageInfo("InvokeValue"),dictErrorMessageInfo("InvokeValueXMLName"),dictErrorMessageInfo("DialogTitle"),dictErrorMessageInfo("ErrorMessageTitle"),dictErrorMessageInfo("ErrorMessageXMLName"),dictErrorMessageInfo("Button"))
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to perform operations on common java error dialogs
	Case "ComonErrorWindow"
		Call Fn_ErrM_ComonErrorWindowOperations(dictErrorMessageInfo("Action"),dictErrorMessageInfo("Perspective"),dictErrorMessageInfo("InvokeOption"),dictErrorMessageInfo("InvokeValue"),dictErrorMessageInfo("InvokeValueXMLName"),dictErrorMessageInfo("DialogTitle"),dictErrorMessageInfo("ErrorMessageTitle"),dictErrorMessageInfo("ErrorMessageXMLName"),dictErrorMessageInfo("Button"))
End Select

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Function Name											|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -|- - - - - - - - - - - - - - - -| - - - - - - - |- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. 	Fn_ErrM_ComonErrorDialogOperations						|	sandeep.navghane@sqs.com 	|	12-Jul-2016	|	Function Used to handle common java error dialogs
'002. 	Fn_ErrM_ComonErrorWindowOperations						|	sandeep.navghane@sqs.com 	|	12-Jul-2016	|	Function Used to handle common java error window
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -|- - - - - - - - - - - - - - - -| - - - - - - - |- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_ErrM_ComonErrorDialogOperations
'
'Function Description	 :	Function Used to handle common java error dialogs
'
'Function Parameters	 :  1.sAction	 			: Action to perform					
'							2.sPerspective			: Perspective Name
'							3.sInvokeOption	 		: Error dialog invoke option
'							4.sInvokeValue			: Error Invoke option value
'							5.sInvokeValueXMLName	: Automation XML name
'							6.sDialogTitle			: Error dialog title
'							7.sErrorMessageTitle	: Error message
'							8.sErrorMessageXMLName	: Error message xml name
'							9.sButton				: Button name
'
'Function Return Value	 : 	True Or False
'
'Wrapper Function	     : 	RAC_ErrM_ErrorMessageOperations
'
'Function Pre-requisite	 :	Should be log in Teamcenter or Error message dialog should appear
'
'Function Usage		     :	bReturn=Fn_ErrM_ComonErrorDialogOperations("DetailMessageVerifyOnMoreButton","structuremanager","","","","Save As...","This Item Revision does not contain a BOMView Revision","","OK")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  12-Jul-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_ErrM_ComonErrorDialogOperations(sAction,sPerspective,sInvokeOption,sInvokeValue,sInvokeValueXMLName,sDialogTitle,sErrorMessageTitle,sErrorMessageXMLName,sButton)
	'Declaring variables
	Dim sErrorMessage
	Dim objErrorDialog,objErrorDialog1
	
	'Selecting Error Invoke option
	If sInvokeOption<>"" Then
		Select Case lcase(sInvokeOption)
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "menu"
				LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select",sInvokeValue
	   		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "toolbar"
				'Click on toolbar button to invoke Error Dialog
				LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click", sInvokeValue, "",""
		End Select
	End If
	
	GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_ErrM_ErrorMessageOperations"
	GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

	'Creating object of error dialog
	Select Case sPerspective
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "","myteamcenter"
			Set objErrorDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ErrorMessage_OR","jdlg_ErrorDialog","")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "structuremanager"
			Set objErrorDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ErrorMessage_OR","jdlg_ErrorDialog@2","")
	End Select
	
	Set objErrorDialog1=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ErrorMessage_OR","jdlg_ErrorDialog@3","")
	'Setting Error Dialog title
	If sDialogTitle<>"" Then
		Call Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","SetTOProperty",objErrorDialog,"","title",sDialogTitle)
		Call Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","SetTOProperty",objErrorDialog1,"","title",sDialogTitle)
	End If
	
	'Retrive error message information
	If sErrorMessageXMLName<>"" Then
		sErrorMessage=Fn_FSOUtil_XMLFileOperations("getvalue",sErrorMessageXMLName,sErrorMessageTitle,"")
	End IF
	
	'Checking Existance of Error Dialog
	If Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","Exist",objErrorDialog,"","","") Then
		'Do nothing
	ElseIf Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","Exist",objErrorDialog1,"","","") Then
		Set objErrorDialog=objErrorDialog1		
	Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as error dialog [ " & Cstr(sDialogTitle) & " ] does not exist","","","","","")
		Call Fn_ExitTest()
	End If
	
	Set objErrorDialog1=Nothing
	
	'Retrive error dialog title
	sDialogTitle=objErrorDialog.GetROProperty("title")
	
	'Capture execution start time
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Comon Error Dialog Operations",sAction,"Error Message",sErrorMessage)
	
	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle basic error dialog message
		Case "Basic"			
			If Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","Exist",objErrorDialog.JavaEdit("jedt_ErrorEdit"),"","","") Then
				If InStr(1,trim(objErrorDialog.JavaEdit("jedt_ErrorEdit").GetROProperty("value")),sErrorMessage) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current error messgae match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
					Set objErrorDialog=Nothing
					Call Fn_ExitTest()
				End If
			ElseIf Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","Exist",objErrorDialog.JavaStaticText("jstx_ErrorText"),"","","") Then
				If InStr(1,trim(objErrorDialog.JavaStaticText("jstx_ErrorText").GetROProperty("label")),sErrorMessage) Then				
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current error messgae match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
					Set objErrorDialog=Nothing
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
				Set objErrorDialog=Nothing
				Call Fn_ExitTest()
			End If
			'Clicking on button
			If sButton<>"" Then
				If  Fn_UI_JavaButton_Operations("Fn_ErrM_ComonErrorDialogOperations","Click",objErrorDialog,"jbtn_" & sButton)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) &" ] button of error dialog [ " & Cstr(sDialogTitle) & " ]","","","","","")
					Set objErrorDialog=Nothing
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(1)
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle error message by clicking on  more button
		Case "DetailMessageVerify"
			'Checking Existance of  [ More Error Edit ]
			If Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","Exist",objErrorDialog.JavaEdit("jedt_MoreErrorEdit"),"","","") Then
				If inStr(1,trim(objErrorDialog.JavaEdit("jedt_MoreErrorEdit").GetROProperty("value")),sErrorMessage) Then
					Fn_ErrM_ComonErrorDialogOperations=True
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current error messgae match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
					Set objErrorDialog=Nothing
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as error message edit box does not exist","","","","","")					
				Set objErrorDialog=Nothing
				Call Fn_ExitTest()
			End If

			'Clicking on button
			If sButton<>"" Then
				If  Fn_UI_JavaButton_Operations("Fn_ErrM_ComonErrorDialogOperations","Click",objErrorDialog,"jbtn_" & sButton)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) &" ] button of error dialog [ " & Cstr(sDialogTitle) & " ]","","","","","")
					Set objErrorDialog=Nothing
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(1)
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle error message by clicking on  more button
		Case "DetailMessageVerifyFromErrorButton"
			objErrorDialog.JavaStaticText("jstx_ObjectName").SetTOProperty "label",sButton
			objErrorDialog.JavaStaticText("jstx_ObjectName@1").SetTOProperty "label",sButton
			If Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","Exist",objErrorDialog.JavaButton("jbtn_DefaultError_16"),"","","") Then
				If inStr(1,trim(objErrorDialog.JavaButton("jbtn_DefaultError_16").GetROProperty("tool_tip_text")),sErrorMessage) Then
					Fn_ErrM_ComonErrorDialogOperations=True
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current error messgae match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
					Set objErrorDialog=Nothing
					Call Fn_ExitTest()
				End If
			ElseIf Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","Exist",objErrorDialog.JavaButton("jbtn_DefaultError_16@1"),"","","") Then
				If inStr(1,trim(objErrorDialog.JavaButton("jbtn_DefaultError_16@1").GetROProperty("tool_tip_text")),sErrorMessage) Then
					Fn_ErrM_ComonErrorDialogOperations=True
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current error messgae match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
					Set objErrorDialog=Nothing
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as error message button does not exist","","","","","")					
				Set objErrorDialog=Nothing
				Call Fn_ExitTest()
			End If

			'Clicking on button
			If  Fn_UI_JavaButton_Operations("Fn_ErrM_ComonErrorDialogOperations","Click",objErrorDialog,"jbtn_OK")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) &" ] button of error dialog [ " & Cstr(sDialogTitle) & " ]","","","","","")
				Set objErrorDialog=Nothing
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle error message by clicking on  more button
		Case "DetailMessageVerifyOnMoreButton"
			'Checking Existance of  [ More ] button
			If Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","Exist",objErrorDialog.JavaCheckBox("jckb_More"),"","","")Then
				'Click on More button
				objErrorDialog.JavaCheckBox("jckb_More").Set "ON"
				'Checking Existamce of  [ More_Error_Edit ]
				If Fn_UI_Object_Operations("Fn_ErrM_ComonErrorDialogOperations","Exist",objErrorDialog.JavaEdit("jedt_MoreErrorEdit"),"","","") Then
					If inStr(1,trim(objErrorDialog.JavaEdit("jedt_MoreErrorEdit").GetROProperty("value")),sErrorMessage) Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current error messgae match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
						Set objErrorDialog=Nothing
						Call Fn_ExitTest()
					End If
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as error message edit box does not exist","","","","","")					
					Set objErrorDialog=Nothing
					Call Fn_ExitTest()
				End If
			End If
			'Clicking on button
			If sButton<>"" Then
				If  Fn_UI_JavaButton_Operations("Fn_ErrM_ComonErrorDialogOperations","Click",objErrorDialog,"jbtn_" & sButton)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) &" ] button of error dialog [ " & Cstr(sDialogTitle) & " ]","","","","","")
					Set objErrorDialog=Nothing
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(1)
			End If
	End Select
	
	'Capturing execution end time
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Comon Error Dialog Operations",sAction,"Error Message",sErrorMessage)
	
	'Releasing Error Dialog object
	Set objErrorDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_ErrM_ComonErrorWindowOperations
'
'Function Description	 :	Function Used to handle common java error windows
'
'Function Parameters	 :  1.sAction	 			: Action to perform					
'							2.sPerspective			: Perspective Name
'							3.sInvokeOption	 		: Error dialog invoke option
'							4.sInvokeValue			: Error Invoke option value
'							5.sInvokeValueXMLName	: Automation XML name
'							6.sDialogTitle			: Error dialog title
'							7.sErrorMessageTitle	: Error message
'							8.sErrorMessageXMLName	: Error message xml name
'							9.sButton				: Button name
'
'Function Return Value	 : 	True Or False
'
'Wrapper Function	     : 	RAC_ErrM_ErrorMessageOperations
'
'Function Pre-requisite	 :	Should be log in Teamcenter or Error message window should appear
'
'Function Usage		     :	bReturn=Fn_ErrM_ComonErrorWindowOperations("Basic","structuremanager","","","","Save As...","This Item Revision does not contain a BOMView Revision","","OK")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  12-Jul-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_ErrM_ComonErrorWindowOperations(sAction,sPerspective,sInvokeOption,sInvokeValue,sInvokeValueXMLName,sDialogTitle,sErrorMessageTitle,sErrorMessageXMLName,sButton)
	'Declaring variables
	Dim sErrorMessage
	Dim objErrorWindow,objErrorWindow1
	
	'Selecting Error Invoke option
	If sInvokeOption<>"" Then
		Select Case lcase(sInvokeOption)
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "menu"
				LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select",sInvokeValue
	   		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "toolbar"
				'Click on toolbar button to invoke Error Dialog
				LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click", sInvokeValue, "",""
		End Select
	End If
	
	GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_ErrM_ErrorMessageOperations"
	GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
	
	'Creating object of error dialog
	Select Case sPerspective
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "","myteamcenter"
			Set objErrorWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ErrorMessage_OR","jwnd_ErrorWindow@3","")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "structuremanager"
			Set objErrorWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ErrorMessage_OR","jwnd_ErrorWindow","")
	End Select
	
	'Setting Error Dialog title
	If sDialogTitle<>"" Then
		Call Fn_UI_Object_Operations("Fn_ErrM_ComonErrorWindowOperations","SetTOProperty",objErrorWindow,"","title",sDialogTitle)
	End If
	
	'Retrive error message information
	If sErrorMessageXMLName<>"" Then
		sErrorMessage=Fn_FSOUtil_XMLFileOperations("getvalue",sErrorMessageXMLName,sErrorMessageTitle,"")
	End IF
	
	'Checking Existance of Error Dialog
	If Fn_UI_Object_Operations("Fn_ErrM_ComonErrorWindowOperations","Exist",objErrorWindow,"","","")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as error dialog [ " & Cstr(sDialogTitle) & " ] does not exist","","","","","")
		Call Fn_ExitTest()
	End If
		
	'Retrive error dialog title
	sDialogTitle=objErrorWindow.GetROProperty("title")
	
	'Capture execution start time
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Comon Error Dialog Operations",sAction,"Error Message",sErrorMessage)
	
	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle basic error dialog message
		Case "Basic"			
			If Fn_UI_Object_Operations("Fn_ErrM_ComonErrorWindowOperations","Exist",objErrorWindow.JavaEdit("jedt_ErrorEdit"),"","","") Then
				If InStr(1,trim(objErrorWindow.JavaEdit("jedt_ErrorEdit").GetROProperty("value")),sErrorMessage) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current error messgae match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
					Set objErrorWindow=Nothing
					Call Fn_ExitTest()
				End If
			ElseIf Fn_UI_Object_Operations("Fn_ErrM_ComonErrorWindowOperations","Exist",objErrorWindow.JavaStaticText("jstx_ErrorText"),"","","") Then
				If InStr(1,trim(objErrorWindow.JavaStaticText("jstx_ErrorText").GetROProperty("label")),sErrorMessage) Then				
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current error messgae match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
					Set objErrorWindow=Nothing
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
				Set objErrorWindow=Nothing
				Call Fn_ExitTest()
			End If
			'Clicking on button
			If sButton<>"" Then
				If  Fn_UI_JavaButton_Operations("Fn_ErrM_ComonErrorWindowOperations","Click",objErrorWindow,"jbtn_" & sButton)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) &" ] button of error dialog [ " & Cstr(sDialogTitle) & " ]","","","","","")
					Set objErrorWindow=Nothing
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(1)
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle basic error dialog tree message
		Case "BasicTree"			
			If Fn_UI_Object_Operations("Fn_ErrM_ComonErrorWindowOperations","Exist",objErrorWindow.JavaTree("jtree_ErrorTree"),"","","") Then
				If InStr(1,trim(objErrorWindow.JavaTree("jtree_ErrorTree").GetColumnValue("#0",1)),sErrorMessage) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current error messgae match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
					Set objErrorWindow=Nothing
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification failed as current error messgae does not match with expected error message [ " & Cstr(sErrorMessage) & " ]","","","","","")					
				Set objErrorWindow=Nothing
				Call Fn_ExitTest()
			End If
			'Clicking on button
			If sButton<>"" Then
				If  Fn_UI_JavaButton_Operations("Fn_ErrM_ComonErrorWindowOperations","Click",objErrorWindow,"jbtn_" & sButton)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) &" ] button of error dialog [ " & Cstr(sDialogTitle) & " ]","","","","","")
					Set objErrorWindow=Nothing
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(1)
			End If
	End Select
	
	'Capturing execution end time
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Comon Error Dialog Operations",sAction,"Error Message",sErrorMessage)
	
	'Releasing Error Dialog object
	Set objErrorWindow=Nothing
End Function

Function Fn_ExitTest()
	ExitTest
End Function

