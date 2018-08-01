'! @Name 			RAC_PSE_BaselineOperations
'! @Details 		Action word to remove Design From Product In BOM Table
'! @InputParam1 	sAction 		: Action to be performed e.g. VerifyBaselineError
'! @InputParam2 	sInvokeOption 	: Method to invoke Remove dialog e.g. menu
'! @InputParam3 	sNodePath 		: Table node path
'! @InputParam5 	sButton		 	: Button Name
'! @Author 			Kundan Kudale kundan.kudale@sqs.com
'! @Reviewer 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Date 			15 June 2017
'! @Version 		1.0
'! @Example 		dictBaselineInfo("CloseErrorMessage") = True
'! @Example 		dictBaselineInfo("ErrorMessage") = "The object 4502035/AA is Checked-Out"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_PSE_BaselineOperations","RAC_PSE_BaselineOperations",OneIteration,"VerifyBaselineError","menu","0000313/AA-Asm Cockpit~0000316/AA-Asm Cockpit","OK"
'! @Example 		dictBaselineInfo.RemoveAll
'! @Example 		LoadAndRunAction "RAC_Common\RAC_PSE_BaselineOperations","RAC_PSE_BaselineOperations",OneIteration,"clickbutton","nooption","","OK"
'! @Example 		dictBaselineInfo("CloseErrorMessage") = True
'! @Example 		dictBaselineInfo("ErrorMessage") = "The object 4502035/AA is Checked-Out"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_PSE_BaselineOperations","RAC_PSE_BaselineOperations",OneIteration,"CreateBaseline","menu","0000313/AA-Asm Cockpit~0000316/AA-Asm Cockpit","OK"

Option Explicit
Err.Clear

'Declaring varaibles	
Dim sAction,sInvokeOption,sNodePath,sButton
Dim objBaseline
Dim dictItems, dictKeys
Dim iCounter

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sNodePath = Parameter("sNodePath")
sButton= Parameter("sButton")

'Selecting node from table
If sNodePath<>"" Then
	LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "Select",sNodePath,"","",""
End If

'inoke Remove dialog
Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsBaseline"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'Use this invoke option when user wants to invoke Remove dialog from outside function
End Select

'Creating object of [ Baseline ] dialog
Set objBaseline=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_Baseline","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BaselineOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of Baseline dialog
If Fn_UI_Object_Operations("RAC_PSE_BaselineOperations", "Exist", objBaseline, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Baseline ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","PSE Baseline Operations",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to create baseline
	Case "createbaseline"
	
		'If user hasn't passed any baseline template value then set it to default value
		If Not(dictBaselineInfo.Exists("Baseline Template")) Then
			dictBaselineInfo("Baseline Template") = ""
		End If

		'Taking Items & Keys from dictionary
		dictItems = dictBaselineInfo.Items
		dictKeys = dictBaselineInfo.Keys
		
		For iCounter=0 to dictSummaryTabInfo.count-1
			Select Case dictKeys(iCounter)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Description", "Job Description"
					
					'Set TO property of java edit and verify its existence
					If Fn_UI_Object_Operations("RAC_PSE_BaselineOperations","settoexistcheck",objBaseline.JavaEdit("jedt_BaselineEdit"),"","attached text",dictKeys(iCounter) & ":")=False Then
						If Fn_UI_Object_Operations("RAC_PSE_BaselineOperations","settoexistcheck",objBaseline.JavaEdit("jedt_BaselineEdit"),"","attached text",dictKeys(iCounter))=False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail : Failed to perform action [" & sAction & "] on Baseline dialog as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
							Call Fn_ExitTest()
						End If
					End IF
					
					'Set the value of java edit
					If Fn_UI_JavaEdit_Operations("RAC_PSE_BaselineOperations", "Set",  objBaseline, "jedt_BaselineEdit", dictItems(iCounter) ) = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Baseline ] dialog as failed to set value of edit box [" & dictKeys(iCounter) & "] as [" & dictItems(iCounter) & "]","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Baseline Template"
					
					'Select value for baseline template drop down
					If Fn_UI_JavaList_Operations("RAC_PSE_BaselineOperations", "Select", objBaseline, "jlst_BaselineTemplate", dictItems(iCounter), "", "") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Baseline ] dialog as failed to set value of java list [" & dictKeys(iCounter) & "] as [" & dictItems(iCounter) & "]","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Open On Create", "Dry Run Creation", "Precise Baseline"
				
					'Set TO property of java edit and verify its existence
					If Fn_UI_Object_Operations("RAC_PSE_BaselineOperations","settoexistcheck",objBaseline.JavaCheckBox("jchk_BaselineCheckbox"),"","attached text",dictKeys(iCounter))=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail : Failed to perform action [" & sAction & "] on Baseline dialog as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End IF
					
					'Set the checkbox value
					If Fn_UI_JavaCheckBox_Operations("RAC_PSE_BaselineOperations", "set", objBaseline, jchk_BaselineCheckbox, dictItems(iCounter)) = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Baseline ] dialog as failed to set value of java checkbox [" & dictKeys(iCounter) & "] as [" & dictItems(iCounter) & "]","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				
			End Select
			
		Next
			
		'Click on button
		If sButton <> "" Then
			'Set TO property of java edit and verify its existence
			If Fn_UI_Object_Operations("RAC_PSE_BaselineOperations","settoexistcheck",objBaseline.JavaButton("jtbn_BaselineButton"),"","label",dictKeys(iCounter))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail : Failed to perform action [" & sAction & "] on Baseline dialog as [ " & Cstr(dictKeys(iCounter)) & " ] button does not exist\available on summary tab","","","","","")
				Call Fn_ExitTest()
			End IF
			
			'Click on java button
			If Fn_UI_JavaButton_Operations("RAC_PSE_BaselineOperations", "Click", objBaseline,"jtbn_BaselineButton") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Baseline ] dialog as failed to click on java button [" & dictKeys(iCounter) & "]","","","","","")
				Call Fn_ExitTest()
			End If			
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify error message displayed while creating baseline
	Case "verifybaselineerror"
		
		'Call actionword to create baseline
		LoadAndRunAction "RAC_Common\RAC_PSE_BaselineOperations","RAC_PSE_BaselineOperations",OneIteration,"createbaseline","nooption","",sButton
		
		If Lcase(sAction) = "verifybaselineerror" Then
			'Verify existence of error dialog
			If Fn_UI_Object_Operations("RAC_PSE_BaselineOperations","Exist", JavaDialog("jdlg_BaselineError"),"5", "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail : Failed to perform action [" & sAction & "] on Baseline dialog as error dialog is not displayed.","","","","","")
				Call Fn_ExitTest()
			End IF
			
			'verify error message
			If Instr(Fn_UI_JavaEdit_Operations("RAC_PSE_BaselineOperations", "gettext", JavaDialog("jdlg_BaselineError"),"jedt_ErrorEdit", "" ), dictBaselineInfo("ErrorMessage")) > 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as error message displayed on performing baseline operation contacins text [ " & Cstr(dictBaselineInfo("ErrorMessage")) & " ]","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as text [ " & Cstr(dictBaselineInfo("ErrorMessage")) & " ] is not displayed in error message on performing baseline operation","","","","","")
				Call Fn_ExitTest()
			End If
			
			'Close the error dialog if required
			If dictBaselineInfo("CloseErrorMessage") = True Then
				If Fn_UI_JavaButton_Operations("RAC_PSE_BaselineOperations", "Click", JavaDialog("jdlg_BaselineError"),"jbtn_Close") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Baseline ] dialog as failed to click on close button of baseline error dialog.","","","","","")
					Call Fn_ExitTest()
				End If			
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End If	
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify error message displayed while creating baseline
	Case "clickbutton"
	
		'Click on button
		If sButton <> "" Then
			'Set TO property of java edit and verify its existence
			If Fn_UI_Object_Operations("RAC_PSE_BaselineOperations","settoexistcheck",objBaseline.JavaButton("jtbn_BaselineButton"),"","label",dictKeys(iCounter))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail : Failed to perform action [" & sAction & "] on Baseline dialog as [ " & Cstr(dictKeys(iCounter)) & " ] button does not exist\available on summary tab","","","","","")
				Call Fn_ExitTest()
			End IF
			
			'Click on java button
			If Fn_UI_JavaButton_Operations("RAC_PSE_BaselineOperations", "Click", objBaseline,"jtbn_BaselineButton") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Baseline ] dialog as failed to click on java button [" & dictKeys(iCounter) & "]","","","","","")
				Call Fn_ExitTest()
			End If			
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		
End Select

Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Remove design from product Operations",sAction,"","")
If Err.Number<>0 then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to action [" & sAction & "] on baseline dialog as following runtime error was encountered: [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing object
Set objBaseline=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objBaseline=Nothing
	ExitTest
End Function

