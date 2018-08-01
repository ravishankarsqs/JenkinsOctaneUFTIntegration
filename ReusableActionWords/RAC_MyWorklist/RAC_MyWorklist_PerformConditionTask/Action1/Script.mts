'! @Name 			RAC_MyWorklist_PerformConditionTask
'! @Details 		This actionword is used to perform operations on Perform condition task dialog
'! @InputParam1 	sMode				: ViewerTab/PerformConditionTask
'! @InputParam2 	sInstructions 		: Task Instructions
'! @InputParam3 	sProcessDescription : Process Description
'! @InputParam4 	sComments 			: Comments
'! @InputParam5 	sTaskResult 		: task result
'! @InputParam5 	sButton	 			: button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			01 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_PerformConditionTask","RAC_MyWorklist_PerformConditionTask",OneIteration,"PerformConditionTask","","performed task","task completed","Accept","OK"

Option Explicit
Err.Clear

'Declaring variables
Dim sMode,sInstructions,sProcessDescription,sComments,sTaskResult,sButton	
Dim objPerformConditionTask

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sMode = Parameter("sMode")
sInstructions = Parameter("sInstructions")
sProcessDescription = Parameter("sProcessDescription")
sComments = Parameter("sComments")
sTaskResult = Parameter("sTaskResult")
sButton = Parameter("sButton")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_PerformConditionTask"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Selecting mode to perform operation
Select Case sMode	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Viewver tab mode to perform condition task
	Case "ViewerTab"
		Set objPerformConditionTask=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","japt_MyWorkListApplet","")
		'Selecting [ Viewer ] tab
		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",OneIteration,"Select", "Viewer", ""
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_PerformConditionTask"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

		'Set [ Task View ] from Viewer Tab
		Call Fn_UI_Object_Operations("RAC_MyWorklist_PerformConditionTask","settoproperty",objPerformConditionTask.JavaRadioButton("jrdb_ViewerMode"),"","attached text","Task View")
		If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_PerformSignoffDecision", "Set", objPerformConditionTask, "jrdb_ViewerMode", "ON")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform condition task operation as fail to select [ task view ] option from [ Viewer ] tab","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Perform Condition Task dialog mode to perform condition task
	Case "PerformConditionTask"
'		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ActionsPerform"
		Set objPerformConditionTask=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","jdlg_PerformConditionTask","")
		'Checking existance of [ Perform Condition Task ] dialog
		If Fn_UI_Object_Operations("RAC_MyWorklist_PerformConditionTask","Exist",objPerformConditionTask,2,"","") = False Then
			'Calling menu
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ActionsPerform"
		End If
End Select		

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_PerformConditionTask"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Checking existance of [ Perform Condition Task ] dialog
If Fn_UI_Object_Operations("RAC_MyWorklist_PerformConditionTask","Exist",objPerformConditionTask,"","","") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform condition task operation as [ Perform Condition Task ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If
		
'Capturing functionality execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_MyWorklist_PerformConditionTask",sMode,"","")

'Set Task Instructions
If sInstructions<>"" Then
	If Fn_UI_Object_Operations("RAC_MyWorklist_PerformConditionTask","getroproperty",objPerformConditionTask.JavaEdit("jedt_TaskInstructions"),"","editable","")=1 Then
		If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_PerformConditionTask", "Set", objPerformConditionTask, "jedt_TaskInstructions", sInstructions)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform condition task operation as fail to set value of Task Instructions field","","","","","")
			Call Fn_ExitTest()
		End iF
		'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
End If

'Set Process Description
If sProcessDescription<>"" Then
	If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_PerformConditionTask", "Set",  objPerformConditionTask, "jedt_ProcessDescription", sProcessDescription)=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Process Description value while performing perform condition task operation","","","","","")
		Call Fn_ExitTest()
	End If
	'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End If

'Set Comments
If sComments<>"" Then
	If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_PerformConditionTask", "Set",  objPerformConditionTask, "jedt_Comments", sComments)=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Comments while performing perform condition task operation","","","","","")
		Call Fn_ExitTest()
	End IF
	'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End If

'Set task result option
If lcase(Cstr(sTaskResult)) <> "" Then
	If Cstr(sTaskResult) = "Signoff CR" Then
		sTaskResult = "OK to Proceed"
	End If
	Call Fn_UI_Object_Operations("RAC_MyWorklist_PerformConditionTask","settoproperty",objPerformConditionTask.JavaRadioButton("jrdb_TaskResult"),"","attached text",sTaskResult)
	If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_PerformConditionTask", "Set", objPerformConditionTask, "jrdb_TaskResult", "ON")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Task Result ] option [ " & cstr(sTaskResult) & " ] while performing perform condition task operation","","","","","")
		Call Fn_ExitTest()
	End IF	
End If	
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)


'Clicking on button
If sButton<>"" Then
	If Fn_UI_JavaButton_Operations("RAC_MyWorklist_PerformConditionTask","Click",objPerformConditionTask,"jbtn_" & sButton)=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ " & Cstr(sButton) & " ] button while performing perform condition task operation","","","","","")
		Call Fn_ExitTest()
	End If	
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End If

'Checking existance of Warning dialog
If JavaWindow("jwnd_MyWorkListWindow").JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_Warning").Exist(6) Then
	If Fn_UI_JavaButton_Operations("RAC_MyWorklist_PerformConditionTask","Click",JavaWindow("jwnd_MyWorkListWindow").JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_Warning"),"jbtn_OK")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ OK ] button from Warning dialog while performing perform condition task operation","","","","","")
		Call Fn_ExitTest()
	End If	
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	
	If objPerformConditionTask.Exist(3) Then
		objPerformConditionTask.Close
	End If
End If

If Err.Number <> 0 then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform condition task operation due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully performed [ perform condition task ] operation from [" & Cstr(sMode) & "]","","","","DONOTSYNC","")	
End If

'Capturing functionality execution end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MyWorklist_PerformConditionTask",sMode,"","")		

'Releasing objects of [ Perform Condition Task ] dialog
Set objPerformConditionTask=Nothing	

Function Fn_ExitTest()
	'Releasing objects of [ Perform Condition Task ] dialog
	Set objPerformConditionTask=Nothing
	ExitTest
End Function


