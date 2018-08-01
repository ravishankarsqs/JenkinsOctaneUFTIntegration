'! @Name 			RAC_MyWorklist_PerformDoTask
'! @Details 		This actionword is used to perform operations on Perform Do task dialog
'! @InputParam1 	sMode				: ViewerTab/PerformConditionTask
'! @InputParam2 	sInstructions 		: Task Instructions
'! @InputParam3 	sProcessDescription : Process Description
'! @InputParam4 	sComments 			: Comments
'! @InputParam5 	bComplete	 		: Complete task option
'! @InputParam5 	sButton	 			: button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			01 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_PerformDoTask","Action1",OneIteration,"PerformDoTask","","performed task","task completed","True","OK"

Option Explicit
Err.Clear

'Declaring variables
Dim sMode,sInstructions,sProcessDescription,sComments,bComplete,sButton	
Dim objPerformDoTask,objPerformConditionTask

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sMode = Parameter("sMode")
sInstructions = Parameter("sInstructions")
sProcessDescription = Parameter("sProcessDescription")
sComments = Parameter("sComments")
bComplete = Parameter("bComplete")
sButton = Parameter("sButton")

'Selecting mode to perform operation
Select Case sMode	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Viewver tab mode to perform condition task
	Case "ViewerTab"
		Set objPerformDoTask=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","japt_MyWorkListApplet","")
		'Selecting [ Viewer ] tab
		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",OneIteration,"Select","Viewer",""
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_PerformDoTask"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

		'Set [ Task View ] from Viewer Tab
		Call Fn_UI_Object_Operations("RAC_MyWorklist_PerformDoTask","settoproperty",objPerformDoTask.JavaRadioButton("jrdb_ViewerMode"),"","attached text","Task View")
		If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_PerformDoTask", "Set", objPerformDoTask, "jrdb_ViewerMode", "ON")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform do task operation as fail to select [ task view ] option from [ Viewer ] tab","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Perform Do Task dialog mode to perform do task
	Case "PerformDoTask"
		Set objPerformDoTask=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","jdlg_PerformDoTask","")
		'Checking existance of [ Perform Do Task ] dialog
		If Fn_UI_Object_Operations("RAC_MyWorklist_PerformDoTask","Exist",objPerformDoTask,2,"","") = False Then
			'calling Toolbar menu
			LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click", "ActionsPerform","",""
		End If
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_PerformDoTask"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Checking existance of [ Perform Do Task ] dialog
If Fn_UI_Object_Operations("RAC_MyWorklist_PerformDoTask","Exist",objPerformDoTask,"","","") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform do task operation as [ Perform Do Task ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If
		
'Capturing functionality execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_MyWorklist_PerformDoTask",sMode,"","")

'Set Task Instructions
If sInstructions<>"" Then
	If Fn_UI_Object_Operations("RAC_MyWorklist_PerformDoTask","getroproperty",objPerformDoTask.JavaEdit("jedt_TaskInstructions"),"","editable","")=1 Then
		If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_PerformDoTask", "Set", objPerformDoTask, "jedt_TaskInstructions", sInstructions)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform do task operation as fail to set value of Task Instructions field","","","","","")
			Call Fn_ExitTest()
		End iF
		'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
End If

'Set Process Description
If sProcessDescription<>"" Then
	If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_PerformDoTask", "Set",  objPerformDoTask, "jedt_ProcessDescription", sProcessDescription)=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Process Description value while performing perform do task operation","","","","","")
		Call Fn_ExitTest()
	End If
	'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End If

'Set Comments
If sComments<>"" Then
	If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_PerformDoTask", "Set",  objPerformDoTask, "jedt_Comments", sComments)=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Comments while performing perform do task operation","","","","","")
		Call Fn_ExitTest()
	End IF
	'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End If

'Set Complete radio button
If lcase(Cstr(bComplete)) = "true" Then
	bComplete="ON"
Else
	bComplete="OFF"
End If

If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_PerformDoTask", "Set", objPerformDoTask, "jrdb_Complete",bComplete)=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select option [ complete ] while performing perform do task operation","","","","","")
	Set objPerformDoTask=Nothing
	Call Fn_ExitTest()
End IF	
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'Clicking on button
If sButton<>"" Then
	If Fn_UI_JavaButton_Operations("RAC_MyWorklist_PerformDoTask","Click",objPerformDoTask,"jbtn_" & sButton)=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ " & Cstr(sButton) & " ] button while performing perform do task operation","","","","","")
		Call Fn_ExitTest()
	End If	
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End If

If Err.Number <> 0 then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform do task operation due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully performed [ perform do task ] operation from [" & Cstr(sMode) & "]","","","","DONOTSYNC","")	
End If

'Capturing functionality execution end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MyWorklist_PerformDoTask",sMode,"","")

'Releasing objects of [ Perform DO Task ] dialog	
Set objPerformDoTask=Nothing	

Function Fn_ExitTest()
	'Releasing objects of [ Perform Do Task ] dialog
	Set objPerformDoTask=Nothing
	ExitTest
End Function

