'! @Name 			RAC_MyWorklist_UpdateWeightOperations
'! @Details 		This actionword is used to perform operations on Update Weight dialog
'! @InputParam1 	sAction					: Action to perform
'! @InputParam2 	sMode					: ViewerTab/PerformConditionTask
'! @InputParam3 	sCustomDescription		: Update weight Custom Description
'! @InputParam4 	sEnggNetWeightUpdate	: Engg Net Weight Update
'! @InputParam5 	sMaterialClassification : Material Classification
'! @InputParam6 	sTargetWeight			: Target Weight
'! @InputParam7 	sProcessDescription 	: Process Description
'! @InputParam8 	sComments 				: Comments
'! @InputParam9 	sTaskResult 			: task result
'! @InputParam10 	sButton	 				: button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			12 Jul 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_UpdateWeightOperations","RAC_MyWorklist_UpdateWeightOperations",OneIteration,"Update","ViewerTab","","40.00","","","","","","Save"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sMode,sCustomDescription,sEnggNetWeightUpdate,sMaterialClassification,sTargetWeight,sInstructions,sProcessDescription,sComments,sTaskResult,sButton	
Dim objPerformConditionTask

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sMode = Parameter("sMode")
sCustomDescription = Parameter("sCustomDescription")
sEnggNetWeightUpdate = Parameter("sEnggNetWeightUpdate")
sMaterialClassification = Parameter("sMaterialClassification")
sTargetWeight = Parameter("sTargetWeight")
sProcessDescription = Parameter("sProcessDescription")
sComments = Parameter("sComments")
sTaskResult = Parameter("sTaskResult")
sButton = Parameter("sButton")

'Selecting mode to perform operation
Select Case sMode	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Viewver tab mode to perform Update Weight
	Case "ViewerTab"
		Set objPerformConditionTask=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","japt_MyWorkListApplet","")
		'Selecting [ Viewer ] tab
		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",OneIteration,"Select", "Viewer", ""
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_UpdateWeightOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

		'Set [ Task View ] from Viewer Tab
		Call Fn_UI_Object_Operations("RAC_MyWorklist_UpdateWeightOperations","settoproperty",objPerformConditionTask.JavaRadioButton("jrdb_ViewerMode"),"","attached text","Task View")
		If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_PerformSignoffDecision", "Set", objPerformConditionTask, "jrdb_ViewerMode", "ON")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform Update Weight operation as fail to select [ task view ] option from [ Viewer ] tab","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End Select		

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_UpdateWeightOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of [ Perform Update Weight ] dialog
If Fn_UI_Object_Operations("RAC_MyWorklist_UpdateWeightOperations","Exist",objPerformConditionTask,"","","") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform Update Weight operation as [ Update Weight ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If
		
'Capturing functionality execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_MyWorklist_UpdateWeightOperations",sAction,"","")
Select Case sAction
	Case "Update"
		If sCustomDescription<>"" Then
			Call Fn_UI_Object_Operations("RAC_MyWorklist_UpdateWeightOperations","settoproperty",objPerformConditionTask.JavaStaticText("jstx_ViewerTabText"),"","label","Custom Description:")
			If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_UpdateWeightOperations", "Set", objPerformConditionTask, "jedt_ViewerTabEdit",sCustomDescription)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Custom Description value while performing Update Weight operation","","","","","")
				Call Fn_ExitTest()				
			End If
		End IF
		
		If sEnggNetWeightUpdate<>"" Then
			Call Fn_UI_Object_Operations("RAC_MyWorklist_UpdateWeightOperations","settoproperty",objPerformConditionTask.JavaStaticText("jstx_ViewerTabText"),"","label","Engg Net Weight Update:")
			If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_UpdateWeightOperations", "Set", objPerformConditionTask, "jedt_ViewerTabEdit",sEnggNetWeightUpdate)=False  Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Engg Net Weight Update value while performing Update Weight operation","","","","","")
				Call Fn_ExitTest()				
			End If
		End IF
		
		If sMaterialClassification<>"" Then
			Call Fn_UI_Object_Operations("RAC_MyWorklist_UpdateWeightOperations","settoproperty",objPerformConditionTask.JavaStaticText("jstx_ViewerTabText"),"","label","Material Classification:")
			If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_UpdateWeightOperations", "Set", objPerformConditionTask, "jedt_ViewerTabEdit",sMaterialClassification)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Material Classification value while performing Update Weight operation","","","","","")
				Call Fn_ExitTest()				
			End If
		End IF
		
		If sTargetWeight<>"" Then
			Call Fn_UI_Object_Operations("RAC_MyWorklist_UpdateWeightOperations","settoproperty",objPerformConditionTask.JavaStaticText("jstx_ViewerTabText"),"","label","Target Weight (g):")
			If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_UpdateWeightOperations", "Set", objPerformConditionTask, "jedt_ViewerTabEdit",sTargetWeight)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Target Weight (g) value while performing Update Weight operation","","","","","")
				Call Fn_ExitTest()				
			End If
		End IF
		
		'Set Task Instructions
		If sInstructions<>"" Then
			If Fn_UI_Object_Operations("RAC_MyWorklist_UpdateWeightOperations","getroproperty",objPerformConditionTask.JavaEdit("jedt_TaskInstructions"),"","editable","")=1 Then
				If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_UpdateWeightOperations", "Set", objPerformConditionTask, "jedt_TaskInstructions", sInstructions)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail Update Weight operation as fail to set value of Task Instructions field","","","","","")
					Call Fn_ExitTest()
				End iF
				'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End If
		End If

		'Set Process Description
		If sProcessDescription<>"" Then
			If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_UpdateWeightOperations", "Set",  objPerformConditionTask, "jedt_ProcessDescription", sProcessDescription)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Process Description value while performing Update Weight operation","","","","","")
				Call Fn_ExitTest()
			End If
			'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If

		'Set Comments
		If sComments<>"" Then
			If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_UpdateWeightOperations", "Set",  objPerformConditionTask, "jedt_Comments", sComments)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Comments while performing Update Weight operation","","","","","")
				Call Fn_ExitTest()
			End IF
			'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If

		'Set task result option
		If lcase(Cstr(sTaskResult)) <> "" Then
			Call Fn_UI_Object_Operations("RAC_MyWorklist_UpdateWeightOperations","settoproperty",objPerformConditionTask.JavaRadioButton("jrdb_Complete"),"","attached text",sTaskResult)
			If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_UpdateWeightOperations", "Set", objPerformConditionTask, "jrdb_Complete", "ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ " & Cstr(sTaskResult) & " ] option [ " & cstr(sTaskResult) & " ] while performing Update Weight operation","","","","","")
				Call Fn_ExitTest()
			End IF	
		End If	
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

		'Clicking on button
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_MyWorklist_UpdateWeightOperations","Click",objPerformConditionTask,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ " & Cstr(sButton) & " ] button while performing Update Weight operation","","","","","")
				Call Fn_ExitTest()
			End If	
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
				
		If Err.Number <> 0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform Update Weight operation due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully performed [ Update Weight ] operation from [" & Cstr(sMode) & "]","","","","DONOTSYNC","")	
		End If

		'Capturing functionality execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MyWorklist_UpdateWeightOperations",sAction,"","")
End Select
'Releasing objects of [ Perform Update Weight ] dialog
Set objPerformConditionTask=Nothing	

Function Fn_ExitTest()
	'Releasing objects of [ Perform Update Weight ] dialog
	Set objPerformConditionTask=Nothing
	ExitTest
End Function

