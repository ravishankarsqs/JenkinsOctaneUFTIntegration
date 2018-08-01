'! @Name 			RAC_MyWorklist_PerformSignoffDecision
'! @Details 		This Action word to perform operations on Perform Signoff Decision dialog
'! @InputParam1 	sNode 			: Myworklist tree node
'! @InputParam2 	sMode 			: ViewerTab/PerformDoTask
'! @InputParam3 	sAutomationID 	: Automation ID
'! @InputParam4 	sDecision 		: Signoff decision
'! @InputParam5 	sComments 		: Signoff decision Comments
'! @InputParam6 	sButton 		: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com 
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			30 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_PerformSignoffDecision","RAC_MyWorklist_PerformSignoffDecision",OneIteration,"","PerformSignoff","TestUser1DraftingEcOkcAlDocController","Approve","Test comment","Close"

Option Explicit
Err.Clear

'Declaring variables
Dim sNode,sMode,sAutomationID,sDecision,sComments,sButton
Dim objPerformSignoffDialog,objSignoffDecisionDialog
Dim sUserGroupRole
Dim iCounter,iPerformSignoffDecisionCount
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sNode = Parameter("sNode")
sMode = Parameter("sMode")
sAutomationID = Parameter("sAutomationID")
sDecision = Parameter("sDecision")
sComments = Parameter("sComments")
sButton = Parameter("sButton")

'Creating object of [ Signoff Decision ] dialog
Set objSignoffDecisionDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","jdlg_SignoffDecision","")	

sUserGroupRole=""
If sAutomationID<>"" Then
	sUserGroupRole = Fn_Setup_GetTestUserDetailsFromExcelOperations("getrole","",sAutomationID)
	sUserGroupRole ="*-*/" & CStr(sUserGroupRole) & "/*"
End If

If sNode<>"" Then
	'to select specific node from [ My Worklist ] tree
	LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_TreeNodeOperations","RAC_MyWorklist_TreeNodeOperations",OneIteration,"Select",sNode,""
End If

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_PerformSignoffDecision"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Selecting mode to perform operation
Select Case sMode
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "ViewerTab"
		Set objPerformSignoffDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","jwnd_TcDefaultApplet","")
		'Selecting [ Viewer ] tab
		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",OneIteration,"Select", "Viewer", ""
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "SignoffDecisionDialogButtonClick"
		'Clicking on specific button of signoff decision dialog
		If Fn_UI_JavaButton_Operations("RAC_MyWorklist_PerformSignoffDecision","Click",objSignoffDecisionDialog,"jbtn_" & sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of signoff decision dialog","","","","","")
			Call Fn_ExitTest()
		End If	
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		'Releasing object of [ Signoff Decision ] dialog
		Set objSignoffDecisionDialog=Nothing
		ExitAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "PerformSignoff"
		Set objPerformSignoffDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","jdlg_PerformSignoff","")
		'Checking existance of [ Perform Signoff ] dialog
		If Fn_UI_Object_Operations("RAC_MyWorklist_PerformSignoffDecision","Exist",objPerformSignoffDialog,3,"","") = False Then
			'Invoking perform signoff dialog by executing menu
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ActionsPerform"
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
End Select	

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_PerformSignoffDecision"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Fetching datatable current selected row
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Perform Signoff Decision",sMode,"","")

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

DataTable.SetCurrentRow 1
Call Fn_CommonUtil_DataTableOperations("AddColumn","PerformSignoffDecisionCount","","")
iPerformSignoffDecisionCount=Fn_CommonUtil_DataTableOperations("GetValue","PerformSignoffDecisionCount","","")
If iPerformSignoffDecisionCount="" Then
	iPerformSignoffDecisionCount=1
Else
	iPerformSignoffDecisionCount=iPerformSignoffDecisionCount+1
End If
Call Fn_CommonUtil_DataTableOperations("SetValue","PerformSignoffDecisionCount",iPerformSignoffDecisionCount,"")
	
'Seleting decision link
If sUserGroupRole="" Then
	iCounter=0
Else
	bFlag=False
	For iCounter = 0 To Cint(objPerformSignoffDialog.JavaTable("jtbl_SignoffTable").GetROProperty("rows"))-1
		If Trim(sUserGroupRole)=Trim(objPerformSignoffDialog.JavaTable("jtbl_SignoffTable").GetCellData(iCounter,"User-Group/Role")) Then
			bFlag=True
			Exit for				
		End If			
	Next
	If bFlag=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform Signoff Decision operation as signoff table does not contain value [ " & Cstr(sUserGroupRole) & " ] under [ User-Group/Role ] column on [ Perform Signoff ] dialog","","","","","")
		Call Fn_ExitTest()
	End If
End If
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
wait 0,500

Call Fn_Setup_ReporterFilter("DisableAll")
On Error Resume Next
'Clicking on cell to open Signoff decision dialog
'If Fn_UI_JavaTable_Operations("RAC_MyWorklist_PerformSignoffDecision","ClickCell",objPerformSignoffDialog,"jtbl_SignoffTable",iCounter,"Decision","","","")=False Then
'	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform Signoff Decision operation as fail to click on cell of row [ " & Cstr(iCounter) & " ] under [ Decision ] column on [ Perform Signoff ] dialog","","","","","")
'	Call Fn_ExitTest()
'End If
objPerformSignoffDialog.JavaTable("jtbl_SignoffTable").SelectCell Cint(iCounter),"Decision"
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

If objSignoffDecisionDialog.Exist(10)=False Then
	objPerformSignoffDialog.JavaTable("jtbl_SignoffTable").ClickCell Cint(iCounter),"Decision"
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End IF
On Error GoTo 0
Err.Clear
Call Fn_Setup_ReporterFilter("EnableAll")

'Checking existance of [ Signoff Decision ] dialog
If Fn_UI_Object_Operations("RAC_MyWorklist_PerformSignoffDecision","Exist",objSignoffDecisionDialog,"","","") Then

	'Setting Signoff Decision option
	If sDecision<>"" Then
		Call Fn_UI_Object_Operations("RAC_MyWorklist_PerformSignoffDecision","settoproperty",objSignoffDecisionDialog.JavaRadioButton("jrdb_DecisionOption"),"","attached text",sDecision)
		If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_PerformSignoffDecision", "Set", objSignoffDecisionDialog, "jrdb_DecisionOption", "ON")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform Signoff Decision operation as fail to select [ Decision ] option from signoff decision dialog","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
	
	'Setting Signoff Decision Comments
	If sComments<>"" Then
		If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_PerformSignoffDecision", "Set",  objSignoffDecisionDialog, "jedt_Comments", sComments)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform Signoff Decision operation as fail to enter comments in [ Comments ] field of signoff decision dialog","","","","","")
			Call Fn_ExitTest()
		End IF
	End If		
	DataTable.SetCurrentRow iPerformSignoffDecisionCount
	
	Call Fn_CommonUtil_DataTableOperations("AddColumn","PerformSignoffDecisionComment","","")
	DataTable.Value("PerformSignoffDecisionComment","Global") = sComments
		
	'Click on [ ok ] button
	If Fn_UI_JavaButton_Operations("RAC_MyWorklist_PerformSignoffDecision","Click",objSignoffDecisionDialog,"jbtn_OK")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform Signoff Decision operation as fail to click [ OK ] button of signoff decision dialog","","","","","")
		Call Fn_ExitTest()
	End If	
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	
	'Checking existance of Warning dialog
	If JavaWindow("jwnd_MyWorkListWindow").JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_Warning").Exist(6) Then
		If Fn_UI_JavaButton_Operations("RAC_MyWorklist_PerformConditionTask","Click",JavaWindow("jwnd_MyWorkListWindow").JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_Warning"),"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ OK ] button from Warning dialog while performing perform condition task operation","","","","","")
			Call Fn_ExitTest()
		End If	
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
	
	'Clciking on button
	If sButton<>"" Then
		If Fn_UI_JavaButton_Operations("RAC_MyWorklist_PerformSignoffDecision","Click",objPerformSignoffDialog,"jbtn_" & sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform Signoff Decision operation as fail to click [ " & Cstr(sButton) & " ] button of perform signoff dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
End If

'Capture business functionality end time	
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Perform Signoff Decision",sMode,"","")

If Err.Number<0 then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Perform Signoff Decision operation due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Performed Signoff Decision operation from [ " & Cstr(sMode) & " ]","","","","DONOTSYNC","")	
End If

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing all objects
Set objPerformSignoffDialog=Nothing
Set objSignoffDecisionDialog=Nothing

Function Fn_ExitTest()
	'Releasing all objects
	Set objPerformSignoffDialog=Nothing
	Set objSignoffDecisionDialog=Nothing
	ExitTest
End Function	

