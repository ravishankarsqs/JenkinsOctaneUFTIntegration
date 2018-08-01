'! @Name 			RAC_PSE_DataPanelTabOperations
'! @Details 		This actionword is used to perform operations on Data panel tabs
'! @InputParam1 	sAction			: Action to be performed
'! @InputParam2		sTabName 		: Inner Tab name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			29 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_PSE\RAC_PSE_DataPanelTabOperations","RAC_PSE_DataPanelTabOperations",OneIteration,"Activate","Attachment"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sTabName
Dim objDataPanel

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sTabName = Parameter("sTabName")

'Creating Object of Teamcenter main window
Set objDataPanel=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jtab_DataPanel","")

If objDataPanel.Exist(1)=False Then
	LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","ShowHideDataPanel","",""
End If

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_DataPanelTabOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Capture business functionality start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Data Panel Tab Operations",sAction,"Tab name",sTabName)

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to activate Inner tabs
 	Case "Activate"
		If Fn_UI_JavaTab_Operations("RAC_PSE_DataPanelTabOperations", "Select",objDataPanel,"",sTabName)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to activate tab [ " & Cstr(sTabName) & " ] tab","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully activated [ " & Cstr(sTabName) & " ] tab","","","","","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to activate Inner tabs
 	Case "Select"
		If Fn_UI_JavaTab_Operations("RAC_PSE_DataPanelTabOperations", "Select",objDataPanel,"",sTabName)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to activate tab [ " & Cstr(sTabName) & " ] tab","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_UI_JavaTab_Operations("RAC_PSE_DataPanelTabOperations", "click",objDataPanel,"",sTabName)
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully activated [ " & Cstr(sTabName) & " ] tab","","","","","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
	'Case to verify specific tab is in active state
	Case "VerifyActivate"
		If Trim(sTabName)=Trim(Fn_UI_Object_Operations("RAC_PSE_DataPanelTabOperations","getroproperty",objDataPanel,"","value","")) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sTabName) & " ] tab is currently activated\selected","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sTabName) & " ] tab is currently not activated\selected","","","","","")
			Call Fn_ExitTest()
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
	'Case to Verify Inner tabs not exist
 	Case "VerifyTabNonExist"
		If Fn_UI_JavaTab_Operations("RAC_PSE_DataPanelTabOperations", "Exist",objDataPanel,"",sTabName)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sTabName) & " ] tab does not exist\available","","","","DONOTSYNC","")
		Else
		    Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sTabName) & " ] tab is exist\available","","","","","")
			Call Fn_ExitTest()
		End If
End Select

'Capture business functionality end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Data Panel Tab Operations",sAction,"Tab name",sTabName)

'validating error number
If Err.Number<>0 then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] and error description [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing Teamcenter main window object
Set objDataPanel=Nothing

Function Fn_ExitTest()
	'Releasing Teamcenter main window object
	Set objDataPanel=Nothing
	ExitTest
End Function


