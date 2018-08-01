'! @Name 			RAC_Common_InnerTabOperations
'! @Details 		This actionword is used to perform operations on inner tabs
'! @InputParam1 	sAction			: Action to be performed
'! @InputParam2		sParentTabName	: Parent Tab name
'! @InputParam3		sTabName 		: Inner Tab name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			29 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Summary","Overview"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sParentTabName,sTabName
Dim objDefaultWindow

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sParentTabName = Parameter("sParentTabName")
sTabName = Parameter("sTabName")

'Selecting parent tab	
If sParentTabName<>"" Then
	LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select",sParentTabName,""
End If

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_InnerTabOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Capture business functionality start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Teamcenter Inner Tab Operations",sAction,"Tab name",sTabName)

'Creating Object of Teamcenter main window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_DefaultWindow","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to activate Inner tabs
 	Case "Activate"
		If Fn_UI_JavaTab_Operations("RAC_Common_InnerTabOperations", "Select",objDefaultWindow,"jtab_InnerTab",sTabName)=False Then
'		If Fn_UI_JavaTab_Operations("RAC_Common_InnerTabOperations", "Select",objDefaultWindow.JavaTab("jtab_InnerTab"),"",sTabName)=False Then
		
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to activate tab [ " & Cstr(sTabName) & " ] tab","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully activated [ " & Cstr(sTabName) & " ] tab","","","","","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
	'Case to verify specific tab is in active state
	Case "VerifyActivate"
		If Trim(sTabName)=Trim(Fn_UI_Object_Operations("RAC_Common_InnerTabOperations","getroproperty",objDefaultWindow.JavaTab("jtab_InnerTab"),"","value","")) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sTabName) & " ] tab is currently activated\selected","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sTabName) & " ] tab is currently not activated\selected","","","","","")
			Call Fn_ExitTest()
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
	'Case to Verify Inner tabs not exist
 	Case "VerifyTabNonExist"
		If Fn_UI_JavaTab_Operations("RAC_Common_InnerTabOperations", "Exist",objDefaultWindow,"jtab_InnerTab",sTabName)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sTabName) & " ] tab does not exist\available","","","","DONOTSYNC","")
		Else
		    Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sTabName) & " ] tab is exist\available","","","","","")
			Call Fn_ExitTest()
		End If
End Select

'Capture business functionality end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Teamcenter Inner Tab Operations",sAction,"Tab name",sTabName)

'validating error number
If Err.Number<>0 then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] and error description [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing Teamcenter main window object
Set objDefaultWindow=Nothing

Function Fn_ExitTest()
	'Releasing Teamcenter main window object
	Set objDefaultWindow=Nothing
	ExitTest
End Function

