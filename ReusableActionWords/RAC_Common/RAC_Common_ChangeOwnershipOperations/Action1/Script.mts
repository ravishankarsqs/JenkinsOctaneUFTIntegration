'! @Name 			RAC_Common_ChangeOwnershipOperations
'! @Details 		This Action word to perform operations on Change Ownership dialog
'! @InputParam1 	sAction 		: Action to be performed
'! @InputParam2 	sAutomationID 	: automation id
'! @InputParam3 	sInvokeOption 	: Invoke Option (menu or nooption)
'! @InputParam4 	sPerspective 	: Perspective name
'! @InputParam5 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com 
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			06 Mar 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ChangeOwnershipOperations","RAC_Common_ChangeOwnershipOperations",OneIteration,"ChangeOwnership","TestUser4","","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction, sAutomationID,sInvokeOption, sPerspective
Dim objChangeOwnership,objOrganizationSelection 
Dim sNewOwningUser,sNode,sButton
Dim aNode
Dim iCounter

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters value in local variables
sAction = Parameter("sAction")
sAutomationID =  Parameter("sAutomationID")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sButton = Parameter("sButton")

'Get [ User Settings ] object from xml file
Select Case lcase(sPerspective)
	Case "myteamcenter","","my teamcenter"
		Set objChangeOwnership=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ChangeOwnership","")
		Set objOrganizationSelection=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_OrganizationSelection","")
End Select

Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditChangeOwnership"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
End Select
GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ChangeOwnershipOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'checking Change Ownership dialogs existence
If Not Fn_UI_Object_Operations("RAC_Common_ChangeOwnershipOperations","Exist", objChangeOwnership, GBL_DEFAULT_TIMEOUT,"","") Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on Change Ownership dialog as [ Change Ownership ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Change Ownership",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "changeownership"		
		'Fetching New Owning user details
		sNewOwningUser = Fn_Setup_GetTestUserDetailsFromExcelOperations("getorganizationtreeusernodepath","",sAutomationID)
		
		'Click on New Owning User button
		If Fn_UI_JavaButton_Operations("RAC_Common_ChangeOwnershipOperations", "Click", objChangeOwnership,"jbtn_NewOwningUser") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to change ownership of selected object as fail to click on [ New Owning User ] button of Change Ownership dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Selecting New Owning User from Organization chart
		aNode = Split(sNewOwningUser,"~")
		If Ubound(aNode) > 1 Then
			For iCounter = 0 to Ubound(aNode) - 1
				If iCounter = 0 Then
					sNode = aNode(0)
				Else
					sNode = sNode & "~" & aNode(iCounter)
				End If
				'expanding node
				Call Fn_UI_JavaTree_Operations("RAC_Common_ChangeOwnershipOperations", "Expand",objOrganizationSelection, "jtree_Organization",sNode,"","")
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)					
			Next
		End If
			
		If Fn_UI_JavaTree_Operations("RAC_Common_ChangeOwnershipOperations", "Select",objOrganizationSelection, "jtree_Organization",sNewOwningUser,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Change Ownership of selected object as fail to select organization tree node [ " & Cstr(sNewOwningUser) & " ] on Change Ownership dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		Call Fn_RAC_ReadyStatusSync(GBL_MAX_SYNC_ITERATIONS)
		
		For iCounter = 1 To 10 Step 1
			If Fn_UI_Object_Operations("RAC_Common_ChangeOwnershipOperations", "Enabled", objOrganizationSelection.JavaButton("jbtn_OK"),"", "", "") = False Then
				wait 1
			End If
		Next
		
		'Click on OK button
		If Fn_UI_JavaButton_Operations("RAC_Common_ChangeOwnershipOperations", "Click", objOrganizationSelection,"jbtn_OK") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to change ownership of selected object as fail to click on [ OK ] button of Organization Selection dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on Yes button
		If Fn_UI_JavaButton_Operations("RAC_Common_ChangeOwnershipOperations", "Click", objChangeOwnership,"jbtn_Yes") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to change ownership of selected object as fail to click on [ Yes ] button of Change Ownership dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Change Ownership",sAction,"","")
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully change ownership of selected objects to new user [ " & Cstr(sNewOwningUser) & " ]","","","","","")
End Select

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on Change Ownership dialog due to error number [" & Cstr(Err.Number) & "] and error description  [" & Cstr(Err.Description) & "]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing object of Change Ownership dialog
Set objOrganizationSelection = Nothing
Set objChangeOwnership = Nothing
	
Function Fn_ExitTest()
	'Releasing object of Change Ownership dialog
	Set objOrganizationSelection = Nothing
	Set objChangeOwnership = Nothing
	ExitTest
End Function
