'! @Name 			RAC_Project_DefinitionOperations
'! @Details 		To perform operations on Project Manager definition tab tasks
'! @InputParam1 	sAction 		: String to indicate what action is to be performed on navigation tree e.g. Select, Expand
'! @InputParam2 	sAutomationID	: Automation ID of the user which will be a node in member selection tree
'! @InputParam3 	sButtonName		: Button name to be clicked after operation is completed
'! @Author 			Kundan Kudale kundan.kudale@sqs.com
'! @Reviewer 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Date 			25 Nov 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Project\RAC_Project_DefinitionOperations", "RAC_Project_DefinitionOperations", oneIteration, "AddMemberToProject", "TestUserEngineeringDesigner1", "Modify"
'! @Example 		LoadAndRunAction "RAC_Project\RAC_Project_DefinitionOperations", "RAC_Project_DefinitionOperations", oneIteration, "VerifyNodeInSelectedMembersList", "TestUserEngineeringDesigner1", ""
'! @Example 		LoadAndRunAction "RAC_Project\RAC_Project_DefinitionOperations", "RAC_Project_DefinitionOperations", oneIteration, "RemoveSelectedMember", "TestUserEngineeringDesigner1", "Modify"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sAutomationID, sMemberDetails, sButtonName
Dim objMemberSelectionTree, objProjectWindow, objSelectedMembersList
Dim aMemberDetails, aAutomationID
Dim iCounter
Dim bPriviledged,bProjectTeamAdministrator

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get parameter values in local variables
sAction = Parameter("sAction")
sAutomationID = Parameter("sAutomationID")
sButtonName = Parameter("sButtonName")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Project_DefinitionOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'creating object of [ Navigation Tree ]
Set objMemberSelectionTree = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jtree_MemberSelection","")

'creating object of [ Project Window ]
Set objProjectWindow = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jwnd_ProjectDefaultWindow","")

'Get object of selected members tree from XML
Set objSelectedMembersList = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jtree_SelectedMembers","")

Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Project Manager member selection tree Node Operations",sAction,"Node name",sAutomationID)

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER = DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select node from navigation tree
	Case "AddMemberToProject"
		
		'Loop to add all users passed by user
		aAutomationID = Split(sAutomationID, "~")
		For iCounter = 0 To Ubound(aAutomationID) Step 1
			
			bPriviledged = False
			bProjectTeamAdministrator = False
			'Check if current user is to be set as priviledged user
			If Instr(Lcase(aAutomationID(iCounter)),":privileged") > 0 Then
				sAutomationID = Trim(Replace(aAutomationID(iCounter), ":privileged", ""))
				bPriviledged = True
			ElseIf Instr(Lcase(aAutomationID(iCounter)),":projectteamadministrator") > 0 Then
				sAutomationID = Trim(Replace(aAutomationID(iCounter), ":projectteamadministrator", ""))
				bProjectTeamAdministrator = True
			Else
				sAutomationID = Trim(aAutomationID(iCounter))
			End If
		
			'Checking existance of [ Member Selection ] Tree
			If Fn_UI_Object_Operations("RAC_Project_DefinitionOperations","Exist", objMemberSelectionTree,"","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on member selection tree as [ Member Selection Tree ] does not exist","","","","","")
				Call Fn_ExitTest()
			End If
			
			'Select the node on which operation is to be performed
			sAutomationID = Fn_Setup_GetTestUserDetailsFromExcelOperations("getusernodetreepathforprojectselectionmembers","",sAutomationID)
			If Fn_RAC_ProjectMemberSelectionTreeOperations("Select", sAutomationID, "", objMemberSelectionTree) = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sAutomationID) & " ] as failed to select the node in member selection tree.","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
			
			'Click on Add button
			If Fn_UI_JavaButton_Operations("RAC_Project_DefinitionOperations", "Click", objProjectWindow, "jbtn_AddUser") = False Then	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [Add] button after selecting node in member selection tree.","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added member [" & sAutomationID & "] to project.","","","","","")
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
			
			'Add the user as priviledged user if specified
			If bPriviledged Then
			
				'Select the user in selected members tree list
				sAutomationID = Trim(Replace(aAutomationID(iCounter), ":privileged", ""))
				sAutomationID = Fn_Setup_GetTestUserDetailsFromExcelOperations("getusernodetreepathforprojectselectedmembers","",sAutomationID)
				If Fn_RAC_ProjectMemberSelectionTreeOperations("Select", sAutomationID, "", objSelectedMembersList) = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sAutomationID) & " ] as failed to select the node in selected members tree.","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
				
				'Right click on it
				objSelectedMembersList.OpenContextMenu sAutomationID
				
				'Select "Set Priviledged Users" menu 
				objProjectWindow.WinMenu("wmnu_ContextMenu").Select "Set Privileged Users"

				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
				
				If Err.Number <> 0 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to set member [" & sAutomationID & "] as privileged user.","","","","","")
					Call Fn_ExitTest()
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully set member [" & sAutomationID & "] as privileged user.","","","","","")
				End If
			ElseIf bProjectTeamAdministrator Then
			
				'Select the user in selected members tree list
				sAutomationID = Trim(Replace(aAutomationID(iCounter), ":projectteamadministrator", ""))
				sAutomationID = Fn_Setup_GetTestUserDetailsFromExcelOperations("getusernodetreepathforprojectselectedmembers","",sAutomationID)
				If Fn_RAC_ProjectMemberSelectionTreeOperations("Select", sAutomationID, "", objSelectedMembersList) = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sAutomationID) & " ] as failed to select the node in selected members tree.","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
				
				'Right click on it
				objSelectedMembersList.OpenContextMenu sAutomationID
				
				'Select "Set Priviledged Users" menu 
				objProjectWindow.WinMenu("wmnu_ContextMenu").Select "Select a Project team administrator"
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
				
				If Err.Number <> 0 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to set member [" & sAutomationID & "] as Project team administrator.","","","","","")
					Call Fn_ExitTest()
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully set member [" & sAutomationID & "] as Project team administrator.","","","","","")
				End If	
			End If
		Next
		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify node in selected members list
	Case "VerifyNodeInSelectedMembersList"
		
		'Get object of selected members tree from XML
		Set objSelectedMembersList = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jtree_SelectedMembers","")
		
		'Checking existance of [ Selected Members ] Tree
		If Fn_UI_Object_Operations("RAC_Project_DefinitionOperations","Exist", objSelectedMembersList,"","","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on selected members tree as the tree does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
		'Verify existence
		sAutomationID = Fn_Setup_GetTestUserDetailsFromExcelOperations("getusernodetreepathforprojectselectedmembers","",sAutomationID)
		If Fn_RAC_ProjectMemberSelectionTreeOperations("Verify", sAutomationID, "", objSelectedMembersList) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify existence of node [" & CStr(sAutomationID) & " ] in the selected members tree list..","","","","DONOTSYNC","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified existence of node [" & CStr(sAutomationID) & " ] in the selected members tree list..","","","","DONOTSYNC","")
		End If
		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify node in selected members list
	Case "RemoveSelectedMember"
		
		'Get object of selected members tree from XML
		Set objSelectedMembersList = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jtree_SelectedMembers","")
		
		'Checking existance of [ Selected Members ] Tree
		If Fn_UI_Object_Operations("RAC_Project_DefinitionOperations","Exist", objSelectedMembersList,"","","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on selected members tree as the tree does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Select the node to be removed
		sAutomationID = Fn_Setup_GetTestUserDetailsFromExcelOperations("getusernodetreepathforprojectselectedmembers","",sAutomationID)
		If Fn_RAC_ProjectMemberSelectionTreeOperations("Select", sAutomationID, "", objSelectedMembersList) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sAutomationID) & " ] as failed to select the node in selected members tree.","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
		
		'Click on Remove button
		If Fn_UI_JavaButton_Operations("RAC_Project_DefinitionOperations", "Click", objProjectWindow, "jbtn_RemoveUser") = False Then	
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [Remove] button after selecting node in selected members tree.","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully removed member [" & sAutomationID & "] from Project","","","","","")
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
		
		'Verify if user is removed successfully
		If Fn_RAC_ProjectMemberSelectionTreeOperations("Verify", sAutomationID, "", objSelectedMembersList) = True Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remove user node [ " & CStr(sAutomationID) & " ] in selected members tree.","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully removed user node [ " & CStr(sAutomationID) & " ] in selected members tree.","","","","","")
		End If
	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
	Case Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Invalid operation [ " & Cstr(sAction) & " ]","","","","","")
		
End Select

'Click on button if specified by user
If sButtonName <> "" Then
	If Fn_UI_JavaButton_Operations("RAC_Project_DefinitionOperations", "Click", objProjectWindow, "jbtn_" & sButtonName) = False Then	
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [" & sButtonName & "] button after performing action [" & sAction & "]","","","","","")
		Call Fn_ExitTest()
	Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on [" & sButtonName & "] button after performing action [" & sAction & "]","","","","","")
	End If
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
	If sButtonName="Modify" Or sButtonName="jbtn_Modify" Then
		If objProjectWindow.JavaWindow("jwnd_MissingProjectTeamAdminInTeam").Exist(10) Then
			If Fn_UI_JavaButton_Operations("RAC_Project_DefinitionOperations", "Click", objProjectWindow.JavaWindow("jwnd_MissingProjectTeamAdminInTeam"), "jbtn_OK") = False Then	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ OK ] button of [ Missing Project Team Admin In Team ] dialog after performing action [" & sAction & "]","","","","","")
				Call Fn_ExitTest()
			End iF
		End If
	End If
End If
Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] operation on navigation tree due to error number as [ " & Cstr(Err.Number) & " ] and error description as [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing all objects
Set objMemberSelectionTree = Nothing
Set objProjectWindow = Nothing
Set objSelectedMembersList = Nothing

Function Fn_ExitTest()	
	'Releasing all objects
	Set objMemberSelectionTree = Nothing
	Set objProjectWindow = Nothing
	Set objSelectedMembersList = Nothing
	ExitTest
End Function

