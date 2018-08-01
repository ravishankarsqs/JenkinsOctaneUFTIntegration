'! @Name 			RAC_Project_MemberSelection
'! @Details 		To perform operations on Project Manager module member selection tasks
'! @InputParam1 	sAction 					: String to indicate what action is to be performed on navigation tree e.g. Select, Expand
'! @InputParam2 	sNodeDetailsAutomationID	: Automation ID of the user which will be a node in member selection tree
'! @Author 			Kundan Kudale kundan.kudale@sqs.com
'! @Reviewer 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Date 			11 Nov 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Project\RAC_Project_MemberSelection", "RAC_Project_MemberSelection", oneIteration, "AddMemberToProject", "TestUserEngineeringDesigner1"
'! @Example 		LoadAndRunAction "RAC_Project\RAC_Project_MemberSelection", "RAC_Project_MemberSelection", oneIteration, "VerifyNodeInSelectedMembersList", "TestUserEngineeringDesigner1"
'! @Example 		LoadAndRunAction "RAC_Project\RAC_Project_MemberSelection", "RAC_Project_MemberSelection", oneIteration, "RemoveSelectedMember", "TestUserEngineeringDesigner1"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeDetailsAutomationID, sMemberDetails
Dim objMemberSelectionTree, objProjectWindow, objSelectedMembersList
Dim aMemberDetails,sTemp,iCounter

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get parameter values in local variables
sAction = Parameter("sAction")
sNodeDetailsAutomationID = Parameter("sNodeDetailsAutomationID")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Project_MemberSelection"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Get the member details from excel file
If sNodeDetailsAutomationID <> "" Then
	sMemberDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sNodeDetailsAutomationID)	
	If sMemberDetails <> False Then
		aMemberDetails = Split(sMemberDetails,"~")
	Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as failed to get user details for [ " & CStr(sNodeDetailsAutomationID) & " ]","","","","","")
		Call Fn_ExitTest()
	End If
End If

'creating object of [ Navigation Tree ]
Set objMemberSelectionTree = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jtree_MemberSelection","")

'creating object of [ Project Window ]
Set objProjectWindow = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jwnd_ProjectDefaultWindow","")

Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Project Manager member selection tree Node Operations",sAction,"Node name",sNodeDetailsAutomationID)

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER = DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select node from navigation tree
	Case "AddMemberToProject"
		
		'Checking existance of [ Member Selection ] Tree
		If Fn_UI_Object_Operations("RAC_Project_MemberSelection","Exist", objMemberSelectionTree,"","","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on member selection tree as [ Member Selection Tree ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If

		sTemp = Split(aMemberDetails(2),".")
		For iCounter = UBound(sTemp) To 0 Step -1
			If iCounter=UBound(sTemp) Then
				aMemberDetails(2)=sTemp(iCounter)
			Else
				aMemberDetails(2)=aMemberDetails(2)+"~"+sTemp(iCounter)
			End If
		Next
			
		sNodeDetailsAutomationID = aMemberDetails(2) & "~" & aMemberDetails(3) & "~" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",Parameter("sNodeDetailsAutomationID")) & " (" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",Parameter("sNodeDetailsAutomationID")) & ")"
		'Select the node on which operation is to be performed
		If Fn_RAC_ProjectMemberSelectionTreeOperations("Select", sNodeDetailsAutomationID, "", objMemberSelectionTree) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sNodeDetailsAutomationID) & " ] as failed to select the node in member selection tree.","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
		
		'Click on Add button
		If Fn_UI_JavaButton_Operations("RAC_Project_MemberSelection", "Click", objProjectWindow, "jbtn_AddUser") = False Then	
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [Add] button after selecting node in member selection tree.","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added member [" & sNodeDetailsAutomationID & "] to project.","","","","","")
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify node in selected members list
	Case "VerifyNodeInSelectedMembersList"
		
		'Get object of selected members tree from XML
		Set objSelectedMembersList = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jtree_SelectedMembers","")
		
		'Checking existance of [ Selected Members ] Tree
		If Fn_UI_Object_Operations("RAC_Project_MemberSelection","Exist", objSelectedMembersList,"","","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on selected members tree as the tree does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Verify existence
		sNodeDetailsAutomationID = aMemberDetails(2) & "." & aMemberDetails(3) & "~" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",Parameter("sNodeDetailsAutomationID")) & " (" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",Parameter("sNodeDetailsAutomationID")) & ")"
		If Fn_RAC_ProjectMemberSelectionTreeOperations("Verify", sNodeDetailsAutomationID, "", objSelectedMembersList) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify existence of node [" & CStr(sNodeDetailsAutomationID) & " ] in the selected members tree list..","","","","DONOTSYNC","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified existence of node [" & CStr(sNodeDetailsAutomationID) & " ] in the selected members tree list..","","","","DONOTSYNC","")
		End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify node in selected members list
	Case "VerifyNodeNotInSelectedMembersList"
		
		'Get object of selected members tree from XML
		Set objSelectedMembersList = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jtree_SelectedMembers","")
		
		'Checking existance of [ Selected Members ] Tree
		If Fn_UI_Object_Operations("RAC_Project_MemberSelection","Exist", objSelectedMembersList,"","","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on selected members tree as the tree does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Verify existence
		sNodeDetailsAutomationID = aMemberDetails(2) & "." & aMemberDetails(3) & "~" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",Parameter("sNodeDetailsAutomationID")) & " (" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",Parameter("sNodeDetailsAutomationID")) & ")"
		If Fn_RAC_ProjectMemberSelectionTreeOperations("VerifyNonExist", sNodeDetailsAutomationID, "", objSelectedMembersList) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify non existence of node [" & CStr(sNodeDetailsAutomationID) & " ] in the selected members tree list..","","","","DONOTSYNC","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified non existence of node [" & CStr(sNodeDetailsAutomationID) & " ] in the selected members tree list..","","","","DONOTSYNC","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify node in selected members list
	Case "RemoveSelectedMember"
		
		'Get object of selected members tree from XML
		Set objSelectedMembersList = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jtree_SelectedMembers","")
		
		'Checking existance of [ Selected Members ] Tree
		If Fn_UI_Object_Operations("RAC_Project_MemberSelection","Exist", objSelectedMembersList,"","","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on selected members tree as the tree does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Select the node to be removed
		sNodeDetailsAutomationID = aMemberDetails(2) & "." & aMemberDetails(3) & "~" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",Parameter("sNodeDetailsAutomationID")) & " (" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",Parameter("sNodeDetailsAutomationID")) & ")"
		If Fn_RAC_ProjectMemberSelectionTreeOperations("Select", sNodeDetailsAutomationID, "", objSelectedMembersList) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sNodeDetailsAutomationID) & " ] as failed to select the node in selected members tree.","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
		
		'Click on Remove button
		If Fn_UI_JavaButton_Operations("RAC_Project_MemberSelection", "Click", objProjectWindow, "jbtn_RemoveUser") = False Then	
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [Remove] button after selecting node in selected members tree.","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully removed member [" & sNodeDetailsAutomationID & "] from Project","","","","","")
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
	Case Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Invalid operation [ " & Cstr(sAction) & " ]","","","","","")
		
End Select

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


'Systemutil.Run "\\s400v001.autoexpr.com\plmclntprod\Environments\Local\Menu_Starter\Starter\Teamcenter_Startup.exe"
