'! @Name 			RAC_ChangeManager_NavigationTreeOperations
'! @Details 		This Action word to perform operations on Change Manager Nav tree
'! @InputParam1 	sAction 			: Action Name
'! @InputParam2		sNodeName 			: Node path
'! @InputParam3 	sPopupMenu 			: Popup menu tag name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com 
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			05 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_ChangeManager\RAC_ChangeManager_NavigationTreeOperations","RAC_ChangeManager_NavigationTreeOperations",OneIteration,"Expand","Change Home",""
'! @Example 		LoadAndRunAction "RAC_ChangeManager\RAC_ChangeManager_NavigationTreeOperations","RAC_ChangeManager_NavigationTreeOperations",OneIteration,"Expand","Change Home~My Open ECR's",""
'! @Example 		LoadAndRunAction "RAC_ChangeManager\RAC_ChangeManager_NavigationTreeOperations","RAC_ChangeManager_NavigationTreeOperations",OneIteration,"Select","Change Home~My Open ECR's~Load All",""
'! @Example 		LoadAndRunAction "RAC_ChangeManager\RAC_ChangeManager_NavigationTreeOperations","RAC_ChangeManager_NavigationTreeOperations",OneIteration,"VerifyExist","Change Home~My Open ECR's~AL-ECR00131/A-EcrChange_302",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeName,sPopupMenu
Dim objNavigationTree
Dim sObjectTypeName
Dim iPath

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get parameter values in local variables
sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sPopupMenu = Parameter("sPopupMenu")

'Creating object of change manager [ Navigation tree ]
Set objNavigationTree=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ChangeManager_OR","jtree_NavigationTree","")

'Retrive popup menu
If sPopupMenu<>"" Then
	sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_ComponentTabTreeOperations","",sPopupMenu)
End If

'Checking existance of change manager [ Navigation tree 
If Fn_UI_Object_Operations("RAC_ChangeManager_NavigationTreeOperations","Exist", objNavigationTree,"","","") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sNodeName) & " ] as change manager [ navigation tree ] does not exist","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Change Manager Navigation Tree Node Operations",sAction,"","")

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "Exist","VerifyExist","VerifyNonExist"
		iPath = Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName(objNavigationTree,sNodeName,"","")
		If iPath = False Then
			If sAction="Exist" Then
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				DataTable.SetCurrentRow 1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_ChangeManager_NavigationTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] does not exist","","","","","")
				Call Fn_ExitTest()
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Change Manager Navigation Tree Node Operations",sAction,"","")
				If GBL_LOG_ADDITIONAL_INFORMATION<>"" Then
					sObjectTypeName=GBL_LOG_ADDITIONAL_INFORMATION
				Else
					sObjectTypeName="node"
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist","","","","DONOTSYNC","")
			End If
		Else
			If sAction="Exist" Then
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				DataTable.SetCurrentRow 1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_ChangeManager_NavigationTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Change Manager Navigation Tree Node Operations",sAction,"","")
				sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
				If sObjectTypeName="" Then
					sObjectTypeName="node"
				End If	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] exist","","","","DONOTSYNC","")
			ElseIf sAction="VerifyNonExist" Then
				If GBL_LOG_ADDITIONAL_INFORMATION<>"" Then
					sObjectTypeName=GBL_LOG_ADDITIONAL_INFORMATION
				Else
					sObjectTypeName="node"
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] exist","","","","","")
				Call Fn_ExitTest()
			End If
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case  "Select"
		iPath = Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName(objNavigationTree, sNodeName, "","")
		If iPath = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ] as specified node does not exist under change manager navigation tree","","","","","")
			Call Fn_ExitTest()
		End If
		objNavigationTree.Select iPath
		objNavigationTree.Click 0,0
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ] from change manager navigation tree","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Change Manager Navigation Tree Node Operations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected node [ " & CStr(sNodeName) & " ] from change manager navigation tree","","","","","") 
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case  "Expand"
		iPath = Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName(objNavigationTree, sNodeName, "","")
		If iPath = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & CStr(sNodeName) & " ] as specified node does not exist under change manager navigation tree","","","","","")
			Call Fn_ExitTest()
		End If
		objNavigationTree.Expand iPath
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & CStr(sNodeName) & " ] from change manager navigation tree","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Change Manager Navigation Tree Node Operations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully expand node [ " & CStr(sNodeName) & " ] from change manager navigation tree","","","","","") 
		End If	
End Select

'Releasing object
Set objNavigationTree = Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objNavigationTree =Nothing
	ExitTest
End Function

