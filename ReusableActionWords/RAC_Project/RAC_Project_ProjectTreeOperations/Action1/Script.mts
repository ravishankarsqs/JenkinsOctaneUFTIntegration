'! @Name 			RAC_Project_ProjectTreeOperations
'! @Details 		To perform operations on Project module Projects tree
'! @InputParam1 	sAction 		: String to indicate what action is to be performed on Projects tree e.g. Select, Expand
'! @InputParam2 	sNodeName 		: Node name in Projects tree on which action is to be performed
'! @InputParam3 	sPopupMenu 		: Menu tag name from XML
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			10 Oct 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Project\RAC_Project_ProjectTreeOperations","RAC_Project_ProjectTreeOperations",oneIteration, "Select", "AutomatedTest",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeName,sPopupMenu
Dim objProjectsTree
Dim bFlag
Dim iCount, iPath
Dim sNode,sNodePath

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get parameter values in local variables
sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sPopupMenu = Parameter("sPopupMenu")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Project_ProjectTreeOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

bFlag = False

'creating object of [ Navigation Tree ]
Set objProjectsTree=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jtree_ProjectsTree","")
		
'Checking existance of [ Navigation ] Tree
If Fn_UI_Object_Operations("RAC_Project_ProjectTreeOperations","Exist",objProjectsTree,"","","") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sNodeName) & " ] as [ Navigation Tree ] does not exist","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Projects tree Node Operations",sAction,"Node name",sNodeName)

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select node from navigation tree
	Case "Select"
		iPath = Fn_RAC_GetJavaTreeNodePath(objProjectsTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [ " & Cstr(sNodeName) & " ] of project tree as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If			
		
		'Selecting node from tree
		objProjectsTree.Select iPath	
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Projects tree Node Operations",sAction,"Node name",sNodeName)
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected node [ " & Cstr(sNodeName) & " ] from Projects tree","","","",GBL_MICRO_SYNC_ITERATIONS,"")
		End If
		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Expand node from navigation tree
	Case "Expand"	
		If Fn_UI_JavaTree_Operations("RAC_Project_ProjectTreeOperations","Exist",objProjectsTree,"",sNodeName,"","")=False Then
			If Fn_UI_JavaTree_Operations("Fn_Project_ProjectTreeOperations","Expand",objProjectsTree,"",sNodeName,"","")=False Then 
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & CStr(sNodeName) & " ]","","","","","") 
				Call Fn_ExitTest()
			Else
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Projects tree Node Operations",sAction,"Node name",sNodeName)
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully expanded node [ " & Cstr(sNodeName) & " ] from Projects tree","","","",GBL_MICRO_SYNC_ITERATIONS,"")
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand Node [ " & Cstr(sNodeName) & " ] of project tree as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If											 	
	 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 'Case to check existance of node of navigation tree
	 Case "Exist","VerifyExist","VerifyNonExist"
		bFlag = True
		If Fn_RAC_GetJavaTreeNodePath(objProjectsTree,sNodeName,"","")=False Then
			bFlag = False
		End If
		
		If bFlag = False Then
			If sAction="Exist" Then
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				DataTable.SetCurrentRow 1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_Project_ProjectTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as project node [ " & Cstr(sNodeName) & " ] does not exist","","","","","")
				Call Fn_ExitTest()
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Projects tree Node Operations",sAction,"Node name",sNodeName)	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ project ] [ " & Cstr(sNodeName) & " ] does not exist","","","","DONOTSYNC","")
			End If
		Else
			If sAction="Exist" Then
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				DataTable.SetCurrentRow 1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_Project_ProjectTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Projects tree Node Operations",sAction,"Node name",sNodeName)		
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Projects tree Node Operations",sAction,"Node name",sNodeName)		
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ project ] [ " & Cstr(sNodeName) & " ] exist","","","","DONOTSYNC","")
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ project ] [ " & Cstr(sNodeName) & " ] exist","","","","","")
				Call Fn_ExitTest()
			End If
		End If	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Expand node from navigation tree
	Case "ExpandAndSelect"		
		sNode = Split(sNodeName,"~",-1,1)								
		For iCount = 0 To Ubound(sNode)-1
			If iCount = 0 Then
				sNodePath = "Home"
			Else
				sNodePath = sNodePath &"~"& sNode(iCount)
			End If				
			'Retrive node path
			If Fn_UI_JavaTree_Operations("RAC_Project_ProjectTreeOperations","Exist",objProjectsTree,"",sNodePath,"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand Node [ " & Cstr(sNodeName) & " ] of project tree as specified node does not exist","","","","","")
				Call Fn_ExitTest()
			Else
				'Expanding node from navigation tree
				objProjectsTree.Expand sNodePath
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			End If
		Next
		If Fn_UI_JavaTree_Operations("Fn_Project_ProjectTreeOperations","Select",objProjectsTree,"",sNodeName,"","")=False Then 
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Projects tree Node Operations",sAction,"Node name",sNodeName)
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected node [ " & Cstr(sNodeName) & " ] from Projects tree","","","",GBL_MICRO_SYNC_ITERATIONS,"")
		End If	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
	Case Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Invalid operation [ " & Cstr(sAction) & " ]","","","","","")
End Select

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] operation on project tree due to error number as [ " & Cstr(Err.Number) & " ] and error description as [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing all objects
Set objProjectsTree=Nothing

Function Fn_ExitTest()	
	'Releasing all objects
	Set objProjectsTree=Nothing	
	ExitTest
End Function


