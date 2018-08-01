'! @Name 			RAC_MyWorklist_TreeNodeOperations
'! @Details 		This Action word to perform operations on my worklist tree
'! @InputParam1 	sAction = Action Name
'! @InputParam2 	sNodeName = Node path
'! @InputParam3 	sPopupMenu = Popup menu name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com 
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			29 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_TreeNodeOperations","RAC_MyWorklist_TreeNodeOperations",OneIteration,"Exist","My Worklist~Sunny Ruparel (502425666) Inbox~Tasks To Perform~WS00000032/01;1-ECR for testing WF issue (Set Pending status and Derive ECN)",""
'! @Example 		LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_TreeNodeOperations","RAC_MyWorklist_TreeNodeOperations",OneIteration,"Select","My Worklist~Sunny Ruparel (502425666) Inbox~Tasks To Perform~WS00000032/01;1-ECR for testing WF issue (Set Pending status and Derive ECN)",""
'! @Example 		LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_TreeNodeOperations","RAC_MyWorklist_TreeNodeOperations",OneIteration,"Expand","My Worklist~Sunny Ruparel (502425666) Inbox~Tasks To Perform~WS00000032/01;1-ECR for testing WF issue (Set Pending status and Derive ECN)",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeName,sPopupMenu
Dim objTcDefaultApplet,objMyWorklistTree
Dim aNodeName,aMenuList
Dim sTreeItem
Dim iCounter
Dim bFlag

bFlag=False

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sPopupMenu = Parameter("sPopupMenu")

'Creating object of [ MyWorklist ] tree
Set objMyWorkListTree=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","jtree_MyWorkListTree","")
Set objTcDefaultApplet=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","jwnd_TcDefaultApplet","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_TreeNodeOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of [ MyWorklist ] tree
If Fn_UI_Object_Operations("RAC_MyWorklist_TreeNodeOperations","Exist",objMyWorkListTree,"","","") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sNodeName) & " ] as [ My WorkList ] tree does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capture business functionality start time	
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","My Worklist Tree Node Operations",sAction,"","")

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Set focus on MyWorkListTree 
objMyWorklistTree.Object.SetFocus
Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "Exist","VerifyExist","VerifyNonExist"
		sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree,sNodeName,"")
		If sTreeItem = False Then
			If sAction="Exist" Then
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				DataTable.SetCurrentRow 1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_MyWorklist_TreeNodeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] does not exist under my worklist tree","","","","","")
				Call Fn_ExitTest()
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","My Worklist Tree Node Operations",sAction,"","")
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified node [ " & Cstr(sNodeName) & " ] does not exist under my worklist tree","","","","DONOTSYNC","")
			End If
		Else
			If sAction="Exist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","My Worklist Tree Node Operations",sAction,"","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				DataTable.SetCurrentRow 1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_MyWorklist_TreeNodeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","My Worklist Tree Node Operations",sAction,"","")
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified node [ " & Cstr(sNodeName) & " ] exist under my worklist tree","","","","DONOTSYNC","")
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] exist under my worklist tree","","","","","")
				Call Fn_ExitTest()
			End If
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case  "Select"
		sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree, sNodeName, "")
		If sTreeItem = False Then
			Call Fn_RAC_ReadyStatusSync(3)
			wait 6
			Call Fn_RAC_ReadyStatusSync(3)
			sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree, sNodeName, "")
			If sTreeItem = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ] as specified node does not exist under my worklist tree","","","","","")
				Call Fn_ExitTest()
			End IF
		End If
		
		objMyWorklistTree.Select sTreeItem
		Call Fn_RAC_ReadyStatusSync(1)
		objMyWorklistTree.Click 0,0		
		Call Fn_RAC_ReadyStatusSync(1)		
		objMyWorklistTree.Select sTreeItem
		Call Fn_RAC_ReadyStatusSync(1)
		objMyWorklistTree.Click 0,0
		
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ] from my worklist tree due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","My Worklist Tree Node Operations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected node [ " & CStr(sNodeName) & " ] from my worklist tree","","","","DONOTSYNC","") 
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "MultiSelect" 
		aNodeName = Split(sNodeName,"^")
		For iCounter = 0 to  Ubound(aNodeName)
			sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree, aNodeName(iCounter), "")					
			If sTreeItem = False Then
				Call Fn_RAC_ReadyStatusSync(3)
				wait 6
				Call Fn_RAC_ReadyStatusSync(3)
				sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree, aNodeName(iCounter), "")					
				If sTreeItem = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multi select node [ " & CStr(sNodeName) & " ] as specified node does not exist under my worklist tree","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			If iCounter <> 0 Then
				objMyWorklistTree.ExtendSelect sTreeItem
			Else
				objMyWorklistTree.Select sTreeItem
			End If
			If Err.Number < 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multi select nodes [ " & CStr(sNodeName) & " ] from my worklist tree","","","","","") 
				Call Fn_ExitTest()
			Else
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","My Worklist Tree Node Operations",sAction,"","")
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully multi selected nodes [ " & CStr(sNodeName) & " ] from my worklist tree","","","","DONOTSYNC","") 
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case  "Expand"
		If Instr(1,sNodeName,"Tasks To Perform") Or Instr(1,sNodeName,"Tasks To Track") Then
			bFlag=True
		End If
		
		sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree, sNodeName, "")
		If sTreeItem = False Then
			Call Fn_RAC_ReadyStatusSync(3)
			wait 6
			Call Fn_RAC_ReadyStatusSync(3)
			sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree, sNodeName, "")
			If sTreeItem = False Then	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & CStr(sNodeName) & " ] as specified node does not exist under my worklist tree","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		objMyWorklistTree.Expand sTreeItem
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & CStr(sNodeName) & " ] from my worklist tree","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","My Worklist Tree Node Operations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully expand node [ " & CStr(sNodeName) & " ] from my worklist tree","","","","DONOTSYNC","") 
		End If
		
		If bFlag=True Then
			wait 0,500
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "Collapse"
		sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree, sNodeName, "")
		If sTreeItem = False Then
			Call Fn_RAC_ReadyStatusSync(3)
			wait 6
			Call Fn_RAC_ReadyStatusSync(3)
			sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree, sNodeName, "")
			If sTreeItem = False Then	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Collapse node [ " & CStr(sNodeName) & " ] as specified node does not exist under my worklist tree","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		objMyWorklistTree.Collapse sTreeItem
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Collapse node [ " & CStr(sNodeName) & " ] from my worklist tree","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","My Worklist Tree Node Operations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Collapse node [ " & CStr(sNodeName) & " ] from my worklist tree","","","","DONOTSYNC","") 
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "PopupMenuSelect"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyWorklist_TreeNodeOperations","",sPopupMenu)
				
		aMenuList = split(sPopupMenu, ":",-1,1)
		iCounter = Ubound(aMenuList)
		
		sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree,sNodeName,"")
		If sTreeItem = False Then
			Call Fn_RAC_ReadyStatusSync(3)
			wait 6
			Call Fn_RAC_ReadyStatusSync(3)
			sTreeItem = Fn_RAC_GetMyWorklistNodePath(objMyWorklistTree,sNodeName,"")	
			If sTreeItem = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select  popup menu [ " & CStr(sPopupMenu) & " ] of node [ " & CStr(sNodeName) & " ] as specified node does not exist under my worklist tree","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		objMyWorklistTree.Select sTreeItem
		objMyWorklistTree.OpenContextMenu sTreeItem
		Select Case iCounter
			Case "0"						
				sPopupMenu = objTcDefaultApplet.WinMenu("wmnu_ContextMenu").BuildMenuPath(aMenuList(0))
			Case "1"
				sPopupMenu = objTcDefaultApplet.WinMenu("wmnu_ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
		End Select		
		objTcDefaultApplet.WinMenu("wmnu_ContextMenu").Select sPopupMenu
		
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & CStr(sPopupMenu) & " ] of node [ " & CStr(sNodeName) & " ] from my worklist tree due to error [ " & Cstr(err.description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","My Worklist Tree Node Operations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected popup menu [ " & CStr(sPopupMenu) & " ] of node [ " & CStr(sNodeName) & " ] from my worklist tree","","","","DONOTSYNC","") 
		End If		
End Select

'Releasing object
Set objMyWorklistTree = Nothing
Set objTcDefaultApplet=Nothing

Function Fn_ExitTest()
	Set objMyWorklistTree =Nothing
	Set objTcDefaultApplet=Nothing
	ExitTest
End Function

