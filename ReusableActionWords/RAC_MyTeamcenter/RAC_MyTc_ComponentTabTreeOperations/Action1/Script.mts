'! @Name 			RAC_MyTc_ComponentTabTreeOperations
'! @Details 		To perform operations on My Teamcenter module Component tab tree
'! @InputParam1 	sAction 		: String to indicate what action is to be performed on Component tab tree e.g. Select, Expand
'! @InputParam2 	sNodeName 		: Node name in Component tab tree on which action is to be performed
'! @InputParam3 	sPopupMenu 		: Menu name tag name of XML node
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			26 Mar 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_ComponentTabTreeOperations", "RAC_MyTc_ComponentTabTreeOperations", oneIteration, "Select","AutomatedTest",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeName,sPopupMenu
Dim objComponentTree,objCurrentNode,objSelection,objContextMenu,objMyTeamcenterWindow
Dim iCounter,iPath,iSelectionCount
Dim sNode,iCount,sNodePath
Dim aNodeName,aPopupMenu
Dim sParentPath,sReturn,sObjectTypeName
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get parameter values in local variables
sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sPopupMenu = Parameter("sPopupMenu")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyTc_ComponentTabTreeOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

bFlag = False

'creating object of [ Component tab tree ]
Set objComponentTree=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyTeamcenter_OR","jtree_ComponentTree","")
'creating object of myteamcenter main window
Set objMyTeamcenterWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyTeamcenter_OR","jwnd_MyTeamcenter","")
'Select RMB menu
Set objContextMenu=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyTeamcenter_OR","wmnu_ContextMenu","")
		
'Checking existance of [ Component ] Tree
If Fn_UI_Object_Operations("RAC_MyTc_ComponentTabTreeOperations","Exist", objComponentTree,"","","") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sNodeName) & " ] as [ Component tab tree ] does not exist","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select node from Component tab tree
	Case "Select"
		'Initial Item parent Path
		aNodeName = Split (sNodeName, "~")
		For iCounter =0 to UBound(aNodeName)-1
			If sParentPath = "" Then
				sParentPath  = aNodeName(iCounter)
			Else
				sParentPath  = sParentPath & "~" & aNodeName(iCounter)
			End If
		Next				
		'Expanding paren node
		If UBound(aNodeName) > 0 Then
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_ComponentTabTreeOperations","RAC_MyTc_ComponentTabTreeOperations", oneIteration, "Expand", sParentPath, ""
		End If				
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [ " & Cstr(sNodeName) & " ] of Component tab tree as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If			
		'Selecting node from tree
		objComponentTree.Select iPath			
		If err.number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="Node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Component tab tree","","","",GBL_MICRO_SYNC_ITERATIONS,"")
		End If				
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Deselect node from Component tab tree
	Case "Deselect"
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to deselect Node [ " & Cstr(sNodeName) & " ] of Component tab tree as specified node does not exist.","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Deselecting node from tree
		objComponentTree.Select iPath				
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to deselect node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="Node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully deselected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] of Component tab tree","","","",GBL_MICRO_SYNC_ITERATIONS,"")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to multi select nodes from Component tab tree
	Case "Multiselect"
		aNodeName=Split(sNodeName,"^")
		For iCounter=0 To UBound(aNodeName)
			iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, aNodeName(iCounter) , "~", "@")
			If iPath=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Multi Select Node [" & Cstr(sNodeName) & "] of Component tab tree as node [" & Cstr(aNodeName(iCounter)) & "] does not exist","","","","","")
				Call Fn_ExitTest()
			Else
				'Multiselecting items
				If iCounter=0 Then
					objComponentTree.Select iPath				
				Else
					objComponentTree.ExtendSelect iPath				
				End If
				
				If Err.Number <> 0 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multi select nodes [ " & CStr(sNodeName) & " ]","","","","","") 
					Call Fn_ExitTest()
				Else
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)
					sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
					If sObjectTypeName="" Then
						sObjectTypeName="nodes"
					End If
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]","","","","","") 
				End If
			End If
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Expand node from Component tab tree
	Case "Expand"	
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Expand Node [" & Cstr(sNodeName) & "] of Component tab tree as specified node does not exist in Component tab tree.","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Expanding node from Component tab tree
		objComponentTree.Expand iPath
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully expanded [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]","","","","","") 
		End If				
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Collapse node from Component tab tree
	Case "Collapse"
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse node [" & Cstr(sNodeName) & "] of Component tab tree as specified node does not exist in Component tab tree.","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Collapse node from Component tab tree
		objComponentTree.Collapse iPath
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully collapsed [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]","","","","","") 
		End If					
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Double Click node from Component tab tree
	Case "DoubleClick"
		Dim intX, intY, intWidth, intHeight		
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to double click node [" & Cstr(sNodeName) & "] of Component tab tree as specified node does not exist in Component tab tree.","","","","","")
			Call Fn_ExitTest()
		End If
		
		objComponentTree.Select iPath
		wait GBL_MIN_MICRO_TIMEOUT
		Call Fn_CommonUtil_KeyBoardOperation("SendKeys", "{ENTER}")
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to double click node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully double click [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]","","","","","") 
		End If				 
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB menu
	Case "PopupMenuSelect"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_ComponentTabTreeOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu of node [" & Cstr(sNodeName) & "]" ,"","","","","")
			Call Fn_ExitTest()
		End If
		
		'Build the Popup menu to be selected
		aPopupMenu = Split(sPopupMenu, ":",-1,1)
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu of node [" & Cstr(sNodeName) & "] specified node does not exist in Component tab tree","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Selecting node from tree
		objComponentTree.Select iPath
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)		
		
		'Opening context menu on selected node
		If Fn_UI_JavaTree_Operations("RAC_MyTc_ComponentTabTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"jtree_ComponentTree",iPath,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu as fail to open context menu of node [ " & Cstr(sNodeName) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
								
		Select Case Ubound(aPopupMenu)
			Case "0"
				 sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0))
				 objContextMenu.Select sPopupMenu
			Case "1"
				sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0),aPopupMenu(1))
				objContextMenu.Select sPopupMenu
			Case "2"
				sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0),aPopupMenu(1),aPopupMenu(2))
				objContextMenu.Select sPopupMenu
			Case Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to build menu path of context menu of node [" & Cstr(sNodeName) & "]","","","","","")
				Call Fn_ExitTest()
		End Select
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform context menu operation with error number [ " & Err.Number & " ] and error description [" & Err.Description & "]","","","","","")
			Call Fn_ExitTest()
		End If
		sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
		If sObjectTypeName="" Then
			sObjectTypeName="node"
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform RMB menu [ " & Cstr(sPopupMenu) & " ] operation of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Component tab tree","","","","","")
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB menu
	Case "PopupMenuSelect_Ext"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_ComponentTabTreeOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to get menu value from XML for XML tag [" & Parameter("sPopupMenu") & "]" ,"","","","","")
			Call Fn_ExitTest()
		End If
		
		'Build the Popup menu to be selected
		aPopupMenu = Split(sPopupMenu, ":",-1,1)
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu of node [" & Cstr(sNodeName) & "] specified node does not exist in Component tab tree","","","","","")
			Call Fn_ExitTest()
		Else
			'Selecting node from tree
			objComponentTree.Select iPath
		End If
				
		'Opening context menu on selected node
		If Fn_UI_JavaTree_Operations("RAC_MyTc_ComponentTabTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"jtree_ComponentTree",iPath,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu as fail to open context menu of node [ " & Cstr(sNodeName) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		'selecting RMB menu	
		sPopupMenu = Fn_UI_JavaMenu_Operations("RAC_MyTc_ComponentTabTreeOperations","Select",objMyTeamcenterWindow,sPopupMenu)				
		If sPopupMenu=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform context menu operation with error number [ " & Err.Number & " ] and error description [" & Err.Description & "]","","","","","")
			Call Fn_ExitTest()
		End If
		sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
		If sObjectTypeName="" Then
			sObjectTypeName="node"
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform RMB menu [ " & Cstr(sPopupMenu) & " ] operation of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Component tab tree","","","","","")
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)
	 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 'Checking existance of RMB menu by selecting multiple nodes
	Case "MultiSelectContextMenuExist"
		aNodeName = split(sNodeName,"^",-1,1)
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_ComponentTabTreeOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to get menu value from XML for XML tag [" & Parameter("sPopupMenu") & "]" ,"","","","","")
			Call Fn_ExitTest()
		End If
		
		'Multiselecting nodes
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_ComponentTabTreeOperations","RAC_MyTc_ComponentTabTreeOperations",OneIteration,"Multiselect",sNodeName,""
								
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")		
		'Open context menu
		If Fn_UI_JavaTree_Operations("RAC_MyTc_ComponentTabTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"jtree_ComponentTree",iPath,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail  to Open Context Menu window","","","","","")
			Call Fn_ExitTest()
		End If
		
		'chekcing existance of RMB menu
		If objContextMenu.GetItemProperty (sPopupMenu,"Exists") = True Then
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified RMB menu [ " & Cstr(sPopupMenu) & " ] appears after multiselecting nodes from Component tab tree","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] does not appears after multiselecting nodes from Component tab tree","","","","","")
			Call Fn_ExitTest()
		End If				
	 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 'Case to check existance of node of Component tab tree
	 Case "Exist","VerifyExist","VerifyNonExist"
		bFlag = True
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		
		If iPath=False Then
			bFlag = False
		Else
			aNodeName = split(Replace(iPath,"#",""),":")
			Set objCurrentNode = objComponentTree.Object
			For iCounter = 0 to UBound(aNodeName) -1
				Set objCurrentNode = objCurrentNode.GetItem(aNodeName(iCounter))
				If cBool(objCurrentNode.getExpanded()) = False Then
					bFlag = False
					Exit for
				End If
			Next
			Set objCurrentNode = Nothing									
		End If
		If bFlag = False Then
			If sAction="Exist" Then
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				DataTable.SetCurrentRow 1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_ComponentTabTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] does not exist","","","","","")
				Call Fn_ExitTest()
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)	
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
				DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_ComponentTabTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)		
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)		
				sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
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
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to check existance of RMB menu
	Case "PopupMenuExist"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_ComponentTabTreeOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] due to invalid RMB menu [ " & Cstr(Parameter("sPopupMenu")) & "]","","","","","")
			Call Fn_ExitTest()
		End If

		'Build the Popup menu to be selected
		aPopupMenu = Split(sPopupMenu, ":",-1,1)
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] of node [" & sNodeName & "] in Component tab tree as the specified node does not exist","","","","","")
			Call Fn_ExitTest()
		Else
			'Selecting node from tree
			objComponentTree.Select iPath
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If		
		'Open context menu
		If Fn_UI_JavaTree_Operations("RAC_MyTc_ComponentTabTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"Component tab tree",iPath,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as Context menu does not exist","","","","","")
			Call Fn_ExitTest()
		End If
				
		'Build RMB menu path
		Select Case Ubound(aPopupMenu)
			Case "0"
				sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0))
			Case "1"
				sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0),aPopupMenu(1))
			Case "2"
				sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0),aPopupMenu(1),aPopupMenu(2))
			Case Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] does not exist for node [ " & Cstr(sNodeName) & " ] in Component tab tree","","","","","")
				Call Fn_ExitTest()						
		End Select
		'Checking existance of RMB menu
		If objContextMenu.GetItemProperty (sPopupMenu,"Exists") = True Then
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If	
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Sucessfully verified RMB menu [ " & Cstr(sPopupMenu) & " ] exist for [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] in Component tab tree","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] does not exist for node [ " & Cstr(sNodeName) & " ] in Component tab tree","","","","","")
			Call Fn_ExitTest()	
		End If
		Call Fn_CommonUtil_KeyBoardOperation("SendKeys", "{ESC}")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to check RMB menu state
	Case "PopupMenuEnabled"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_ComponentTabTreeOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as Context menu does not exist","","","","","")
			Call Fn_ExitTest()	
		End If

		'Build the Popup menu to be selected
		aPopupMenu = Split(sPopupMenu, ":",-1,1)
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform popup menu enabled action on Node [" & Cstr(sNodeName) & "] of Component tab tree as the specified node does not exist.","","","","","")
			Call Fn_ExitTest()
		Else
			'Selecting node from tree
			objComponentTree.Select iPath
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)			
		End If

		'Open context menu
		If Fn_UI_JavaTree_Operations("RAC_MyTc_ComponentTabTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"Component tab tree",iPath,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as Context menu does not exist","","","","","")
			Call Fn_ExitTest()
		End If
				
		'Build RMB menu path
		Select Case Ubound(aPopupMenu)
			Case "0"
				sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0))
			Case "1"
				sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0),aPopupMenu(1))
			Case "2"
				sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0),aPopupMenu(1),aPopupMenu(2))
			Case Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] does not exist for node [ " & Cstr(sNodeName) & " ] in Component tab tree","","","","","")
				Call Fn_ExitTest()
		End Select
		
		'Checking existance of RMB menu
		If objContextMenu.GetItemProperty (sPopupMenu,"Enabled") = True Then
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)	
			sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Sucessfully verified RMB menu [ " & Cstr(sPopupMenu) & " ] enabled for [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] in Component tab tree","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] disabled for node [ " & Cstr(sNodeName) & " ] in Component tab tree","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'getting selected node path
	Case "GetSelectedNodePath"
		If objComponentTree.Object.getSelectionCount() > 0 Then
			Set objSelection = objComponentTree.Object.getSelection()
			iSelectionCount =objComponentTree.Object.getSelectionCount()
			For iCounter = 0 to iSelectionCount - 1 
				sReturn = ""
				Set objCurrentNode = objSelection.mic_arr_get(iCounter)
				sReturn = objCurrentNode.getData().toString() 
				Do while not(objCurrentNode.getParentItem() is nothing)
					Set objCurrentNode = objCurrentNode.getParentItem()
					sReturn = objCurrentNode.getData().toString() & "~" & sReturn 
				Loop
				Set objCurrentNode = Nothing				
				IF iCounter=0 Then
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
					DataTable.SetCurrentRow 1		
					DataTable.Value("ActionWordName","Global")= "RAC_MyTc_ComponentTabTreeOperations"
					DataTable.Value("ActionWordReturnValue","Global")= sReturn
				Else
					DataTable.Value("ActionWordReturnValue","Global")= DataTable.Value("ActionWordReturnValue","Global") & "^" & sReturn
				End IF	
			Next
			Set objSelection =Nothing
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "GetIndex"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ActionWordName","Global")= "RAC_MyTc_ComponentTabTreeOperations"
		DataTable.Value("ActionWordReturnValue","Global")= Cstr(Fn_RAC_GetJavaTreeNodeIndex(objComponentTree, sNodeName,"",""))
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "GetFirtNodeName"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ActionWordName","Global")= "RAC_MyTc_ComponentTabTreeOperations"
		DataTable.Value("ActionWordReturnValue","Global")= objComponentTree.Object.getItem(0).getData().toString()				
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Expand node from Component tab tree
	Case "ExpandAndSelect"		
		sNode = Split(sNodeName,"~",-1,1)								
		For iCount = 0 To Ubound(sNode)-1
			If sNode(iCount) <> "" Then
				If iCount = 0 Then
					sNodePath = "Home"
				Else
					sNodePath = sNodePath &"~"& sNode(iCount)
				End If				
				'Retrive node path
				iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodePath , "~", "@")
				If iPath=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand adn select Node [" + sNodePath + "] of Component tab tree as the specified node does not exist in Component tab tree.","","","","","")
					Call Fn_ExitTest()
				Else
					'Expanding node from Component tab tree
					objComponentTree.Expand iPath
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				End If
			End If
		Next		
		'Selecting node from Component tab tree
		iPath = Fn_RAC_GetJavaTreeNodePath(objComponentTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [" & Cstr(sNodeName) & "] of Component tab tree as the specified node does not exist in Component tab tree","","","","","")
			Call Fn_ExitTest()
		Else
			'Expanding node from Component tab tree
			objComponentTree.Select iPath
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)			
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter Component tab tree Node Operations",sAction,"Node name",sNodeName)	
		
		sObjectTypeName=Fn_RAC_GetTreeNodeType("ComponentTabTree","getobjecttypename")
		If sObjectTypeName="" Then
			sObjectTypeName="node"
		End If		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Component tab tree","","","","","")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
	Case Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Invalid operation [ " & Cstr(sAction) & " ]","","","","","")
End Select

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] operation on Component tab tree due to error number as [ " & Cstr(Err.Number) & " ] and error description as [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing all objects
Set objComponentTree=Nothing
Set objContextMenu = Nothing
Set objMyTeamcenterWindow = Nothing

Function Fn_ExitTest()	
	'Releasing all objects
	Set objComponentTree=Nothing
	Set objContextMenu = Nothing
	Set objMyTeamcenterWindow = Nothing
	ExitTest
End Function

