'! @Name 			RAC_MyTc_NavigationTreeOperations
'! @Details 		To perform operations on My Teamcenter module Navigation tree
'! @InputParam1 	sAction 		: String to indicate what action is to be performed on navigation tree e.g. Select, Expand
'! @InputParam2 	sNodeName 		: Node name in navigation tree on which action is to be performed
'! @InputParam3 	sPopupMenu 		: Menu tag name from XML
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			26 Mar 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Select", "AutomatedTest",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeName,sPopupMenu
Dim objNavigationTree,objCurrentNode,objSelection,objContextMenu,objMyTeamcenterWindow
Dim iCounter,iPath,iSelectionCount,iInstanceHandler
Dim sNode,iCount,sNodePath,sChildNodePath
Dim aNodeName,aPopupMenu, aPropertyNames
Dim sParentPath,sReturn,sObjectTypeName
Dim bFlag
Dim sDatasetType, sTempValue
Dim iActualOccuranceCount, iExpectedOccuranceCount
Dim sCurrentlySelectedNode

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get parameter values in local variables
sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sPopupMenu = Parameter("sPopupMenu")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyTc_NavigationTreeOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

bFlag = False

'creating object of [ Navigation Tree ]
Set objNavigationTree=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyTeamcenter_OR","jtree_NavigationTree","")
'creating object of myteamcenter main window
Set objMyTeamcenterWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyTeamcenter_OR","jwnd_MyTeamcenter","")
'Select RMB menu
Set objContextMenu=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyTeamcenter_OR","wmnu_ContextMenu","")
		
'Checking existance of [ Navigation ] Tree
If Fn_UI_Object_Operations("RAC_MyTc_NavigationTreeOperations","Exist", objNavigationTree,"","","") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & CStr(sNodeName) & " ] as [ Navigation Tree ] does not exist","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select node from navigation tree
	Case "Select","SelectAndRefresh"
		'Initial Item parent Path
		aNodeName = Split (sNodeName, "~")
		For iCounter =0 to UBound(aNodeName)-1
			If sParentPath = "" Then
				sParentPath  = aNodeName(iCounter)
			Else
				sParentPath  = sParentPath + "~" + aNodeName(iCounter)
			End If
		Next				
		'Expanding paren node
		If UBound(aNodeName) > 0 Then
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations", oneIteration, "ExpandExt", sParentPath, ""
		End If				
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [ " & Cstr(sNodeName) & " ] of navigation tree as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If			
		'Selecting node from tree
		objNavigationTree.Select iPath			
		If err.number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="Node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from navigation tree","","","","DONOTSYNC","")
		End If		
		If sAction="SelectAndRefresh" Then
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewRefresh"
		End If		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Deselect node from navigation tree
	Case "Deselect"
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to deselect Node [ " & Cstr(sNodeName) & " ] of navigation tree as specified node does not exist.","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Deselecting node from tree
		objNavigationTree.Deselect iPath				
		If err.number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to deselect node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="Node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully deselected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] of navigation tree","","","","DONOTSYNC","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to multi select nodes from navigation tree
	Case "Multiselect"
		aNodeName=Split(sNodeName,"^")
		For iCounter=0 To UBound(aNodeName)
			iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, aNodeName(iCounter) , "~", "@")
			If iPath=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Multi Select Node [" & Cstr(sNodeName) & "] of navigation tree as node [" & Cstr(aNodeName(iCounter)) & "] does not exist","","","","","")
				Call Fn_ExitTest()
			Else
				'Multiselecting items
				If iCounter=0 Then
					objNavigationTree.Select iPath				
				Else
					objNavigationTree.ExtendSelect iPath				
				End If
				
				If Err.Number <> 0 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multi select nodes [ " & CStr(sNodeName) & " ]","","","","","") 
					Call Fn_ExitTest()
				Else
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
					sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
					If sObjectTypeName="" Then
						sObjectTypeName="nodes"
					End If
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]","","","","DONOTSYNC","") 
				End If
			End If
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Expand node from navigation tree
	Case "Expand","ExpandExt"
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Expand Node [" & Cstr(sNodeName) & "] of navigation tree as specified node does not exist in navigation tree.","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Expanding node from navigation tree
		objNavigationTree.Expand iPath
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			If sAction="Expand" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
				sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
				If sObjectTypeName="" Then
					sObjectTypeName="node"
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully expanded [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]  from navigation tree","","","","DONOTSYNC","") 
			End If
		End If				
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Collapse node from navigation tree
	Case "Collapse"
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse node [" & Cstr(sNodeName) & "] of navigation tree as specified node does not exist in navigation tree.","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Collapse node from navigation tree
		objNavigationTree.Collapse iPath
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully collapsed [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]  from navigation tree","","","","DONOTSYNC","") 
		End If					
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Double Click node from navigation tree
	Case "DoubleClick"
		Dim intX, intY, intWidth, intHeight		
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to double click node [" & Cstr(sNodeName) & "] of navigation tree as specified node does not exist in navigation tree.","","","","","")
			Call Fn_ExitTest()
		End If
		
		objNavigationTree.Select iPath
		wait GBL_MIN_MICRO_TIMEOUT
		Call Fn_CommonUtil_KeyBoardOperation("SendKeys", "{ENTER}")
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to double click node [ " & CStr(sNodeName) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully double click [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]  from navigation tree","","","","DONOTSYNC","") 
		End If				 
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB menu
	Case "PopupMenuSelect", "MultiselectPopupMenu"
		If sAction = "PopupMenuSelect" Then
			
			'Retrive popup menu
			sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_NavigationTreeOperations","",sPopupMenu)
			
			If sPopupMenu = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu of node [" & Cstr(sNodeName) & "]" ,"","","","","")
				Call Fn_ExitTest()
			End If
			
			If sPopupMenu="Send To:Structure Manager" Then
				GBL_SENDTOPSM_FROM_NAVTREE_FLAG=True
			End If
			
			'Build the Popup menu to be selected
			aPopupMenu = Split(sPopupMenu, ":",-1,1)
			'Retrive node path
			iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
			
			If iPath=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu of node [" & Cstr(sNodeName) & "] specified node does not exist in navigation tree","","","","","")
				Call Fn_ExitTest()
			End If
			
			'Selecting node from tree
			objNavigationTree.Select iPath
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
			
			'Opening context menu on selected node
			If Fn_UI_JavaTree_Operations("RAC_MyTc_NavigationTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"jtree_NavigationTree",iPath,"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu as fail to open context menu of node [ " & Cstr(sNodeName) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		
		Else
			sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_NavigationTreeOperations","",sPopupMenu)
			
			If sPopupMenu="Send To:Structure Manager" Then
				GBL_SENDTOPSM_FROM_NAVTREE_FLAG=True
			End If
			
			aPopupMenu = Split(sPopupMenu, ":",-1,1)
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Multiselect", sNodeName,""
			
			iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, Split(sNodeName,"^")(0) , "~", "@")
			
			'Opening context menu on selected node
			If Fn_UI_JavaTree_Operations("RAC_MyTc_NavigationTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"jtree_NavigationTree",iPath,"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu as fail to open context menu of node [ " & Cstr(sNodeName) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		
					
			
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
		
		If Err.number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform context menu operation with error number [ " & Err.Number & " ] and error description [" & Err.Description & "]","","","","","")
			Call Fn_ExitTest()
		End If
		sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
		If sObjectTypeName="" Then
			sObjectTypeName="node"
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform RMB menu [ " & Cstr(sPopupMenu) & " ] operation of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from navigation tree","","","","DONOTSYNC","")
		
		If GBL_SENDTOPSM_FROM_NAVTREE_FLAG=True Then
			Call Fn_HandlePreferenceManagerError(sNodeName)
			GBL_SENDTOPSM_FROM_NAVTREE_FLAG=False
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB menu
	Case "PopupMenuSelect_Ext"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_NavigationTreeOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to get menu value from XML for XML tag [" & Parameter("sPopupMenu") & "]" ,"","","","","")
			Call Fn_ExitTest()
		End If
		
		If sPopupMenu="Send To:Structure Manager" Then
			GBL_SENDTOPSM_FROM_NAVTREE_FLAG=True
		End If
		
		'Build the Popup menu to be selected
		aPopupMenu = Split(sPopupMenu, ":",-1,1)
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu of node [" & Cstr(sNodeName) & "] specified node does not exist in navigation tree","","","","","")
			Call Fn_ExitTest()
		Else
			'Selecting node from tree
			objNavigationTree.Select iPath
		End If
				
		'Opening context menu on selected node
		If Fn_UI_JavaTree_Operations("RAC_MyTc_NavigationTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"jtree_NavigationTree",iPath,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu as fail to open context menu of node [ " & Cstr(sNodeName) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		'selecting RMB menu	
		sPopupMenu = Fn_UI_JavaMenu_Operations("RAC_MyTc_NavigationTreeOperations","Select",objMyTeamcenterWindow,sPopupMenu)				
		If sPopupMenu=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform context menu operation with error number [ " & Err.Number & " ] and error description [" & Err.Description & "]","","","","","")
			Call Fn_ExitTest()
		End If
		sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
		If sObjectTypeName="" Then
			sObjectTypeName="node"
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform RMB menu [ " & Cstr(sPopupMenu) & " ] operation of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from navigation tree","","","","DONOTSYNC","")
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		If GBL_SENDTOPSM_FROM_NAVTREE_FLAG=True Then
			Call Fn_HandlePreferenceManagerError(sNodeName)
			GBL_SENDTOPSM_FROM_NAVTREE_FLAG=False
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
	 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 'Checking existance of RMB menu by selecting multiple nodes
	Case "MultiSelectContextMenuExist"
		aNodeName = split(sNodeName,"^",-1,1)
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_NavigationTreeOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to get menu value from XML for XML tag [" & Parameter("sPopupMenu") & "]" ,"","","","","")
			Call Fn_ExitTest()
		End If
		
		'Multiselecting nodes
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Multiselect",sNodeName,""
								
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")		
		'Open context menu
		If Fn_UI_JavaTree_Operations("RAC_MyTc_NavigationTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"jtree_NavigationTree",iPath,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail  to Open Context Menu window","","","","","")
			Call Fn_ExitTest()
		End If
		
		'chekcing existance of RMB menu
		If objContextMenu.GetItemProperty (sPopupMenu,"Exists") = True Then
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified RMB menu [ " & Cstr(sPopupMenu) & " ] appears after multiselecting nodes from navigation tree","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] does not appears after multiselecting nodes from navigation tree","","","","","")
			Call Fn_ExitTest()
		End If				
	 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 'Case to check existance of node of navigation tree
	 Case "Exist","VerifyExist","VerifyNonExist"
		bFlag = True
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		
		If iPath=False Then
			bFlag = False
		Else
			aNodeName = split(Replace(iPath,"#",""),":")
			Set objCurrentNode = objNavigationTree.Object
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
				DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] does not exist","","","","","")
				Call Fn_ExitTest()
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)	
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
				DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)		
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)		
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
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to check existance of RMB menu
	Case "PopupMenuExist"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_NavigationTreeOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] due to invalid RMB menu [ " & Cstr(Parameter("sPopupMenu")) & "]","","","","","")
			Call Fn_ExitTest()
		End If

		'Build the Popup menu to be selected
		aPopupMenu = Split(sPopupMenu, ":",-1,1)
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] of node [" & sNodeName & "] in navigation tree as the specified node does not exist","","","","","")
			Call Fn_ExitTest()
		Else
			'Selecting node from tree
			objNavigationTree.Select iPath
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If		
		'Open context menu
		If Fn_UI_JavaTree_Operations("RAC_MyTc_NavigationTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"navigation tree",iPath,"","")=False Then
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
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] does not exist for node [ " & Cstr(sNodeName) & " ] in navigation tree","","","","","")
				Call Fn_ExitTest()						
		End Select
		'Checking existance of RMB menu
		If objContextMenu.GetItemProperty (sPopupMenu,"Exists") = True Then
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If	
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Sucessfully verified RMB menu [ " & Cstr(sPopupMenu) & " ] exist for [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] in navigation tree","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] does not exist for node [ " & Cstr(sNodeName) & " ] in navigation tree","","","","","")
			Call Fn_ExitTest()	
		End If
		Call Fn_CommonUtil_KeyBoardOperation("SendKeys", "{ESC}")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to check RMB menu state
	Case "PopupMenuEnabled"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_NavigationTreeOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as Context menu does not exist","","","","","")
			Call Fn_ExitTest()	
		End If

		'Build the Popup menu to be selected
		aPopupMenu = Split(sPopupMenu, ":",-1,1)
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform popup menu enabled action on Node [" & Cstr(sNodeName) & "] of navigation tree as the specified node does not exist.","","","","","")
			Call Fn_ExitTest()
		Else
			'Selecting node from tree
			objNavigationTree.Select iPath
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)			
		End If

		'Open context menu
		If Fn_UI_JavaTree_Operations("RAC_MyTc_NavigationTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"navigation tree",iPath,"","")=False Then
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
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] does not exist for node [ " & Cstr(sNodeName) & " ] in navigation tree","","","","","")
				Call Fn_ExitTest()
		End Select
		
		'Checking existance of RMB menu
		If objContextMenu.GetItemProperty (sPopupMenu,"Enabled") = True Then
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)	
			sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Sucessfully verified RMB menu [ " & Cstr(sPopupMenu) & " ] enabled for [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] in navigation tree","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as RMB menu [ " & Cstr(sPopupMenu) & " ] disabled for node [ " & Cstr(sNodeName) & " ] in navigation tree","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'getting selected node path
	Case "GetSelectedNodePath"
		If objNavigationTree.Object.getSelectionCount() > 0 Then
			Set objSelection = objNavigationTree.Object.getSelection()
			iSelectionCount =objNavigationTree.Object.getSelectionCount()
			For iCounter = 0 to iSelectionCount - 1 
				sReturn = ""
				Set objCurrentNode = objSelection.mic_arr_get(iCounter)
				If GBL_HP_QTP_PRODUCTNAME=Environment.Value("ProductName") Then
					If IsObject(objCurrentNode) then
						sReturn = objCurrentNode.getData().toString()
						Do while IsObject(objCurrentNode.getParentItem())
							Set objCurrentNode = objCurrentNode.getParentItem()
							sReturn = objCurrentNode.getData().toString() & "~" & sReturn 
						Loop
					End If
				Else
					If not objCurrentNode is Nothing then
						sReturn = objCurrentNode.getData().toString()
						Do while lcase(typename(objCurrentNode.getParentItem())) <> "nothing"
							Set objCurrentNode = objCurrentNode.getParentItem()
							sReturn = objCurrentNode.getData().toString() & "~" & sReturn 
						Loop
					End If
				End If				
				Set objCurrentNode = Nothing
				
				IF iCounter=0 Then
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
					DataTable.SetCurrentRow 1		
					DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
					DataTable.Value("ReusableActionWordReturnValue","Global")= sReturn
				Else
					DataTable.Value("ReusableActionWordReturnValue","Global")= DataTable.Value("ReusableActionWordReturnValue","Global") & "^" & sReturn
				End IF	
			Next
			Set objSelection =Nothing
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "GetIndex"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= Cstr(Fn_RAC_GetJavaTreeNodeIndex(objNavigationTree, sNodeName,"",""))
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "GetFirstNodeName"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= objNavigationTree.Object.getItem(0).getData().toString()
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Retruns Child of any given Name in the tree
	'sNodeName := Parent folder path ^ Node name
	Case "GetChildrenByName"
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Expand",aNodeName(0),""
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"		
		For iCounter=0 to Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
			If instr(1,GBL_JAVATREE_CURRENTNODE_OBJECT.getItem(iCounter).getData().toString(),aNodeName(1)) Then						
				DataTable.Value("ReusableActionWordReturnValue","Global")= GBL_JAVATREE_CURRENTNODE_OBJECT.getItem(iCounter).getData().toString()
				Exit For
			End If
		Next		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Expand node from navigation tree
	Case "ExpandAndSelect","ExpandAndSelectExt"		
		sNode = Split(sNodeName,"~",-1,1)								
		For iCount = 0 To Ubound(sNode)-1
			If sNode(iCount) <> "" Then
				If iCount = 0 Then
					If sAction="ExpandAndSelectExt" Then
						sNodePath =sNode(0)
					Else
						sNodePath = "Home"					
					End If
				Else
					sNodePath = sNodePath &"~"& sNode(iCount)
				End If				
				'Retrive node path
				iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodePath , "~", "@")
				If iPath=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand and select Node [" + sNodePath + "] of navigation tree as the specified node does not exist in navigation tree.","","","","","")
					Call Fn_ExitTest()
				Else
					'Expanding node from navigation tree
					objNavigationTree.Expand iPath
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Wait 1
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				End If
			End If
		Next		
		'Selecting node from navigation tree
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [" & Cstr(sNodeName) & "] of navigation tree as the specified node does not exist in navigation tree","","","","","")
			Call Fn_ExitTest()
		Else
			'Expanding node from navigation tree
			objNavigationTree.Select iPath
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)			
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)	
		
		sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
		If sObjectTypeName="" Then
			sObjectTypeName="node"
		End If		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from navigation tree","","","","DONOTSYNC","")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify of specific type dataset exist under node	
	Case "VerifyDatasetOfSpecificTypeExist", "VerifyDatasetOfSpecificTypeNotExist"
		bFlag = False
		'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"VerifyDatasetOfSpecificTypeExist","Item1~Item1Revision^Dataset1","DatasetType"
		'sNodeName :- Full path of node under user wants to check existamce of dataset ^ Dataset name
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",aNodeName(0),""
		DataTable.SetCurrentRow 1		
		If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		sDatasetType = sPopupMenu
		If sPopupMenu <> "" Then
			sPopupMenu = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewDatasetValues_APL",sPopupMenu,"")
		End If
		
		If sPopupMenu="DirectModel" Then
			
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"GetSelectedNodePath",aNodeName(0),""
			sCurrentlySelectedNode=DataTable.Value("ReusableActionWordReturnValue","Global")
			
			If Trim(sCurrentlySelectedNode)=Trim(aNodeName(0)) Then
				If objNavigationTree.Object.getItem(0).getData().toString()="Home" Then
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select","Home",""
				End If								
			End If
			
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",aNodeName(0),""
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewRefresh"
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			
			If sAction = "VerifyDatasetOfSpecificTypeExist" Then
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Deselect",aNodeName(0),""
			    Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			    LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",aNodeName(0),""
			    Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			    LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click", "Refresh", "",""
			    Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End If
			
		End If 
		
		iInstanceHandler=1		
		iCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
		sChildNodePath=aNodeName(0) & "~" & aNodeName(1)		
		For iCounter=0 to iCount-1
			sNodePath = sChildNodePath & "@" & iInstanceHandler
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				If sPopupMenu="DirectModel" Then
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",aNodeName(0),""
					LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewRefresh"
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""				
					DataTable.SetCurrentRow 1
				End If
				
				If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" and sAction = "VerifyDatasetOfSpecificTypeExist" Then					
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as dataset [ " & Cstr(aNodeName(1)) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If				
			End If
			'sPopupMenu : - For this specific case use this parameter to pass dataset type
			If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(sPopupMenu) Then
				bFlag = True
				Exit For
			End If
			iInstanceHandler=iInstanceHandler+1
		Next
		If bFlag = False Then
			If sAction = "VerifyDatasetOfSpecificTypeExist"  Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify that dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
			End If
		End If
		
		If bFlag = True Then
			If sAction = "VerifyDatasetOfSpecificTypeNotExist"  Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify that dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
			End If
		End If
		
		If sPopupMenu="DirectModel" Then
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",sCurrentlySelectedNode,""
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify of specific type dataset exist under node	
	Case "VerifyFormOfSpecificTypeExist"
		bFlag = False
		'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"VerifyDatasetOfSpecificTypeExist","Item1~Item1Revision^Dataset1","DatasetType"
		'sNodeName :- Full path of node under user wants to check existamce of form ^ Form name
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",aNodeName(0),""
		DataTable.SetCurrentRow 1		
		If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		sDatasetType = sPopupMenu
		If sPopupMenu <> "" Then
			sPopupMenu = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewFormValues_APL",sPopupMenu,"")
		End If
		iInstanceHandler=1		
		iCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
		sChildNodePath=aNodeName(0) & "~" & aNodeName(1)		
		For iCounter=0 to iCount-1
			sNodePath = sChildNodePath & "@" & iInstanceHandler
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as form [ " & Cstr(aNodeName(1)) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			'sPopupMenu : - For this specific case use this parameter to pass dataset type
			If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(sPopupMenu) Then
				bFlag = True
				Exit For
			End If
			iInstanceHandler=iInstanceHandler+1
		Next
		If bFlag = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify that form [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified form [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify of specific type dataset exist under node	
	Case "VerifyFormOfSpecificTypeNotExist"
		bFlag = False
		'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"VerifyDatasetOfSpecificTypeExist","Item1~Item1Revision^Dataset1","DatasetType"
		'sNodeName :- Full path of node under user wants to check existamce of form ^ Form name
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",aNodeName(0),""
		DataTable.SetCurrentRow 1		
		If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		sDatasetType = sPopupMenu
		If sPopupMenu <> "" Then
			sPopupMenu = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewFormValues_APL",sPopupMenu,"")
		End If
		iInstanceHandler=1		
		iCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
		sChildNodePath=aNodeName(0) & "~" & aNodeName(1)		
		For iCounter=0 to iCount-1
			sNodePath = sChildNodePath & "@" & iInstanceHandler
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				bFlag = False
				Exit For
			End If
			'sPopupMenu : - For this specific case use this parameter to pass dataset type
			If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(sPopupMenu) Then
				bFlag = True
				Exit For
			Else
				bFlag = False
			End If
			iInstanceHandler=iInstanceHandler+1
		Next
		If bFlag = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified form [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Veriication Fail as form [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
			Call Fn_ExitTest()
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Expand node from navigation tree
	Case "ExpandAll"
		sNode = Split(sNodeName,"~",-1,1)								
		For iCount = 0 To Ubound(sNode)-1
			If sNode(iCount) <> "" Then
				If iCount = 0 Then
					sNodePath =sNode(0)
				Else
					sNodePath = sNodePath &"~"& sNode(iCount)
				End If				
				'Retrive node path
				iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodePath , "~", "@")
				If iPath=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand and select Node [" + sNodePath + "] of navigation tree as the specified node does not exist in navigation tree.","","","","","")
					Call Fn_ExitTest()
				Else
					'Expanding node from navigation tree
					objNavigationTree.Expand iPath
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Wait 1
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				End If
			End If
		Next		
		'Selecting node from navigation tree
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [" & Cstr(sNodeName) & "] of navigation tree as the specified node does not exist in navigation tree","","","","","")
			Call Fn_ExitTest()
		Else
			'Expanding node from navigation tree
			objNavigationTree.Expand iPath
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Wait 1
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)	
		
		sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
		If sObjectTypeName="" Then
			sObjectTypeName="node"
		End If		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully expanded [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from navigation tree","","","","DONOTSYNC","")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to sselect of specific type dataset exist under node	
	Case "SelectDatasetOfSpecificType","ExpandDatasetOfSpecificType","GetInstanceNumberOfSpecificDatasetType"
		bFlag = False
		'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"SelectDatasetOfSpecificType","Item1~Item1Revision^Dataset1","DatasetType"
		'sNodeName :- Full path of node under user wants to check existamce of dataset ^ Dataset name
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",aNodeName(0),""
		DataTable.SetCurrentRow 1		
		If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		sDatasetType = sPopupMenu
		If sPopupMenu <> "" Then
			sPopupMenu = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewDatasetValues_APL",sPopupMenu,"")
		End If
		iInstanceHandler=1		
		iCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
		sChildNodePath=aNodeName(0) & "~" & aNodeName(1)		
		For iCounter=0 to iCount-1
			sNodePath = sChildNodePath & "@" & iInstanceHandler
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as dataste [ " & Cstr(aNodeName(1)) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			'sPopupMenu : - For this specific case use this parameter to pass dataset type
			If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(sPopupMenu) Then
				If sAction="SelectDatasetOfSpecificType" Then
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",sNodePath,""
				ElseIf sAction="ExpandDatasetOfSpecificType" Then
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Expand",sNodePath,""
				ElseIf sAction="GetInstanceNumberOfSpecificDatasetType" Then
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
					DataTable.SetCurrentRow 1		
					DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
					DataTable.Value("ReusableActionWordReturnValue","Global")= Cstr(iInstanceHandler)
				End If
				bFlag = True
				Exit For
			End If
			iInstanceHandler=iInstanceHandler+1
		Next		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "VerifyLastModifiedTimeGreaterThanCreatedTime"		
		'Retrive node path
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [ " & Cstr(sNodeName) & " ] of navigation tree as specified node does not exist","","","","","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as last modified time is not greater ( current time ) than created time of object [ " & Cstr(sNodeName) & " ] * This node is not exist in navigation tree","","","","","")
			Call Fn_ExitTest()
		End If			
		
		If DateDiff("s",GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getProperty("creation_date"),GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getProperty("last_mod_date"))>0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified last modified time is greater ( current time ) than created time of object [ " & Cstr(sNodeName) & " ]","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as last modified time is not greater ( current time ) than created time of object [ " & Cstr(sNodeName) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
	''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	Case "SelectFormOfSpecificTypeExist"
	     bFlag = False
		'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"VerifyDatasetOfSpecificTypeExist","Item1~Item1Revision^Dataset1","DatasetType"
		'sNodeName :- Full path of node under user wants to check existamce of form ^ Form name
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",aNodeName(0),""
		DataTable.SetCurrentRow 1		
		If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		sDatasetType = sPopupMenu
		If sPopupMenu <> "" Then
			sPopupMenu = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewFormValues_APL",sPopupMenu,"")
		End If
		iInstanceHandler=1		
		iCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
		sChildNodePath=aNodeName(0) & "~" & aNodeName(1)		
		For iCounter=0 to iCount-1
			sNodePath = sChildNodePath & "@" & iInstanceHandler
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as form [ " & Cstr(aNodeName(1)) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			'sPopupMenu : - For this specific case use this parameter to pass dataset type
			If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(sPopupMenu) Then
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",sNodePath,""
					bFlag = True
				    Exit For
			End If
			iInstanceHandler=iInstanceHandler+1
		Next	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -					
	Case "GetLatestRevision", "VerifyLatestRevision"
		
		'Select the node for whom the child revision are to be retrieved
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Select", sNodeName,""
		
		'Get all the revision list for this node
		sTempValue = GBL_JAVATREE_CURRENTNODE_OBJECT.getdata().getcomponent().getProperty("revision_list")
		
		'Store details in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
		'sTempValue = Split(sTempValue, sPopupmenu & "/")
		'DataTable.Value("ReusableActionWordReturnValue","Global") = sPopupmenu & "/" & sTempValue(Ubound(sTempValue))
		
		sTempValue = Split(sTempValue, ",")
		If sAction = "VerifyLatestRevision" Then
			If Trim(Ubound(sTempValue)) = sPopupmenu Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified latest revision as [ " & sPopupmenu & " ]","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Failed to verify latest revision as [ " & sPopupmenu & " ]","","","","DONOTSYNC","")
				Call Fn_ExitTest()
			End If
		Else
			DataTable.Value("ReusableActionWordReturnValue","Global") = Ubound(sTempValue)
		End If
		
	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	Case "GetNativeProperties"
	
		'Select the node for whom properties are to be retrieved
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Select", sNodeName,""
		
		aPropertyNames = Split(dictProperties("PropertyNames"), "~")
		dictProperties("PropertyValues") = ""
		For iCounter = 0 To Ubound(aPropertyNames) Step 1
			If iCounter = 0 Then
				dictProperties("PropertyValues") = Trim(Cstr(GBL_JAVATREE_CURRENTNODE_OBJECT.getdata().getcomponent().getProperty(Cstr(aPropertyNames(iCounter)))))
			Else
				dictProperties("PropertyValues") = dictProperties("PropertyValues") & "~" & Trim(Cstr(GBL_JAVATREE_CURRENTNODE_OBJECT.getdata().getcomponent().getProperty(Cstr(aPropertyNames(iCounter)))))
			End If
		Next
		
		'Store has migration form property value
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global") = Trim(Cstr(GBL_JAVATREE_CURRENTNODE_OBJECT.getdata().getcomponent().getProperty("Ng5_rHasMigration")))
	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify of specific type dataset exist under node	
	Case "VerifyDatasetOfSpecificTypeCount"
		bFlag = False
		iActualOccuranceCount = 0
		
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",aNodeName(0),""
		DataTable.SetCurrentRow 1		
		If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		sDatasetType = Split(sPopupMenu, "^")(0)
		iExpectedOccuranceCount = Cint(Split(sPopupMenu, "^")(1))
		
		If sPopupMenu <> "" Then
			sPopupMenu = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewDatasetValues_APL",Split(sPopupMenu, "^")(0),"")
		End If
		
		If sPopupMenu="DirectModel" Then
			
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"GetSelectedNodePath",aNodeName(0),""
			sCurrentlySelectedNode=DataTable.Value("ReusableActionWordReturnValue","Global")
			
			If Trim(sCurrentlySelectedNode)=Trim(aNodeName(0)) Then			
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select","Home",""	
			End If
			
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",aNodeName(0),""
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewRefresh"
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If 
		
		iInstanceHandler=1		
		iCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
		sChildNodePath=aNodeName(0) & "~" & aNodeName(1)		
		For iCounter=0 to iCount-1
			sNodePath = sChildNodePath & "@" & iInstanceHandler
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				If sPopupMenu="DirectModel" Then
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",aNodeName(0),""
					LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewRefresh"
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""				
					DataTable.SetCurrentRow 1
				End If
				If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as dataste [ " & Cstr(aNodeName(1)) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			'sPopupMenu : - For this specific case use this parameter to pass dataset type
			If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(sPopupMenu) Then
				iActualOccuranceCount = iActualOccuranceCount + 1
				If iActualOccuranceCount = iExpectedOccuranceCount Then
					Exit For
				End If
			End If
			iInstanceHandler=iInstanceHandler+1
		Next
		If iActualOccuranceCount <> iExpectedOccuranceCount Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify that dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist for [ " & Cstr(iActualOccuranceCount) & " ] times while the expected occurrence count was [" & Cstr(iExpectedOccuranceCount) & "]","","","","DONOTSYNC","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist for [ " & Cstr(iActualOccuranceCount) & " ] times while the expected occurrence count was [" & Cstr(iExpectedOccuranceCount) & "]","","","","DONOTSYNC","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Select all child node exit under specific node
	Case "SelectAllChildNode"
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree,sNodeName , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail select all child node under node [ " & Cstr(sNodeName) & " ] of navigation tree as node [" & Cstr(sNodeName) & "] does not exist","","","","","")
			Call Fn_ExitTest()
		End If		
		For iCounter=0 to Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
			If iCounter=0 Then
				objNavigationTree.Select iPath & "~#" & iCounter
			Else
				objNavigationTree.ExtendSelect iPath & "~#" & iCounter				
			End If				
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Select all child node exit under specific node
	Case "SelectNodeAndVerifyProperty","VerifyProperty"
		If sAction="SelectNodeAndVerifyProperty" Then
			'Retrive node path
			iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, sNodeName , "~", "@")
			If iPath=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [ " & Cstr(sNodeName) & " ] of navigation tree as specified node does not exist","","","","","")
				Call Fn_ExitTest()
			End If			
			'Selecting node from tree
			objNavigationTree.Select iPath			
			If Err.number <> 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ]","","","","","") 
				Call Fn_ExitTest()
			Else
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
				If sObjectTypeName="" Then
					sObjectTypeName="Node"
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from navigation tree","","","","DONOTSYNC","")
			End If
		End If
		
		aPopupMenu=Split(sPopupMenu,"^")
		'aPopupMenu(0) = Property Name
		'aPopupMenu(1) = Property Value
		
		JavaWindow("jwnd_MyTeamcenter").JavaStaticText("jstx_SummaryTabText").SetTOProperty "label",aPopupMenu(0) & ":"
		JavaWindow("jwnd_MyTeamcenter").JavaEdit("jedt_SummaryTabEdit1").SetTOProperty "attached text",aPopupMenu(0) & ":"
		
		If JavaWindow("jwnd_MyTeamcenter").JavaEdit("jedt_SummaryTabEdit").Exist(0) Then
			If Fn_UI_JavaEdit_Operations("RAC_MyTc_NavigationTreeOperations", "gettext",  JavaWindow("jwnd_MyTeamcenter"),"jedt_SummaryTabEdit", "" )=aPopupMenu(1) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aPopupMenu(0)) & " ] property contain value [ " & Cstr(aPopupMenu(1)) & " ]","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aPopupMenu(0)) & " ] property does not contain value [ " & Cstr(aPopupMenu(1)) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		ElseIf JavaWindow("jwnd_MyTeamcenter").JavaEdit("jedt_SummaryTabEdit1").Exist(0) Then
			If Fn_UI_JavaEdit_Operations("RAC_MyTc_NavigationTreeOperations", "gettext",  JavaWindow("jwnd_MyTeamcenter"),"jedt_SummaryTabEdit1", "" )=aPopupMenu(1) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aPopupMenu(0)) & " ] property contain value [ " & Cstr(aPopupMenu(1)) & " ]","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aPopupMenu(0)) & " ] property does not contain value [ " & Cstr(aPopupMenu(1)) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aPopupMenu(0)) & " ] property does not exist\available on summury page","","","","","")
			Call Fn_ExitTest()
		End IF	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Retruns Name of item by given Name in the tree
	'sNodeName := Parent folder path ^ Node ID
	'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "GetNameByID","Home~AutomatedTest~TestFolder~4501297-Asm Cockpit^4501297/AB01",""
	'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "GetNameByID","Home~AutomatedTest~TestFolder^4501297",""
	
	Case "GetNameByID"
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Expand",aNodeName(0),""
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"		
		For iCounter=0 to Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
			If instr(1,GBL_JAVATREE_CURRENTNODE_OBJECT.getItem(iCounter).getData().toString(),aNodeName(1)) Then						
				aPopupMenu= Split(GBL_JAVATREE_CURRENTNODE_OBJECT.getItem(iCounter).getData().toString(),"-")
				DataTable.Value("ReusableActionWordReturnValue","Global")=aPopupMenu(1)
				Exit For
			End If
		Next		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -				
	Case "VerifyTwoIDHasUniqueValue"
		aPopupMenu=Split(sPopupMenu,"~")
		If Cstr(aPopupMenu(0))=Cstr(aPopupMenu(1)) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as both Item id has same number","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified Item Id [ " & Cstr(aPopupMenu(0)) & " ] and Item Is [ " & Cstr(aPopupMenu(1)) & " ] has unique values","","","","DONOTSYNC","")	
		End If	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Retruns Name of item by given Name in the tree
	'sNodeName := Parent folder path ^ Node ID
	'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "GetAllInfoByID","Home~AutomatedTest~TestFolder~4501297-Asm Cockpit^4501297/AB01",""
	'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "GetAllInfoByID","Home~AutomatedTest~TestFolder^4501297",""
	
	Case "GetAllInfoByID"
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Expand",aNodeName(0),""
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"		
		For iCounter=0 to Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
			If instr(1,GBL_JAVATREE_CURRENTNODE_OBJECT.getItem(iCounter).getData().toString(),aNodeName(1)) Then						
				aPopupMenu= Split(GBL_JAVATREE_CURRENTNODE_OBJECT.getItem(iCounter).getData().toString(),"/")
				If Ubound(aPopupMenu)=0 Then
					aPopupMenu= Split(GBL_JAVATREE_CURRENTNODE_OBJECT.getItem(iCounter).getData().toString(),"-",2)
					DataTable.Value("ReusableActionWordReturnValue","Global")=aPopupMenu(1)
				Else
					DataTable.Value("ReusableActionWordReturnValue","Global")=aPopupMenu(1)
				End If
				Exit For
			End If
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Retruns first node Name from tree
	'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "GetFirstNodeName","",""	
	Case "GetFirstNodeObjectName"		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"		
		'For iCounter=0 to Cint(objNavigationTree.Object.getItemCount())-1
		'	If instr(1,objNavigationTree.Object.getItem(0).getData().toString(),aNodeName(1)) Then						
				aPopupMenu= Split(objNavigationTree.Object.getItem(0).getData().toString(),"-",2)
				DataTable.Value("ReusableActionWordReturnValue","Global")=aPopupMenu(1)
				'Exit For
		'	End If
		'Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to multi select nodes from navigation tree
	Case "MultiselectPopupMenu_Ext"
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MyTc_NavigationTreeOperations","",sPopupMenu)
		
		If sPopupMenu="Send To:Structure Manager" Then
			GBL_SENDTOPSM_FROM_NAVTREE_FLAG=True
		End If
		
		aNodeName=Split(sNodeName,"^")
		For iCounter=0 To UBound(aNodeName)
			iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, aNodeName(iCounter) , "~", "@")
			If iPath=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Multi Select Node [" & Cstr(sNodeName) & "] of navigation tree as node [" & Cstr(aNodeName(iCounter)) & "] does not exist","","","","","")
				Call Fn_ExitTest()
			Else
				'Multiselecting items
				If iCounter=0 Then
					objNavigationTree.Select iPath				
				Else
					objNavigationTree.ExtendSelect iPath						
				End If
				
				If Err.Number <> 0 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multi select nodes [ " & CStr(aNodeName(iCounter)) & " ]","","","","","") 
					Call Fn_ExitTest()
				Else
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)
					sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
					If sObjectTypeName="" Then
						sObjectTypeName="nodes"
					End If
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(aNodeName(iCounter)) & " ]","","","","DONOTSYNC","") 
				End If
			End If
		Next
		iPath = Fn_RAC_GetJavaTreeNodePath(objNavigationTree, aNodeName(UBound(aNodeName)) , "~", "@")
		If iPath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [ " & Cstr(aNodeName(UBound(aNodeName))) & " ] of navigation tree as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If	
		
		'Opening context menu on selected node
		If Fn_UI_JavaTree_Operations("RAC_MyTc_NavigationTreeOperations","OpenContextMenu",objMyTeamcenterWindow,"jtree_NavigationTree",iPath,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu as fail to open context menu of node [ " & Cstr(aNodeName(UBound(aNodeName))) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		'selecting RMB menu	
		sPopupMenu = Fn_UI_JavaMenu_Operations("RAC_MyTc_NavigationTreeOperations","Select",objMyTeamcenterWindow,sPopupMenu)				
		If sPopupMenu=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform context menu operation with error number [ " & Err.Number & " ] and error description [" & Err.Description & "]","","","","","")
			Call Fn_ExitTest()
		End If
		sObjectTypeName=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
		If sObjectTypeName="" Then
			sObjectTypeName="node"
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform RMB menu [ " & Cstr(sPopupMenu) & " ] operation of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from navigation tree","","","","DONOTSYNC","")
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter navigation tree Node Operations",sAction,"Node name",sNodeName)	
	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	
	'Case to DoubleClick of specific type dataset exist under node	
	Case "DoubleClickDatasetOfSpecificType"
		bFlag = False
		'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"SelectDatasetOfSpecificType","Item1~Item1Revision^Dataset1","DatasetType"
		'sNodeName :- Full path of node under user wants to check existamce of dataset ^ Dataset name
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",aNodeName(0),""
		DataTable.SetCurrentRow 1		
		If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		sDatasetType = sPopupMenu
		If sPopupMenu <> "" Then
			sPopupMenu = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewDatasetValues_APL",sPopupMenu,"")
		End If
		iInstanceHandler=1		
		iCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
		sChildNodePath=aNodeName(0) & "~" & aNodeName(1)
		For iCounter=0 to iCount-1
			'sNodePath = sChildNodePath & "~" & GBL_JAVATREE_CURRENTNODE_OBJECT.getItem(iCounter).getData().toString() & "@" & iInstanceHandler
			sNodePath = sChildNodePath  & "@" & iInstanceHandler
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as dataste [ " & Cstr(aNodeName(1)) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			'sPopupMenu : - For this specific case use this parameter to pass dataset type
			If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(sPopupMenu) Then
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"DoubleClick",sNodePath,""				
				bFlag = True
				Exit For
			End If
			iInstanceHandler=iInstanceHandler+1
		Next
		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify of specific type dataset exist under node
    'It returns True or False value	
	Case "GetValueforDatasetOfSpecificTypeExistOrNot"
		bFlag = False
		'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"VerifyDatasetOfSpecificTypeExist","Item1~Item1Revision^Dataset1","DatasetType"
		'sNodeName :- Full path of node under user wants to check existamce of dataset ^ Dataset name
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",aNodeName(0),""
		DataTable.SetCurrentRow 1		
		If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		sDatasetType = sPopupMenu
		If sPopupMenu <> "" Then
			sPopupMenu = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewDatasetValues_APL",sPopupMenu,"")
		End If
		
		If sPopupMenu="DirectModel" Then
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"GetSelectedNodePath",aNodeName(0),""
			sCurrentlySelectedNode=DataTable.Value("ReusableActionWordReturnValue","Global")
			
			If Trim(sCurrentlySelectedNode)=Trim(aNodeName(0)) Then
				If objNavigationTree.Object.getItem(0).getData().toString()="Home" Then
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select","Home",""
				End If								
			End If
			
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",aNodeName(0),""
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewRefresh"
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Deselect",aNodeName(0),""
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",aNodeName(0),""
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click", "Refresh", "",""
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If 
		
		iInstanceHandler=1		
		iCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
		sChildNodePath=aNodeName(0) & "~" & aNodeName(1)		
		For iCounter=0 to iCount-1
			sNodePath = sChildNodePath & "@" & iInstanceHandler
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				If sPopupMenu="DirectModel" Then
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",aNodeName(0),""
					LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewRefresh"
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""				
					DataTable.SetCurrentRow 1
				End If
				
				If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" and sAction = "GetValueforDatasetOfSpecificTypeExistOrNot" Then					
					DataTable.SetCurrentRow 1		
					DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
					DataTable.Value("ReusableActionWordReturnValue","Global")=False
					DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
				End If				
			End If
			'sPopupMenu : - For this specific case use this parameter to pass dataset type
			If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(sPopupMenu) Then
				bFlag = True
				Exit For
			End If
			iInstanceHandler=iInstanceHandler+1
		Next
		If bFlag = False Then
			If sAction = "GetValueforDatasetOfSpecificTypeExistOrNot"  Then
				DataTable.SetCurrentRow 1		
					DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
					DataTable.Value("ReusableActionWordReturnValue","Global")=False
					DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
			Else
				DataTable.SetCurrentRow 1		
					DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
					DataTable.Value("ReusableActionWordReturnValue","Global")=True
					DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
			End If
		End If
		
		If sPopupMenu="DirectModel" Then
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Select",sCurrentlySelectedNode,""
		End If	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	Case "GetInstanceNumberByCreationTime"
		bFlag = False
		'LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"GetInstanceNumberByCreationTime","Item1~Item1Revision^Dataset1","DatasetType"
		'sNodeName :- Full path of node under user wants to check existamce of dataset ^ Dataset name
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",aNodeName(0),""
		DataTable.SetCurrentRow 1		
		If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		iInstanceHandler=1
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= Cstr(0)
		sTempValue=""
		
		iCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
		sChildNodePath=aNodeName(0) & "~" & aNodeName(1)		
		For iCounter=0 to iCount-1
			sNodePath = sChildNodePath & "@" & iInstanceHandler
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				Exit For
			End If
			If sTempValue="" Then
				sTempValue=GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getProperty("creation_date")
				aPropertyNames=1
			Else
				sTempValue=sTempValue & "~" & GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getProperty("creation_date")
				aPropertyNames=aPropertyNames & "~" & aPropertyNames+1
			End If	
			iInstanceHandler=iInstanceHandler+1			
		Next
		
		iInstanceHandler=aPropertyNames	
		
		If sTempValue<>"" Then			
			aNodeName=Split(sTempValue,"~")
			aPopupMenu =Split(iInstanceHandler,"~")
			
			For iCount = UBound(aNodeName) - 1 To 0 Step -1
				For iCounter= 0 to iCount
					  If aNodeName(iCounter)>aNodeName(iCounter+1) then 
						  sTempValue=aNodeName(iCounter+1) 
						  aNodeName(iCounter+1)=aNodeName(iCounter)
						  aNodeName(iCounter)=sTempValue
						  
						  sTempValue=aPopupMenu(iCounter+1) 
						  aPopupMenu(iCounter+1)=aPopupMenu(iCounter)
						  aPopupMenu(iCounter)=sTempValue
					  End if 
				Next 
			Next
					  
			For iCounter=0 to Ubound(aPopupMenu)
				DataTable.SetCurrentRow iCounter+1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_MyTc_NavigationTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= Cstr(aPopupMenu(iCounter))
			Next
		End If
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
Set objNavigationTree=Nothing
Set objContextMenu = Nothing
Set objMyTeamcenterWindow = Nothing

Function Fn_ExitTest()	
	'Releasing all objects
	Set objNavigationTree=Nothing
	Set objContextMenu = Nothing
	Set objMyTeamcenterWindow = Nothing
	ExitTest
End Function






Function Fn_HandlePreferenceManagerError(sNavTreeNodePath)
	Dim bErrorFlag
	Dim iErrorCounter
	Dim iNodePath
	
	bErrorFlag=False
	
	Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	
	For iErrorCounter = 1 To 3
		If JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaTable("jtbl_BOMTable").Exist(6) Then
			If Cint(JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaTable("jtbl_BOMTable").GetROProperty("rows"))=0 Then
				bErrorFlag=False
			Else
				Exit For				
			End If
		End If
	
		If bErrorFlag=False Then					
			If Window("wwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaDialog("jdlg_Error").Exist(1) Then
				Window("wwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaDialog("jdlg_Error").JavaButton("jbtn_OK").Click
				bErrorFlag=True
			ElseIf JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaDialog("jdlg_Error").Exist(1) Then
				JavaWindow("wwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaDialog("jdlg_Error").JavaButton("jbtn_OK").Click
				bErrorFlag=True
			End If
			If bErrorFlag=True Then
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations", oneIteration, "Select", "FileClose"
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "ExpandAndSelect",sNavTreeNodePath,""
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "SelectAndRefresh",sNavTreeNodePath,""
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				
				iNodePath = Fn_RAC_GetJavaTreeNodePath(JavaWindow("jwnd_MyTeamcenter").JavaTree("jtree_NavigationTree"), sNavTreeNodePath , "~", "@")
				
				'Opening context menu on selected node
				If Fn_UI_JavaTree_Operations("Fn_HandlePreferenceManagerError","OpenContextMenu",JavaWindow("jwnd_MyTeamcenter"),"jtree_NavigationTree",iNodePath,"","")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu as fail to open context menu of node [ " & Cstr(sNavTreeNodePath) & " ]","","","","","")
					ExitTest
				End If				
				wait 2
				JavaWindow("jwnd_MyTeamcenter").JavaMenu("label:=Send To").JavaMenu("label:=Structure Manager").Select
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
				bErrorFlag=False
			Else
				Exit For
			End If	
		Else
			Exit For
		End If
	Next
End Function

