'! @Name 			RAC_Search_SearchResultsTreeOperations
'! @Details 		This Action word to perform operations on search results tree
'! @InputParam1 	sAction 			: Action Name
'! @InputParam2 	sNodeName 			: Search results node path
'! @InputParam3 	sMenu 				: Popup menu tag name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			30 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Search\RAC_Search_SearchResultsTreeOperations","RAC_Search_SearchResultsTreeOperations",OneIteration,"Select","Problem Report~PR-00003", ""
'! @Example 		LoadAndRunAction "RAC_Search\RAC_Search_SearchResultsTreeOperations","RAC_Search_SearchResultsTreeOperations",OneIteration,"VerifyExist","Problem Report~PR-00003", ""
'! @Example 		LoadAndRunAction "RAC_Search\RAC_Search_SearchResultsTreeOperations","RAC_Search_SearchResultsTreeOperations",OneIteration,"VerifySingleResultDisplayedInSearchTree","", ""

Option Explicit
Err.Clear

'Declare variable
Dim sAction, sNodeName,sPopupMenu
Dim objDefaultWindow,objSearchResultTree,objNoAccessibleObjects
Dim objCurrentNode,objContextMenu
Dim sNodePath,sObjectTypeName,sDatasetType,sChildNodePath
Dim aNodeName,aPopupMenu,aNodePath, aTempValue
Dim iCounter,iInstanceHandler,iCount,iPath
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading input parameter values
sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sPopupMenu = Parameter("sPopupMenu")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Search_SearchResultsTreeOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Creating object of Teamcenter Default Window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","jwnd_SearchDefaultWindow","")
'Creating object of Search result tree
Set objSearchResultTree=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","jtree_SearchResultTree","")
'Creating object of No Accessible Object Window
Set objNoAccessibleObjects=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","jdlg_NoAccessibleObjects","")
'Select RMB menu
Set objContextMenu=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","wmnu_ContextMenu","")

'Checked Existance of Search Result Tree
If Fn_UI_Object_Operations("RAC_Search_SearchResultsTreeOperations","Exist",objSearchResultTree,"","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on Search result as search result tree does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Expand the root node of search result tree.
objSearchResultTree.Expand "#0"

'Retrive and append exact node name of first node
If sNodeName<>"" Then
	If InStr(1,sNodeName,"~") = 0  Then
		'Do nothihng
	Else
		aNodeName = Split(sNodeName,"~")
		aNodeName(0) = objSearchResultTree.Object.getItem(0).getData().toString()
		sNodeName = Join(aNodeName,"~")
	End If
Else
   sNodeName = objSearchResultTree.Object.getItem(0).getData().toString()
End If

'Capture business functionality start time	
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Search Results Operations",sAction,"Node Path",sNodeName)

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Clcik on load all button to load all results
If CBool(objDefaultWindow.JavaToolbar("jtlbr_SearchResultPaneToolBar").GetItemProperty("Load All","enabled")) Then
	Call Fn_UI_JavaToolbar_Operations("RAC_Search_SearchResultsTreeOperations", "Click", objDefaultWindow,"jtlbr_SearchResultPaneToolBar", "Load All", "", "", "")
	If objDefaultWindow.Dialog("text:=Info").Exist(2) Then
		objDefaultWindow.Dialog("text:=Info").WinButton("text:=OK").Click
		Call Fn_RAC_ReadyStatusSync(1)
	End If
End If

'checking existance of progree information
Call Fn_UI_Object_Operations("RAC_Search_SearchResultsTreeOperations","settoproperty",objNoAccessibleObjects,"","title","Progress Information")
While Fn_UI_Object_Operations("RAC_Search_SearchResultsTreeOperations","Exist",objNoAccessibleObjects,GBL_DEFAULT_MIN_TIMEOUT,"","")
	Wait 0,200
Wend

Select Case sAction
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select node by number from search results tree
	Case "SelectByNumber"
		sPopupMenu=sPopupMenu-1		
		sNodePath = Fn_RAC_GetJavaTreeNodePath(objSearchResultTree, sNodeName , "~", "@")		
		objSearchResultTree.Select sNodePath & "~#" & sPopupMenu
		sNodePath=sNodePath & "~" & sPopupMenu
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodePath) & " ] from search results due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Search Results Operations",sAction,"Node Path",sNodeName)
			If sObjectTypeName="" Then
				sObjectTypeName="Node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from search results tree","","","","DONOTSYNC","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select node from search results tree
	Case "Select"
		'Retrive node path
		sNodePath = Fn_RAC_GetJavaTreeNodePath(objSearchResultTree, sNodeName , "~", "@")
		If sNodePath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Node [ " & Cstr(sNodeName) & " ] of search results tree as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If			
		'Selecting node from tree
		objSearchResultTree.Select sNodePath			
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ] from search results due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Search Results Operations",sAction,"Node Path",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("SearchResultsTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="Node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from search results tree","","","","DONOTSYNC","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Expand node from navigation tree
	Case "Expand"	
		'Retrive node path
		sNodePath = Fn_RAC_GetJavaTreeNodePath(objSearchResultTree, sNodeName , "~", "@")
		If sNodePath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Expand Node [" & Cstr(sNodeName) & "] of search results tree as specified node does not exist in search results tree.","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Expanding node from search results tree
		objSearchResultTree.Expand sNodePath
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & CStr(sNodeName) & " ] from search results due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Search Results Operations",sAction,"Node Path",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("SearchResultsTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully expanded [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]","","","","DONOTSYNC","") 
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Double Click node from navigation tree
	Case "DoubleClick"
		sNodePath = Fn_RAC_GetJavaTreeNodePath(objSearchResultTree, sNodeName , "~", "@")
		
		If sNodePath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to double click node [ " & Cstr(sNodeName) & " ] of search results tree as specified node does not exist in search results tree.","","","","","")
			Call Fn_ExitTest()
		End If
		
		objSearchResultTree.Select sNodePath
		wait GBL_MIN_MICRO_TIMEOUT
		Call Fn_CommonUtil_KeyBoardOperation("SendKeys", "{ENTER}")
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to double click node [ " & CStr(sNodeName) & " ] from search results due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Search Results Operations",sAction,"Node Path",sNodeName)
			sObjectTypeName=Fn_RAC_GetTreeNodeType("SearchResultsTree","getobjecttypename")
			If sObjectTypeName="" Then
				sObjectTypeName="node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully double click [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]","","","","DONOTSYNC","") 
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 'Case to check existance of node of navigation tree
	 Case "Exist","VerifyExist","VerifyNonExist"
		bFlag = True
		sNodePath = Fn_RAC_GetJavaTreeNodePath(objSearchResultTree, sNodeName , "~", "@")
		
		If sNodePath=False Then
			bFlag = False
		Else
			aNodeName = split(Replace(sNodePath,"#",""),":")
			Set objCurrentNode = objSearchResultTree.Object
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
				DataTable.Value("ReusableActionWordName","Global")= "RAC_Search_SearchResultsTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] does not exist under search result tree","","","","","")
				Call Fn_ExitTest()
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Search Results Operations",sAction,"Node Path",sNodeName)
				If GBL_LOG_ADDITIONAL_INFORMATION<>"" Then
					sObjectTypeName=GBL_LOG_ADDITIONAL_INFORMATION
				Else
					sObjectTypeName="node"
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist under search result tree","","","","DONOTSYNC","")
			End If
		Else
			If sAction="Exist" Then
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				DataTable.SetCurrentRow 1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_Search_SearchResultsTreeOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Search Results Operations",sAction,"Node Path",sNodeName)
			ElseIf sAction="VerifyExist" Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Search Results Operations",sAction,"Node Path",sNodeName)
				sObjectTypeName=Fn_RAC_GetTreeNodeType("SearchResultsTree","getobjecttypename")
				If sObjectTypeName="" Then
					sObjectTypeName="node"
				End If	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] exist under search result tree","","","","DONOTSYNC","")
			ElseIf sAction="VerifyNonExist" Then
				If GBL_LOG_ADDITIONAL_INFORMATION<>"" Then
					sObjectTypeName=GBL_LOG_ADDITIONAL_INFORMATION
				Else
					sObjectTypeName="node"
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] exist under search result tree","","","","","")
				Call Fn_ExitTest()
			End If
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB menu
	Case "PopupMenuSelect"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_Search_SearchResultsTreeOperations","",sPopupMenu)

		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu of node [ " & Cstr(sNodeName) & " ] under search results tree" ,"","","","","")
			Call Fn_ExitTest()
		End If
		
		'Build the Popup menu to be selected
		aPopupMenu = Split(sPopupMenu,":",-1,1)
		'Retrive node path
		sNodePath = Fn_RAC_GetJavaTreeNodePath(objSearchResultTree, sNodeName , "~", "@")
		
		If sNodePath=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu of node [ " & Cstr(sNodeName) & " ] as specified node does not exist in search results tree","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Selecting node from tree
		objSearchResultTree.Select sNodePath
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)		
		
		'Opening context menu on selected node
		If Fn_UI_JavaTree_Operations("RAC_Search_SearchResultsTreeOperations","OpenContextMenu",objSearchResultTree,"",sNodePath,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu as fail to open context menu of node [ " & Cstr(sNodeName) & " ] under search results tree","","","","","")
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
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to build menu path of context menu of node [ " & Cstr(sNodeName) & " ]","","","","","")
				Call Fn_ExitTest()
		End Select
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		If Err.number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform context menu operation with error number [ " & Err.Number & " ] and error description [" & Err.Description & "] under search results tree","","","","","")
			Call Fn_ExitTest()
		End If
		
		sObjectTypeName=Fn_RAC_GetTreeNodeType("SearchResultsTree","getobjecttypename")
		If sObjectTypeName="" Then
			sObjectTypeName="node"
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform RMB menu [ " & Cstr(sPopupMenu) & " ] operation of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from search results tree","","","","","")
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Search Results Operations",sAction,"Node Path",sNodeName)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify of specific type dataset exist under node	
	Case "VerifyDatasetOfSpecificTypeExist"
		bFlag = False
		'LoadAndRunAction "RAC_Search\RAC_Search_SearchResultsTreeOperations","RAC_Search_SearchResultsTreeOperations",OneIteration,"VerifyDatasetOfSpecificTypeExist","Item1~Item1Revision^Dataset1","DatasetType"
		'sNodeName :- Full path of node under user wants to check existamce of dataset ^ Dataset name
		aNodeName=Split(sNodeName,"^")
		LoadAndRunAction "RAC_Search\RAC_Search_SearchResultsTreeOperations","RAC_Search_SearchResultsTreeOperations",OneIteration,"Exist",aNodeName(0),""
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
			LoadAndRunAction "RAC_Search\RAC_Search_SearchResultsTreeOperations","RAC_Search_SearchResultsTreeOperations",OneIteration,"Exist",sNodePath,""
			DataTable.SetCurrentRow 1		
			If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as dataste [ " & Cstr(aNodeName(1)) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","","")
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
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify that dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
		End If
		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 'Case to verify existence on multiple search results
	 Case "VerifyMultipleResultsDisplayedInSearchTree"
		
		If objSearchResultTree.GetROProperty("items count") > 2 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified multiple results displayed under search tree","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify multiple results displayed under search tree","","","","DONOTSYNC","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get name of the first node in search tree
	Case "GetFirstNodeName"
		'Retrive node path
		sNodePath = Trim(Cstr(objSearchResultTree.Object.getItem(0).getItem(0).getData().toString()))	

		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_Search_SearchResultsTreeOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= sNodePath
		Datatable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get name of the first node in search tree
	Case "GetFirstNodeIDAndName"
		'Retrive node path
		sNodePath=""
		Call Fn_Setup_ReporterFilter("DisableAll")
		sNodePath = Trim(Cstr(objSearchResultTree.Object.getItem(0).getItem(0).getData().toString()))	
		Call Fn_Setup_ReporterFilter("enableall")
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_Search_SearchResultsTreeOperations"
		If sNodePath="" Then
			DataTable.Value("ReusableActionWordReturnValue","Global")= ""
		Else
			aNodePath=Split(sNodePath,"-")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","SearchResultsTreeFirstNodeName","","")
			DataTable.Value("SearchResultsTreeFirstNodeName","Global")= aNodePath(1)
			aNodePath=Split(sNodePath,"/")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","SearchResultsTreeFirstNodeID","","")
			DataTable.Value("SearchResultsTreeFirstNodeID","Global")= aNodePath(0)
			DataTable.Value("ReusableActionWordReturnValue","Global")= sNodePath
		End If
		Datatable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select node by number from search results tree
	Case "SelectByNumber"
		sPopupMenu=sPopupMenu-1		
		sNodePath = Fn_RAC_GetJavaTreeNodePath(objSearchResultTree, sNodeName , "~", "@")		
		objSearchResultTree.Select sNodePath & "~#" & sPopupMenu
		sNodePath=sNodePath & "~" & sPopupMenu
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodePath) & " ] from search results due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Search Results Operations",sAction,"Node Path",sNodeName)
			If sObjectTypeName="" Then
				sObjectTypeName="Node"
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from search results tree","","","","DONOTSYNC","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	Case "VerifyResultItemIDSortOrderIsAscending"
		aNodePath=""
		If sPopupMenu="" Then
			sPopupMenu="/"
		End If

		For iCounter = 0 to Cint(objSearchResultTree.Object.getItem(0).getItemCount()) - 1
			sNodeName=objSearchResultTree.Object.getItem(0).getItem(iCounter).getData().toString() 
			aNodeName=SPlit(sNodeName,sPopupMenu)
			If aNodePath="" Then
				aNodePath=aNodeName(0)
			Else
				aNodePath=aNodePath & "^" & aNodeName(0)
			End If	
		Next
		
		If aNodePath="" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as search result does not contain any items to verify sorting order","","","","","") 
			Call Fn_ExitTest()
		Else
			aNodePath=Split(aNodePath,"^")
		End If
		
		If Fn_CommonUtil_StringArrayOperations("VerifyOrder",aNodePath,"Acending") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as search result items not sorted in [ Ascending ] order","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified order search result items sorted in [ Ascending ] order","","","","","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	Case "VerifySpecificContentAvailableInAllNodes"		
		iCount=objSearchResultTree.Object.getItem(0).getItemCount()-1
		If Cint(iCount)>21 Then
			iCount=20	
		End If
		
		For iCounter = 0 to Cint(iCount)
			sNodeName=objSearchResultTree.Object.getItem(0).getItem(iCounter).getData().toString() 
			If Instr(1,sNodeName,sPopupMenu) Then
				bFlag=True
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as search result node [ " & Cstr(sNodeName) & " ] does not contain value [ " & Cstr(sPopupMenu) & " ]","","","","","") 
				Call Fn_ExitTest()
				Exit For
			End If
		Next		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified order all search results contain value [ " & Cstr(sPopupMenu) & " ]","","","","DONOTSYNC","")
		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 'Case to verify existence on single search result
	 Case "VerifySingleResultDisplayedInSearchTree"
		
		If objSearchResultTree.GetROProperty("items count") = 2 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified single result displayed under search tree","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify single result displayed under search tree","","","","DONOTSYNC","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to multi select nodes from navigation tree
		Case "Multiselect"
			aNodeName=Split(sNodeName,"^")
			For iCounter=0 To UBound(aNodeName)
				aTempValue = Split(aNodeName(iCounter), "~")
				aTempValue(0) = objSearchResultTree.object.getItem(0).getData().tostring()
				aNodeName(iCounter) = Join(aTempValue, "~")
				iPath = Fn_RAC_GetJavaTreeNodePath(objSearchResultTree, aNodeName(iCounter) , "~", "@")
				If iPath=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Multi Select Node [" & Cstr(sNodeName) & "] of search result tree as node [" & Cstr(aNodeName(iCounter)) & "] does not exist","","","","","")
					Call Fn_ExitTest()
				Else
					'Multiselecting items
					If iCounter=0 Then
						objSearchResultTree.Select iPath				
					Else
						objSearchResultTree.ExtendSelect iPath				
					End If
					
					If Err.Number <> 0 Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multi select nodes [ " & CStr(sNodeName) & " ]","","","","","") 
						Call Fn_ExitTest()
					Else
						Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
						Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Myteamcenter search result tree Node Operations",sAction,"Node name",sNodeName)
						sObjectTypeName=Fn_RAC_GetTreeNodeType("SearchResultsTree","getobjecttypename")
						If sObjectTypeName="" Then
							sObjectTypeName="nodes"
						End If
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & CStr(sNodeName) & " ]","","","","DONOTSYNC","") 
					End If
				End If
			Next	
End Select

'Releasing all objects
Set objDefaultWindow=Nothing
Set objSearchResultTree=Nothing
Set objNoAccessibleObjects=Nothing
Set objContextMenu=Nothing

Function Fn_ExitTest()
	'Releasing all objects
	Set objDefaultWindow=Nothing
	Set objSearchResultTree=Nothing
	Set objNoAccessibleObjects=Nothing
	Set objContextMenu=Nothing
	ExitTest
End Function

