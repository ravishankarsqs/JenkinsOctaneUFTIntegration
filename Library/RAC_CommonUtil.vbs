Option Explicit
Err.Clear

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Function Name											|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -|- - - - - - - - - - - - - - - -| - - - - - - - |- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. 	Fn_RAC_ReadyStatusSync									|	vrushali.sahare@sqs.com		|	26-Feb-2016	|	Function used to waits till Application comes to Ready state
'002. 	Fn_RAC_GetSanitizedJavaTreeNodeName						|	sandeep.navghane@sqs.com	|	11-Mar-2016 |	Function used to remove unnecessary text from Java Tree Node name
'003. 	Fn_RAC_GetJavaTreeNodePath								|	sandeep.navghane@sqs.com	|	11-Mar-2016 |	Function used to retrive java tree node path by accessing java tree native methods
'004. 	Fn_RAC_GetJavaTreeNodeIndex								|	sandeep.navghane@sqs.com	|	11-Mar-2016 |	Function used to retrive java tree node index
'005. 	Fn_RAC_GetRealPropertyName								|	sandeep.navghane@sqs.com	|	11-Mar-2016 |	Function used to retrive real column / property name
'006. 	Fn_RAC_GetTreeNodeType									|	sandeep.navghane@sqs.com	|	20-Mar-2016 |	Function used to get object type of navigation tree node
'007. 	Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode	|	sandeep.navghane@sqs.com	|	20-Mar-2016 |	Function used to select the node in the assign all task tree in New Workflow Process Dialog
'008. 	Fn_RAC_GetMyWorklistNodePath							|	sandeep.navghane@sqs.com	|	29-Jun-2016 |	Function used to retrive my worklist node path by accessing java tree native methods
'009. 	Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName			|	sandeep.navghane@sqs.com	|	05-Jul-2016 |	Function used to retrive java tree node path by accessing java tree native [ getDisplayName ] methods
'010. 	Fn_RAC_PSEBOMTableRowOperations							|	sandeep.navghane@sqs.com	|	05-Jul-2016 |	Function used to perform operations on PSE BOM table rows
'011. 	Fn_RAC_PSEBOMTableColumnOperations						|	sandeep.navghane@sqs.com	|	08-Jul-2016 |	Function used to perform operations on PSE BOM table columns
'012. 	Fn_RAC_LOVTableOperations								|	sandeep.navghane@sqs.com	|	08-Jul-2016 |	Function used to perform operations on LOV tables
'013. 	Fn_RAC_GetActivePerspectiveName							|	sandeep.navghane@sqs.com	|	28-Jul-2016 |	Function used to get teamcenter active perspective name
'014. 	Fn_RAC_GetXMLNodeValue									|	sandeep.navghane@sqs.com	|	04-Aug-2016 |	Function used to get xml node value
'015. 	Fn_RAC_ProjectMemberSelectionTreeOperations				|	kundan.kudale@sqs.com		|	17-Nov-2016 |	Function used to perform operations on member selection tree
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_ReadyStatusSync
'
'Function Description	 :	Function used to waits till Application comes to Ready state
'
'Function Parameters	 :  1.iIterations: No. of times to be checked for Ready text						
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Teamcenter application should be displayed
'
'Function Usage		     :	Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  26-Feb-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Public Function Fn_RAC_ReadyStatusSync(iIterations)
'	'Declaring Variables
'	Dim iCounter, iCount, iProgressBarTimeout
'	Dim bFound
'	Dim objDefaultWindow,objReady
'	
'	bFound =  false
'	
'	'Creating object of Teamcenter Default Window
'	Set objDefaultWindow = JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0")
'	Set objReady=objDefaultWindow.JavaStaticText("path:=CLabel;StatusLine;Shell;","label:=Ready")
'	
'	If GBL_TCOBJECTS_SYNC_FLAG = False Then
'		If Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync","Exist", objDefaultWindow.JavaEdit("attached text:=Search"),GBL_ZERO_TIMEOUT,"","") Then
'			GBL_TCOBJECTS_SYNC_XAXIS = cInt(Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync", "getroproperty", objDefaultWindow.JavaEdit("attached text:=Search"), "","abs_x","")) + cInt(Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync", "getroproperty", objDefaultWindow.JavaEdit("attached text:=Search"), "","width","")) + 40  ' 251
'			GBL_TCOBJECTS_SYNC_YAXIS = cInt(Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync", "getroproperty", objDefaultWindow.JavaEdit("attached text:=Search"), "","abs_y","")) + cInt(Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync", "getroproperty", objDefaultWindow.JavaEdit("attached text:=Search"), "","height",""))/2     '150
'			GBL_TCOBJECTS_SYNC_FLAG = True
'		End If
'	End If
'	
'	For iCounter = 1 to iIterations
'		If Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync","Exist", objDefaultWindow ,GBL_ZERO_TIMEOUT,"","") Then
'			For iCount = 1 to 125
'				If Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync","Exist",objReady,GBL_MICRO_TIMEOUT,"","") Then
'					bFound=True
'					Exit for
'				End If
'			Next
'		Else
'			Fn_RAC_ReadyStatusSync = False
'			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_ReadyStatusSync ] : Teamcenter window does not exist.")	
'			'Release Object
'			Set objDefaultWindow = Nothing
'			Exit function
'		End If
'		If bFound Then Exit for
'	Next
'	
'	bFound =  false
'	If Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync","Exist",objReady,1,"","") Then
'		For iCounter = 1 to iIterations
'			iProgressBarTimeout = 0
'			For iCount = 1 to 240							
'				'Exit from inner loop if progressbar disappears
'				If Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync","Exist", objDefaultWindow.JavaObject("toolkit class:=org.eclipse.swt.widgets.ProgressBar","index:=0") ,iProgressBarTimeout,"","") = False Then
'					bFound = True
'					Exit for
'				Else
'					Wait 0,100
'				End If
'			Next
'			'Exit from main loop if progressbar disappears
'			If bFound Then Exit for
'		Next
'	End If
'	If Fn_UI_Object_Operations("Fn_RAC_ReadyStatusSync","Exist", objReady ,GBL_MIN_TIMEOUT,"","") = False OR bFound = False Then
'		Fn_RAC_ReadyStatusSync = False
'		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_ReadyStatusSync ] : Teamcenter Not Ready after [" + CStr(iIterations) + "] sync iterations")		
'	Else
'		Fn_RAC_ReadyStatusSync = True
'		'Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "PASS >> Teamcenter is Ready in [" + CStr(iIterations) + "] sync iterations")		
'	End If
'	
'	'Release Object
'	Set objReady=Nothing
'	Set objDefaultWindow = Nothing
'End Function

Public Function Fn_RAC_ReadyStatusSync(iIterations)
	'Declaring Variables
	Dim iCounter, iCount, iProgressBarTimeout
	Dim bFound
	Dim objDefaultWindow,objReady
	
	bFound =  false
	
	'Creating object of Teamcenter Default Window
	Set objDefaultWindow = JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0")
	Set objReady=objDefaultWindow.JavaStaticText("path:=CLabel;StatusLine;Shell;","label:=Ready")
	
'	If GBL_TCOBJECTS_SYNC_FLAG = False Then
'		If objDefaultWindow.JavaEdit("attached text:=Search").Exist(0) Then
'			GBL_TCOBJECTS_SYNC_XAXIS = cInt(objDefaultWindow.JavaEdit("attached text:=Search").GetROProperty("abs_x")) + cInt(objDefaultWindow.JavaEdit("attached text:=Search").GetROProperty("width")) + 40  '251
'			GBL_TCOBJECTS_SYNC_YAXIS = cInt(objDefaultWindow.JavaEdit("attached text:=Search").GetROProperty("abs_y")) + cInt(objDefaultWindow.JavaEdit("attached text:=Search").GetROProperty("height"))/2     '150
'			GBL_TCOBJECTS_SYNC_FLAG = True
'		End If
'	End If
	
	For iCounter = 1 to iIterations
		If objDefaultWindow.Exist(0) Then
			For iCount = 1 to 125
				If objReady.Exist(1) Then
					bFound=True
					Exit for
				End If
			Next
		Else
			Fn_RAC_ReadyStatusSync = False
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_ReadyStatusSync ] : Teamcenter window does not exist.")	
			'Release Object
			Set objDefaultWindow = Nothing
			Exit function
		End If
		If bFound Then Exit for
	Next
	
	bFound =  false
	If objReady.Exist(6) Then
		For iCounter = 1 to iIterations
			iProgressBarTimeout = 0
			For iCount = 1 to 240							
				'Exit from inner loop if progressbar disappears
				If objDefaultWindow.JavaObject("toolkit class:=org.eclipse.swt.widgets.ProgressBar","index:=0").Exist(iProgressBarTimeout) = False Then
					bFound = True
					Exit for
				Else
					Wait 0,100
				End If
			Next
			'Exit from main loop if progressbar disappears
			If bFound Then Exit for
		Next
	End If
	
	If objReady.Exist(1) = False OR bFound = False Then
		Fn_RAC_ReadyStatusSync = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_ReadyStatusSync ] : Teamcenter Not Ready after [" + CStr(iIterations) + "] sync iterations")
	Else
		Fn_RAC_ReadyStatusSync = True
	End If
	
	'Release Object
	Set objReady=Nothing
	Set objDefaultWindow = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_GetSanitizedJavaTreeNodeName
'
'Function Description	 :	Function used to remove unnecessary text from Java Tree Node name
'
'Function Parameters	 :  1.objJavaTreeNode : Tree Node Object					
'
'Function Return Value	 : 	Sanitized Node Text
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Java tree node should be available
'
'Function Usage		     :	bReturn=Fn_RAC_GetSanitizedJavaTreeNodeName(objTree.Object.getItem(0))
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  11-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Replacement for function : Fn_UI_JavaTree_GetSanitizedNodeName ' Delete this comment once implementation is completed
Public Function Fn_RAC_GetSanitizedJavaTreeNodeName(objJavaTreeNode)

	Fn_RAC_GetSanitizedJavaTreeNodeName = objJavaTreeNode.getData().toString()
	
	Select Case objJavaTreeNode.getData().getClass().toString()
		Case "class com.teamcenter.rac.cme.framework.application.model.impl.OccurrenceGroupNode"
			Fn_RAC_GetSanitizedJavaTreeNodeName = Replace(Fn_RAC_GetSanitizedJavaTreeNodeName," (OccurrenceGroupNode)","")
		Case "class com.teamcenter.rac.cme.framework.application.model.impl.StructureContextNode"
			Fn_RAC_GetSanitizedJavaTreeNodeName = Replace(Fn_RAC_GetSanitizedJavaTreeNodeName," (StructureContextNode)","")
		Case "class com.teamcenter.rac.cme.framework.application.model.impl.EndItemNode"
			Fn_RAC_GetSanitizedJavaTreeNodeName = Replace(Fn_RAC_GetSanitizedJavaTreeNodeName," (EndItemNode)","")
		Case "class com.teamcenter.rac.cme.framework.application.model.impl.ConfigurationContextNode"
			Fn_RAC_GetSanitizedJavaTreeNodeName = Replace(Fn_RAC_GetSanitizedJavaTreeNodeName," (ConfigurationContextNode)","")
		Case "class com.teamcenter.rac.cme.framework.application.model.impl.CCObjectRootNode"
			Fn_RAC_GetSanitizedJavaTreeNodeName = Replace(Fn_RAC_GetSanitizedJavaTreeNodeName," (CCObjectRootNode)","")
'		Case "class com.teamcenter.rac.cm.ui.changehome.ChangeHomePseudoFolder",_
'			 "class com.teamcenter.rac.cm.ui.changehome.ChangeHomeQueryPseudoFolder",_			 
'			Fn_RAC_GetSanitizedJavaTreeNodeName = objJavaTreeNode.getData().getDisplayName()
		Case Else
			'Do Nothing
	End Select
	
	Fn_RAC_GetSanitizedJavaTreeNodeName = Trim(Fn_RAC_GetSanitizedJavaTreeNodeName)
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_GetJavaTreeNodePath
'
'Function Description	 :	Function used to retrive java tree node path by accessing java tree native methods
'
'Function Parameters	 :  1.objJavaTree 		: Java Tree Object					
'							2.sTreeNode	  		: Tree node
'							3.sDelimiter 		: Tree node delimiter
'							4.sInstanceHandler 	: Tree node Instance Handler
'
'Function Return Value	 : 	False or Tree node path
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Java tree node should be available
'
'Function Usage		     :	bReturn=Fn_RAC_GetJavaTreeNodePath(JavaWindow("MyTeamcenter").JavaTree("NavTree"),"Home:000021-Test~000021/A;1-Test","~","@")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  11-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Replacement for function : Fn_UI_JavaTreeGetItemPathExt ' Delete this comment once implementation is completed
Public Function Fn_RAC_GetJavaTreeNodePath(objJavaTree,sTreeNode,sDelimiter,sInstanceHandler)
   'Variable Declaration
	Dim iCounter,iNodeItemsCount,iCount,iInstance,iOccurrence
	Dim sItemPath,sTempTreeNode,sSanitizedJavaTreeNode
	Dim aTreeNode,aTreeSubNode
	Dim objCurrentNode
	Dim bFlag

	If sDelimiter = "" Then sDelimiter = "~"
	If sInstanceHandler = "" Then sInstanceHandler = "@"
	
	Fn_RAC_GetJavaTreeNodePath = False
	
	Set GBL_JAVATREE_NODEBOUNDS_OBJECT = Nothing

	'Initial Item Path
	aTreeNode = Split (sTreeNode,sDelimiter)
	sItemPath=False
	bFlag=False
	
	'To handle the situation where operation needs to be performed on Root Node
	iOccurrence = 1
	'For iCount = 0 to Cint(objJavaTree.Object.getItemCount()) - 1
	 For iCount = 0 to Cint(objJavaTree.getroproperty("items count")) - 1
		If Instr(aTreeNode(0), sInstanceHandler) > 0 Then
			aTreeSubNode = split(aTreeNode(0),sInstanceHandler)
			sTempTreeNode = trim(aTreeSubNode(0))
			iInstance = Cint(aTreeSubNode(1))
		Else
			sTempTreeNode = trim(aTreeNode(0))
			iInstance = 1
		End If
		If objJavaTree.Object.getItem(iCount).getData().toString() = sTempTreeNode Then
			If  iOccurrence = iInstance Then
				Set objCurrentNode = objJavaTree.Object.getItem(iCount)
				sItemPath = "#" & iCount
				bFlag = True
				Exit For
			else
				iOccurrence = iOccurrence + 1
			End If
		Else
			sSanitizedJavaTreeNode = Fn_RAC_GetSanitizedJavaTreeNodeName(objJavaTree.Object.getItem(iCount))
			If sSanitizedJavaTreeNode = sTempTreeNode Then
				If  iOccurrence = iInstance Then
					Set objCurrentNode = objJavaTree.Object.getItem(iCount)
					sItemPath = "#" & iCount
					bFlag = True
					Exit For
				else
					iOccurrence = iOccurrence + 1
				End If
			End If
		End If
	Next
	If UBound(aTreeNode) = 0 Then
		Fn_RAC_GetJavaTreeNodePath = sItemPath
		If sItemPath=False Then
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_RAC_GetJavaTreeNodePath ] : Failed to find tree node [ " & Cstr(sTreeNode) & " ] under [ " & Cstr(objJavaTree.toString) & " ] tree")
		Else
			Set GBL_JAVATREE_NODEBOUNDS_OBJECT = objCurrentNode.getBounds()
			Set GBL_JAVATREE_CURRENTNODE_OBJECT=objCurrentNode
		End If
		Exit Function
	End If
	If bFlag Then
		bFlag = False
	Else
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_RAC_GetJavaTreeNodePath ] : Failed to find tree node [ " & Cstr(sTreeNode) & " ] under [ " & Cstr(objJavaTree.toString) & " ] tree")
		Exit function
	End If
	'To Select first Occurance of Node
	For iCount = 1 to UBound(aTreeNode)
		sTempTreeNode = aTreeNode(iCount)
		iNodeItemsCount = objCurrentNode.getItemCount()
		bFlag=False
		iOccurrence = 1
		If Instr(sTempTreeNode, sInstanceHandler) > 0 Then
			aTreeSubNode = split(sTempTreeNode,sInstanceHandler)
			sTempTreeNode = trim(aTreeSubNode(0))
			iInstance = Cint(aTreeSubNode(1))
		Else
			iInstance = 1
		End If
		For iCounter = 0 to iNodeItemsCount - 1
			If Trim(objCurrentNode.getItem(iCounter).getData().toString()) = Trim(sTempTreeNode) Then
				If  iOccurrence = iInstance Then
					Set objCurrentNode = objCurrentNode.getItem(iCounter)
					sItemPath = sItemPath & "~#" & iCounter
					bFlag=True
					Exit For
				else
					iOccurrence = iOccurrence + 1
				End If
			Else
				sSanitizedJavaTreeNode = Fn_RAC_GetSanitizedJavaTreeNodeName(objCurrentNode.getItem(iCounter))
				If sSanitizedJavaTreeNode = Trim(sTempTreeNode) Then
					If  iOccurrence = iInstance Then
						Set objCurrentNode = objCurrentNode.getItem(iCounter)
						sItemPath = sItemPath & "~#" & iCounter
						bFlag=True
						Exit For
					else
						iOccurrence = iOccurrence + 1
					End If
				End If
			End If
		Next
		If bFlag=False Then
			Exit For
		End If
	Next 
	If bFlag=True Then
		'Function Returns Item Path
		Fn_RAC_GetJavaTreeNodePath = sItemPath
		Set GBL_JAVATREE_NODEBOUNDS_OBJECT = objCurrentNode.getBounds()
		Set GBL_JAVATREE_CURRENTNODE_OBJECT=objCurrentNode
	Else
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_RAC_GetJavaTreeNodePath ] : Failed to find tree node [ " & Cstr(sTreeNode) & " ] under [ " & Cstr(objJavaTree.toString) & " ] tree")
		Fn_RAC_GetJavaTreeNodePath = False
	End If
	'releasing the objects
	Set objCurrentNode =Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_GetJavaTreeNodeIndex
'
'Function Description	 :	Function used to retrive java tree node index
'
'Function Parameters	 :  1.objJavaTree 		: Java Tree Object					
'							2.sTreeNode	  		: Tree node
'							3.sDelimiter 		: Tree node delimiter
'							3.sInstanceHandler 	: Tree node Instance Handler
'
'Function Return Value	 : 	False or Tree node index
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Java tree node should be available
'
'Function Usage		     :	bReturn=Fn_RAC_GetJavaTreeNodeIndex(JavaWindow("MyTeamcenter").JavaTree("NavTree"),"Home:000021-Test~000021/A;1-Test","~","@")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  11-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Replacement for function : Fn_UI_getJavaTreeIndexExt ' Delete this comment once implementation is completed
Public Function Fn_RAC_GetJavaTreeNodeIndex(objJavaTree,sTreeNode,sDelimiter,sInstanceHandler)
	'Variable Declaration
	Dim iOccurrence,iTempCounter,iNodeItemsCount,iCounter,iCount,iInstance	
	Dim aTreeNode,aTreeSubNode,aTreeSubNode1
	Dim sItemPath, sNode
	sNode=sTreeNode
	Fn_RAC_GetJavaTreeNodeIndex = -1
	
	If sDelimiter = ""  Then sDelimiter = "~"
	If sInstanceHandler = "" Then sInstanceHandler = "@"
	
	iInstance = 1

	iNodeItemsCount = cInt(objJavaTree.GetROProperty("items count"))
	iTempCounter = 0
	sNode = sTreeNode
	If InStr(sNode,sInstanceHandler) > 0 Then
		'Multiple instances
		aTreeNode = split(sNode, sDelimiter)
		If UBound(aTreeNode) <> 0 Then
			'Path with multiple instance handler
			For iCounter = 0 to ubound(aTreeNode)
				If iTempCounter = iNodeItemsCount Then
					'Node not found
					Exit FOR
				End If
				aTreeSubNode1 = split(aTreeNode(iCounter), sInstanceHandler)
				iInstance = 1
				If UBound(aTreeSubNode1) = 0 Then
					iOccurrence = 1
				Else
					iOccurrence = cInt(Trim(aTreeSubNode1(1)))
				End If
				sItemPath = ""
				'Generating node path
				For iCount = 0 to iCounter
					aTreeSubNode = split(aTreeNode(iCount), sInstanceHandler)
					If sItemPath = "" Then
						sItemPath = Trim(aTreeSubNode(0))
					Else
						sItemPath = sItemPath & "~" & Trim(aTreeSubNode(0))
					End If
				Next
				'verifiyig path
				Do While iTempCounter < iNodeItemsCount
					If objJavaTree.GetItem(iTempCounter) = sItemPath Then
						If iInstance = iOccurrence Then
							iTempCounter = iTempCounter + 1
							Exit do
						End If
						iInstance = iInstance + 1
					End If
					iTempCounter = iTempCounter + 1
				loop
			Next
			If iTempCounter <= iNodeItemsCount Then
				Fn_RAC_GetJavaTreeNodeIndex = iTempCounter - 1
			End If
		Else
			'With instance with no child items
			aTreeSubNode1 = split(aTreeNode(0), sInstanceHandler)
			sItemPath = Trim(aTreeSubNode1(0))
			iInstance = cInt(Trim(aTreeSubNode1(1)))
			For iTempCounter = 0 to iNodeItemsCount - 1
				If objJavaTree.GetItem(iTempCounter) = sItemPath Then
					If iInstance = iInstance Then
						Fn_RAC_GetJavaTreeNodeIndex = iTempCounter
						Exit for
					End If
					iInstance = iInstance + 1
				End If
			Next
		End If
	Else
		'Normal path without instance handler
		sItemPath = Replace(sNode,sDelimiter,"~")
		For iTempCounter = 0 to iNodeItemsCount - 1
			If Trim(Split(objJavaTree.GetItem(iTempCounter),"[")(0)) = sItemPath Then
				Fn_RAC_GetJavaTreeNodeIndex = iTempCounter
				Exit for
			End If
		Next
		If Fn_RAC_GetJavaTreeNodeIndex=-1 Then
			For iTempCounter = 0 to iNodeItemsCount - 1
				If InStr(1,objJavaTree.GetItem(iTempCounter), Split(sItemPath,"~")(UBound(Split(sItemPath,"~")))) Then
					Fn_RAC_GetJavaTreeNodeIndex = iTempCounter
					Exit for
				End If
			Next
		End If
	End If	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_GetRealPropertyName
'
'Function Description	 :	Function used to retrive real column / property name
'
'Function Parameters	 :  1.sProeprtyDisplayName	: column / property display name				
'
'Function Return Value	 : 	Real column / property name
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn=Fn_RAC_GetRealPropertyName("Type")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  11-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Replacement for function : Fn_UI_GetRealPropertyName ' Delete this comment once implementation is completed
Public Function Fn_RAC_GetRealPropertyName(sProeprtyDisplayName)

	Select Case trim(sProeprtyDisplayName)
		Case "Type","Object Type"
			Fn_RAC_GetRealPropertyName = "object_type"
		Case "Owner"
			Fn_RAC_GetRealPropertyName = "owning_user"
		Case "Object"
			Fn_RAC_GetRealPropertyName = "object_string"
		Case "Group ID"
			Fn_RAC_GetRealPropertyName = "owning_group"
		Case "Last Modified Date", "Date Modified"
			Fn_RAC_GetRealPropertyName = "last_mod_date"
		Case "Checked-Out"
			Fn_RAC_GetRealPropertyName = "checked_out"
		Case "Release Status"
			Fn_RAC_GetRealPropertyName = "release_status_list"
		Case "Checked-Out By"
			Fn_RAC_GetRealPropertyName = "checked_out_user"
		Case "Checked-Out Date"
			Fn_RAC_GetRealPropertyName = "checked_out_date"			
		Case "Description","Program Name"
			Fn_RAC_GetRealPropertyName = "object_desc"	
		Case "Name","Customer Part Name"
			Fn_RAC_GetRealPropertyName = "object_name"	
		Case "BOM Line"
			Fn_RAC_GetRealPropertyName = "bl_indented_title"
		Case "Item Type"
			Fn_RAC_GetRealPropertyName = "bl_item_object_type"
		Case "Item Description"
			Fn_RAC_GetRealPropertyName = "bl_item_object_desc"
		Case "Object Name"
			Fn_RAC_GetRealPropertyName="object_name"		
		Case "Customer Part Revision"
			Fn_RAC_GetRealPropertyName="ng5_customer_part_revision"
		Case "Customer"
			Fn_RAC_GetRealPropertyName="ng5_customer"
		Case "Comments"
			Fn_RAC_GetRealPropertyName="fnd0Comments"
		Case "User"
			Fn_RAC_GetRealPropertyName="fnd0AssigneeUser"
		Case "Performer"
			Fn_RAC_GetRealPropertyName="fnd0UserId"	
		Case "Event Type Name"
			Fn_RAC_GetRealPropertyName="fnd0EventTypeName"			
		Case Else
			Fn_RAC_GetRealPropertyName = sProeprtyDisplayName
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_GetTreeNodeType
'
'Function Description	 :	Function used to get object type of tree node
'
'Function Parameters	 :  1.sTreeName	: Tree name on user wants to perform operations
'							2.sAction	: Action to perform
'
'Function Return Value	 : 	Object type / empty
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn=Fn_RAC_GetTreeNodeType("NavigationTree","getobjecttypename")
'Function Usage		     :	bReturn=Fn_RAC_GetTreeNodeType("SearchResultsTree","getobjecttypename")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  20-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_RAC_GetTreeNodeType(sTreeName,sAction)	
	'Declaring Variables
	Dim sNodeType
	'Initially function return value assign to empty
	Fn_RAC_GetTreeNodeType=""
	sNodeType=""
	
	Select Case Lcase(sTreeName)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "navigationtree","searchresultstree","componenttabtree"
			Select Case Lcase(sAction)
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				'Case to get dataset instance number under specific node
				Case "getobjecttypename"
					sNodeType=GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getProperty("object_type")
					Fn_RAC_GetTreeNodeType=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NavigationTreeNodeTypeInformation_APL",sNodeType,"")
					If Fn_RAC_GetTreeNodeType=False Then
						Fn_RAC_GetTreeNodeType=sNodeType
					End If			
		   End Select
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "psebomtable"
			Select Case Lcase(sAction)
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "getitemtypename"				
					sNodeType=GBL_JAVATREE_CURRENTNODE_OBJECT.getProperty("bl_item_object_type")
					If sNodeType="" Then
						Fn_RAC_GetTreeNodeType="node"
					Else
						Fn_RAC_GetTreeNodeType=sNodeType
					End If
		   End Select
   End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode
'
'Function Description	 :	Function used to select the node in the assign all task tree in New Workflow Process Dialog
'
'Function Parameters	 :  1.objTree  	: Tree object
'							2.sNode 	: Node path to select
'							2.sReserve 	: Reserve parameter in case required in future
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	New Process Dialog assign all task tab should be selected
'
'Function Usage		     :	bReturn=Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode(objTree,"Process~Review Task~*/Designer/1", "")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  25-May-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode(objTree,sNode,sReserve)
	'Declaring variables
	Dim iItemsCount, iCounter
	Dim sPath,aPath
	'Initially function return value assign to false
	Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode = False
	'Expand parent path
	aPath = Split(sNode,"~")
	objTree.Object.expandToLevel(ubound(aPath))
	Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	
	sPath = Replace(sNode,"~",", ")
	sPath = "[" & sPath & "]"
	iItemsCount = objTree.GetROProperty("items count")
	For iCounter = 0 To iItemsCount-1 Step 1
		If objTree.Object.getPathForRow(iCounter).toString() = sPath Then
			objTree.Object.setSelectionRow(iCounter)
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Exit For
		End If
	Next
	If Cint(iCounter) = Cint(iItemsCount) Then
		Exit Function
	End If
	
	If Err.Number < 0 then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode ] Fail to select node [ " & Cstr(sNode) & " ] of new process assign all task tree due to error number [ " & Cstr(Err.Number) & " ] and error description [ " & Cstr(Err.Description) & " ]")
		Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode=False
	Else
		Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode = True
	End If	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_GetMyWorklistNodePath
'
'Function Description	 :	Function used to retrive my worklist node path by accessing java tree native methods
'
'Function Parameters	 :  1.objTree	 		: My Worklist Tree Object					
'							2.sTreeNode	  		: Tree node
'							3.sDelimiter 		: Tree node delimiter
'
'Function Return Value	 : 	False or Tree node path
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	My worklist tree node should be available
'
'Function Usage		     :	bReturn=Fn_RAC_GetMyWorklistNodePath(JavaWindow("jwnd_MyWorkListWindow").JavaTree("jtree_MyWorkListTree"),"My Worklist~TestUser1 (TestUser1) Inbox~Tasks To Perform~000021/A;1-Test","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  29-Jun-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_RAC_GetMyWorklistNodePath(objTree,ByVal sTreeNode,sDelimiter)

	'Variable Declaration
	Dim iCount,iCounter,iItemCount,iInstance
	Dim sNodePath,sTreeNodeToStr2
	Dim aTreeNode,aNodePath	
	Dim objItem

	'Set default delimiter
	If sDelimiter="" Then
		sDelimiter="~"
	End If
	
	'Initially set function return value as False
	Fn_RAC_GetMyWorklistNodePath = False
	
	'Spliting Node path
	aTreeNode = Split(sTreeNode, sDelimiter, -1, 1)
	Set objItem =objTree.Object.getItem(0)
	sNodePath = "#0"

	If UBound(aTreeNode)=0 Then
		Fn_RAC_GetMyWorklistNodePath = sNodePath
		Exit function
	End If
	For iCounter = 1 to UBound(aTreeNode)
		iItemCount = cInt(objItem.getItemCount())
		'Getting instance
		aNodePath = split(trim(aTreeNode(iCounter)), "@")
		If UBound(aNodePath) > 0 Then
			aNodePath(0) = trim(aNodePath(0))
			iInstance = cint(aNodePath(1))
		Else
			iInstance = 1
		End If
		For iCount = 0 to iItemCount -1
			'verify node
			If Trim(objItem.getItem(iCount).getData().toString()) = Trim(aNodePath(0)) Then
				If  iInstance = 1 Then
					If instr(aNodePath(0), "#") > 0 Then
						sNodePath = sNodePath & "~" &  aNodePath(0)
					Else
						sNodePath = sNodePath & "~#" &  cstr(iCount)
					End If
					Set objItem = objItem.getItem(iCount)
					Exit For
				Else
					iInstance = iInstance - 1
				End If
			ElseIf instr(aTreeNode(iCounter),")") > 0  Then
				sTreeNodeToStr2 = ""
				sTreeNodeToStr2 = Trim(objItem.getItem(iCount).getData().getComponent().toString())

				If sTreeNodeToStr2 = aNodePath(0) Then
					If  iInstance = 1 Then
						sNodePath = sNodePath & "~#" &  cstr(iCount)
						Set objItem = objItem.getItem(iCount)
						Exit For
					Else
						iInstance = iInstance - 1
					End If
				Else
					On Error Resume Next
					sTreeNodeToStr2 = ""
					Call Fn_Setup_ReporterFilter("DisableAll")
					sTreeNodeToStr2 = Trim(objItem.getItem(iCount).getData().getComponent().toString2()) 
					Call Fn_Setup_ReporterFilter("EnableAll")
					If sTreeNodeToStr2 = "" Then
						sTreeNodeToStr2 = Trim(objItem.getItem(iCount).getData().getComponent().toString())
						If Instr(1,Cstr(aNodePath(0)),Cstr(sTreeNodeToStr2)) Then
							If  iInstance = 1 Then
								sNodePath = sNodePath & "~#" &  cstr(iCount)
								Set objItem = objItem.getItem(iCount)
								Exit For
							Else
								iInstance = iInstance - 1
							End If
						End If
					Else
						If sTreeNodeToStr2 = aNodePath(0) Then
							If  iInstance = 1 Then
								sNodePath = sNodePath & "~#" & cstr(iCount)
								Set objItem = objItem.getItem(iCount)
								Exit For
							Else
								iInstance = iInstance - 1
							End If
						End If
					End If
				End If

			End If
		Next
		If iCount = iItemCount Then
			Set objItem = Nothing
			Fn_RAC_GetMyWorklistNodePath = False
'			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_RAC_GetMyWorklistNodePath ] : Fail to perform operation due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
			Exit Function
		End If
	Next
	Set objItem = Nothing
	
	If Err.Number < 0 Then
		Fn_RAC_GetMyWorklistNodePath=False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_RAC_GetMyWorklistNodePath ] : Fail to perform operation due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	Else
		Fn_RAC_GetMyWorklistNodePath = sNodePath
	End If	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName
'
'Function Description	 :	Function used to retrive java tree node path by accessing java tree native [ getDisplayName ] methods
'
'Function Parameters	 :  1.objJavaTree 		: Java Tree Object					
'							2.sTreeNode	  		: Tree node
'							3.sDelimiter 		: Tree node delimiter
'							4.sInstanceHandler 	: Tree node Instance Handler
'
'Function Return Value	 : 	False or Tree node path
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Java tree node should be available
'
'Function Usage		     :	bReturn=Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName(JavaWindow("ChangeManager").JavaTree("NavigationTree"),"Home~000021-Test~000021/A;1-Test","~","@")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  05-Jul-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName(objJavaTree,sTreeNode,sDelimiter,sInstanceHandler)
	'Variable Declaration
	Dim iCounter,iCount,iItemCount,iInstance
	Dim sItemPath,sTempTreeNode
	Dim objCurrentNode
	Dim aTreeNode
	
	If sDelimiter = "" Then sDelimiter = "~"
	If sInstanceHandler = "" Then sInstanceHandler = "@"

	Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName = False
	
	'Initial Item Path
	aTreeNode = split(sTreeNode, sDelimiter, -1, 1)
	Set objCurrentNode = objJavaTree.Object.getItem(0)

	If sTreeNode <> "" Then
		sTempTreeNode = ""
		sTempTreeNode = Trim(objJavaTree.Object.getItem(0).getData().toString())
		If sTempTreeNode = "" Then
			sTempTreeNode = Trim(objJavaTree.Object.getItem(0).getData().getDisplayName())
		End If
		
		If  sTempTreeNode = Trim(aTreeNode(0)) Then
			sItemPath = "#0"
		 Else
			Exit Function			
		End If
	Else
		Exit Function 
	End If	

	For iCount = 1 to UBOund(aTreeNode)
		iItemCount = cInt(objCurrentNode.getItemCount())
		bFlag=False
		aTreeNodePath = Split(Trim(aTreeNode(iCount)),"@")
		If uBound(aTreeNodePath) > 0 Then
			aTreeNodePath(0) = Trim(aTreeNodePath(0))
			iInstance = cint(aTreeNodePath(1))
		Else
			iInstance = 1
		End If
		For iCounter = 0 to iItemCount -1
			sTempTreeNode = ""
			sTempTreeNode = Trim(objCurrentNode.getItem(iCounter).getData().toString()) 
			If sTempTreeNode = "" Then
				sTempTreeNode = Trim(objCurrentNode.getItem(iCounter).getData().getDisplayName()) 
			End If
			If instr(sTempTreeNode,"(") > 0 Then
				If instr(sTempTreeNode, aTreeNodePath(0)) > 0 Then
					If  iInstance = 1 Then
						sItemPath = sItemPath & "~#" &  cstr(iCounter)
						Set objCurrentNode = objCurrentNode.getItem(iCounter)
						bFlag=True
						Exit For
					Else
						iInstance = iInstance - 1
					End If
				End If
			Else
				If sTempTreeNode = aTreeNodePath(0) Then
					If  iInstance = 1 Then
						sItemPath = sItemPath & "~#" &  cstr(iCounter)
						Set objCurrentNode = objCurrentNode.getItem(iCounter)
						bFlag=True
						Exit For
					Else
						iInstance = iInstance - 1
					End If
				End If
			End If
		Next
        If bFlag=False Then
			Exit For
		End If
	Next
		
	If bFlag=True Then
		'Function Returns Item Path
		Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName = sItemPath
		Set GBL_JAVATREE_CURRENTNODE_OBJECT=objCurrentNode
	Else
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName ] : Failed to find tree node [ " & Cstr(sTreeNode) & " ] under [ " & Cstr(objJavaTree.toString) & " ] tree")
		Fn_RAC_GetJavaTreeNodePathUsingGetDisplayName = False
	End If	
	'releasing the objects
	Set objCurrentNode = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_PSEBOMTableRowOperations
'
'Function Description	 :	Function used to perform operations on PSE BOM table rows
'
'Function Parameters	 :  1.sAction	 		: Action to perform					
'							2.objJavaTabl  		: PSE BOM table object
'							3.sNodeName 		: Tree\Table node path
'
'Function Return Value	 : 	-1 or Row number of node
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	PSE BOM table should be available
'
'Function Usage		     :	bReturn=Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,"518611/A;1-Item_518611 (view)~001270/A;1-ffff")
'Function Usage		     :	bReturn=Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,"518611/A;1-Item_518611 (view)~001270/A;1-ffff")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  29-Jun-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_RAC_PSEBOMTableRowOperations(sAction,ByVal objJavaTable,sNodeName)
	'Variable Declaration
	Dim iInstance,iOccurance,iRows,iCounter,iCount,iColumnIndex,iRowIndex,iCols
	Dim sNodePath,sNodeName1,sNodePath1,sPath,sPath1,sNodePath2,sTopNode
	Dim aNodePath,aRowNode
	Dim bFlag
	Dim objComponent,aNodeName
	
	'Initially set function return value as -1
	Fn_RAC_PSEBOMTableRowOperations = -1
	
	'Checking existance of structure manager BOM table
	If Fn_UI_Object_Operations("Fn_RAC_PSEBOMTableColumnOperations","Exist",objJavaTable,"","","") = False Then
		Exit function
	End If
	
	Select Case LCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getnodeindex"
			If InStr(sNodeName,"@") > 0 Then
				aNodePath = Split(sNodeName,"@",-1, 1)
				sNodeName = aNodePath(0)
				if isNumeric(Trim(aNodePath(1))) = False Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_PSEBOMTableRowOperations ] : Fail to get row index of node [ " & Cstr(sNodeName) & " ]")
					Exit Function
				end If
				iInstance = cInt(aNodePath(1))
			Else
				iInstance = 1
			End If
			
			'Format the Inout as per Table Default Nodes
			sNodeName = Replace(sNodeName,"~",", ",1,-1,1)
			iOccurance = 0
			iRows=objJavaTable.GetROProperty("rows")
			'Get the Row No. of required Node
			For iCounter = 1 to iInstance
				For iCount = iOccurance to iRows-1
					sNodePath = objJavaTable.Object.getPathForRow(iCount).toString
					sNodePath = Right(sNodePath,(Len(sNodePath)-InStr(1, sNodePath, ",", 1)))					
					sNodePath = Left(sNodePath, Len(sNodePath)-1)					
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					'Added Code to match node removing (view) word
                    sNodePath1=replace(LCase(sNodePath)," (view)","")
					sNodeName1=replace(LCase(sNodeName)," (view)","")
					sNodePath1=replace(LCase(sNodePath1)," ,",",")
					sNodeName1=replace(LCase(sNodeName1)," ,",",")
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					If Trim(sNodePath) = Trim(sNodeName) Then
						iRowIndex = iCount
						iOccurance = iRowIndex + 1 
						Exit For
                    ElseIf Trim(sNodePath1) = Trim(sNodeName1) Then
						iRowIndex = iCount
						iOccurance = iRowIndex + 1 
						Exit For
					End If
				Next
			Next
			If  cStr(iCount) = cStr(iRows) Then
				Fn_RAC_PSEBOMTableRowOperations = -1
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_PSEBOMTableRowOperations ] : Fail to get row index of node [ " & Cstr(sNodeName) & " ] as node is not exist in table")
			Else
				Fn_RAC_PSEBOMTableRowOperations = iRowIndex
				Set GBL_JAVATREE_CURRENTNODE_OBJECT=objJavaTable.Object.getComponentForRow(iRowIndex)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<PASS>:  [ Fn_RAC_PSEBOMTableRowOperations ] : Row index of node [ " & Cstr(sNodeName) & " ] is [ " & CStr(iRowIndex) & " ]")
			End If
			'Releasing table object	
			Set objJavaTable = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getnodeindexext"
			iColumnIndex = 0
			bFlag = False
			iRows = cInt(objJavaTable.GetROProperty ("rows"))
			aNodeName = split(sNodeName,"~")
			iRowIndex = 0
			sPath = ""
			For iCounter=0 to UBound(aNodeName)
				aRowNode = Split(trim((aNodeName(iCounter))),"@")
				If sPath = "" Then
					sPath =  trim(aRowNode(0))
				Else
					sPath = sPath & "~" & trim(aRowNode(0))
				End If
			Next
			For iCounter=0 to UBound(aNodeName)
				If iRowIndex = iRows  Then
					Exit for
				End If
				aRowNode = split(trim((aNodeName(iCounter))),"@")
				iInstance = 0
				bFlag = False
				Do While iRowIndex < iRows
					If uBound(aRowNode) > 0 Then
						'instance number exist in name
						'initialize instance number
						'ith row matches with aRowNode(0) then
						sNodePath = objJavaTable.object.getValueAt(iRowIndex,iColumnIndex).toString()
						sNodePath1 = Replace(LCase(sNodePath)," (view)","") 
						sTopNode=Replace(LCase(aRowNode(0))," (view)","")
							
						If trim(sNodePath) = trim(aRowNode(0)) or trim(sNodePath1) = trim(sTopNode) then
							Set objComponent = objJavaTable.object.getComponentForRow(iRowIndex)
							sNodePath2 = ""
							Do while NOT (objComponent is Nothing)
								If sNodePath2 = "" Then
									sNodePath2 = objComponent.getProperty("bl_indented_title")
								Else
									sNodePath2 = objComponent.getProperty("bl_indented_title") & ", " & sNodePath2
								End If
								If Environment.Value("ProductName")=GBL_HP_QTP_PRODUCTNAME Then
									If IsObject(objComponent.parent())=True Then
										Set objComponent = objComponent.parent()
									Else
										Exit do
									End If
								Else
									Set objComponent = objComponent.parent()
									If  objComponent is Nothing Then
										Exit do
									End If	
								End If
							Loop

							If instr(sNodePath2, "@BOM::") > 0 Then
								sNodePath2 = trim(replace(sNodePath2,"""",""))
								aNodePath = split(sNodePath2,",")
								sNodePath2 = ""
								For iCount = 0 to uBound(aNodePath)
									aNodePath(iCount) = Left(aNodePath(iCount), instr(aNodePath(iCount),"@")-1)
									If sNodePath2 = "" Then
										sNodePath2 = trim(aNodePath(iCount))
									else
										sNodePath2 = sNodePath2 & ", " & trim(aNodePath(iCount))
									End If
								Next
							End If
								
							sNodePath2 = trim(replace(sNodePath2,", ","~"))
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
							'Added Code to match node removing (view) word
							sNodePath1=replace(LCase(sNodePath2)," (view)","")
							sPath1=replace(LCase(sPath)," (view)","")
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -																									
							If instr(sPath, sNodePath2 ) > 0 Then
								iInstance = iInstance +1
								If iInstance = cInt(aRowNode(1)) Then 
									If UBound(aNodeName) = iCounter Then
										bFlag = True
									End If
									Exit do
								End If												
							ElseIf instr(sPath1, sNodePath1 ) > 0 Then
								iInstance = iInstance +1
								If iInstance = cInt(aRowNode(1)) Then 
									If UBound(aNodeName) = iCounter Then
										bFlag = True
									End If
									Exit do
								End If
							End If
						End if
					Else
						'if row matches with aRowNode(0) then
						If objJavaTable.object.getPathForRow(iRowIndex).getLastPathComponent().getClass().toString() <> "class com.teamcenter.rac.treetable.HiddenSiblingNode" Then
							sNodePath = objJavaTable.object.getValueAt(iRowIndex, iColumnIndex).toString()
						Else
							sNodePath = ""
						End If				
						sNodePath1 = Replace(LCase(sNodePath)," (view)","") 
						sTopNode=Replace(LCase(aRowNode(0))," (view)","") 

						If trim(sNodePath) = trim(aRowNode(0)) or trim(sNodePath1) =sTopNode  then
							set objComponent = objJavaTable.object.getComponentForRow(iRowIndex)
							sNodePath2 = ""
							Do while NOT (objComponent is Nothing)
								If sNodePath2 = "" Then
									sNodePath2 = objComponent.getProperty("bl_indented_title")
								Else
									sNodePath2 =objComponent.getProperty("bl_indented_title") & ", " & sNodePath2
								End If
								If Environment.Value("ProductName")=GBL_HP_QTP_PRODUCTNAME Then
									If IsObject(objComponent.parent())=True Then
										Set objComponent = objComponent.parent()
									Else
										Exit do
									End If
								Else
									Set objComponent = objComponent.parent()
									If  objComponent is Nothing Then
										Exit do
									End If	
								End If
								
							Loop
							If instr(sNodePath2, "@BOM::") > 0 Then
								sNodePath2 = trim(replace(sNodePath2,"""",""))
								aNodePath = split(sNodePath2,",")
								sNodePath2 = ""
								For iCount = 0 to uBound(aNodePath)
									aNodePath(iCount) = Left(aNodePath(iCount), instr(aNodePath(iCount),"@")-1)
									If sNodePath2 = "" Then
										sNodePath2 = trim(aNodePath(iCount))
									else
										sNodePath2 = sNodePath2 & ", " & trim(aNodePath(iCount))
									End If
								Next
							End If
						
							sNodePath2 = trim(replace(sNodePath2,", ","~"))
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
							'Added Code to match node removing (view) word
							sNodePath1=replace(LCase(sNodePath2)," (view)","")
							sPath1=replace(LCase(sPath)," (view)","")
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
							If instr(sPath, sNodePath2 ) > 0 Then
								If UBound(aNodeName) = iCounter Then
									bFlag = True
								End If
								Exit do
								'exit loop
							ElseIf instr(sPath1, sNodePath1 ) > 0 Then
								If UBound(aNodeName) = iCounter Then
									bFlag = True
								End If
								Exit do
							End if
						End if
					End If
					iRowIndex = iRowIndex + 1
				Loop
			Next
			If bFlag=False Then
				Fn_RAC_PSEBOMTableRowOperations = -1
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_PSEBOMTableRowOperations ] : Fail to get row index of node [ " & Cstr(sNodeName) & " ] as node is not exist in table")
			Else
				Fn_RAC_PSEBOMTableRowOperations = iRowIndex
				Set GBL_JAVATREE_CURRENTNODE_OBJECT=objJavaTable.Object.getComponentForRow(iRowIndex)
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<PASS>:  [ Fn_RAC_PSEBOMTableRowOperations ] : Row index of node [ " & Cstr(sNodeName) & " ] is [ " & CStr(iRowIndex) & " ]")
			End If
			'Releasing table object	
			Set objJavaTable = Nothing	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getattachmentstablenodeindex"
			iRows = objJavaTable.GetROProperty("rows")
			iCols =  objJavaTable.GetROProperty("cols")
			iColumnIndex = Fn_RAC_PSEBOMTableColumnOperations("getcolumnindex",objJavaTable,"Line")
			
			If  instr(sNodeName, "@") > 0 Then
				aNodePath = Split(sNodeName, "@",-1, 1)
				sNodeName = trim(aNodePath(0))
				iInstance = cint(aNodePath(1))
			Else
				iInstance = 1
			End If
			
			aNodePath = Split(sNodeName, "~")
			sNodeName=aNodePath(Ubound(aNodePath))
			iOccurance = 0
			For iCounter = 0 to iRows -1
				If trim(Trim(objJavaTable.Object.getValueAt(iCounter,iColumnIndex).toString)) = trim(sNodeName) Then
					iOccurance = iOccurance + 1
					If iOccurance = iInstance Then
						Fn_RAC_PSEBOMTableRowOperations  = iCounter
						Exit for
					End If
				End If
			Next			
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_PSEBOMTableColumnOperations
'
'Function Description	 :	Function used to perform operations on PSE BOM table columns
'
'Function Parameters	 :  1.sAction	 		: Action to perform					
'							2.objJavaTabl  		: PSE BOM table object
'							3.sColumnName 		: Table column name
'
'Function Return Value	 : 	-1 or Column number
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	PSE BOM table should be available
'
'Function Usage		     :	bReturn=Fn_RAC_PSEBOMTableColumnOperations("getcolumnindex",objBOMTable,"BOM Line")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  08-Jul-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_RAC_PSEBOMTableColumnOperations(sAction,ByVal objJavaTable,sColumnName)
	'Variable Declaration
	Dim iCols,iCounter
	Dim objTable
	
	'Initially set function return value as -1
	Fn_RAC_PSEBOMTableColumnOperations = -1
	
	'Checking existance of structure manager BOM table
	If Fn_UI_Object_Operations("Fn_RAC_PSEBOMTableColumnOperations","Exist",objJavaTable,"","","") = False Then
		Exit function
	End If
	
	Select case LCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getcolumnindex"
			'Get the No. of columns present in the BOM Table
			iCols = objJavaTable.GetROProperty("cols")
			Set objTable = objJavaTable.Object
			'Get the Col No. of required Column
			For iCounter = 0 to iCols -1
				'Comparing column names
				If Trim(objTable.getColumnName(iCounter)) = Trim(sColumnName) Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<PASS>:  [ Fn_RAC_PSEBOMTableColumnOperations ] : Column index for column [ " & Cstr(sColumnName) & " ] is [ " & CStr(iCounter) & " ]")
					Fn_RAC_PSEBOMTableColumnOperations = iCounter
					Exit For
				End If
			Next
			If Fn_RAC_PSEBOMTableColumnOperations=-1 Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_PSEBOMTableColumnOperations ] : Fail to get column index of column [ " & Cstr(sColumnName) & " ] as specified column does not exist in table")
			End If
			'Release the Table object
			Set objTable = Nothing
	End Select
		
	If  Err.number <> 0 Then
		Fn_RAC_PSEBOMTableColumnOperations = -1
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  [ Fn_RAC_PSEBOMTableColumnOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] on BOM table due to error [ " & Cstr(Err.Description) & " ]")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_LOVTableOperations
'
'Function Description	 :	Function used to perform operations on LOV tables
'
'Function Parameters	 :  1.sAction	 				: Action to perform					
'							2.objTableContainer			: Parent object of LOV table
'							3.sValue	 				: Table content value
'							4.sLOVDropDownButtonName	: LOV Table drop down button name
'							5.sColumnName				: LOV Table column name
'
'Function Return Value	 : 	True Or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	LOV table should be available
'
'Function Usage		     :	bReturn=Fn_RAC_LOVTableOperations("setvalue",objNewPartDialog,"each","jbtn_LOVDropDown","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  08-Jul-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_RAC_LOVTableOperations(sAction,objTableContainer,sValue,sLOVDropDownButtonName,sColumnName)
	'Declaring variables
	Dim objDescription,objChildObjects
	
	'Initially function returns false value
	Fn_RAC_LOVTableOperations=False
	
	Select Case Lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select and set value from LOV tree table
		Case "setvalue"
			IF sLOVDropDownButtonName<>"" Then
				'Click on LOV drop down button		
				If Fn_UI_JavaButton_Operations("Fn_RAC_LOVTableOperations", "Click", objTableContainer,sLOVDropDownButtonName)=False Then
					Exit Function
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End If
			'Creating Description object for LOV Table
			Set objDescription = Description.Create
			objDescription("Class Name").value="JavaTable"
			objDescription("tagname").Value = "LOVTreeTable"
			Set objChildObjects = objNewPartDialog.ChildObjects(objDescription)
			'Check if value is present in LOV table. If found return true, else return false
			If objChildObjects.Count > 0 Then
				bFlag=False
				For iCounter=0 to objChildObjects(0).GetROProperty("rows")
					If trim(sValue)=trim(objChildObjects(0).Object.getValueAt(iCounter,0).getDisplayableValue()) Then
						objChildObjects(0).DoubleClickCell iCounter,0
						Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
						bFlag=True
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_RAC_LOVTableOperations=True
				End If
			Else
				Exit Function
			End If
			'Releasing objects
			Set objChildObjects =Nothing
			Set objDescription = Nothing
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_GetActivePerspectiveName
'
'Function Description	 :	Function used to get teamcenter active perspective name
'
'Function Parameters	 :  1.sAction : Action to perform						
'
'Function Return Value	 : 	Perspective name
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Teamcenter application should be displayed
'
'Function Usage		     :	Call Fn_RAC_GetActivePerspectiveName("")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  26-Feb-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_RAC_GetActivePerspectiveName(sAction)
	'Declaring Variables
	Dim objDefaultWindow
	Dim sPerspectiveName
	
	Fn_RAC_GetActivePerspectiveName=""
	
	'Creating object of Teamcenter Default Window
	Set objDefaultWindow = JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0")
	
	Select Case Lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "","getnameext"
			sPerspectiveName=Split(objDefaultWindow.GetROProperty("title"),"-")
			sPerspectiveName(0)=Trim(sPerspectiveName(0))
			Fn_RAC_GetActivePerspectiveName=Replace(sPerspectiveName(0)," ","")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		Case "getname"
			sPerspectiveName=Split(objDefaultWindow.GetROProperty("title"),"-")
			Fn_RAC_GetActivePerspectiveName=Trim(sPerspectiveName(0))			
	End Select	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_GetXMLNodeValue
'
'Function Description	 :	Function used to get xml node value
'
'Function Parameters	 :  1.sActionWordName : Action word name
'							2.sAction		  : sub action name
'							3.sXMLTag		  : XML tag name
'
'Function Return Value	 : 	False \ XML node value
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	XML node should be available
'
'Function Usage		     :	Call Fn_RAC_GetXMLNodeValue("RAC_MyWorklist_TreeNodeOperations","","ViewAuditLogs")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  04-Aug-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_RAC_GetXMLNodeValue(sActionWordName,sAction,sXMLTag)

	Fn_RAC_GetXMLNodeValue=False
	
	Select Case sActionWordName
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "RAC_MyWorklist_TreeNodeOperations","RAC_Search_SearchResultsTreeOperations","RAC_MyTc_ComponentTabTreeOperations","RAC_ChangeManager_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations","RAC_PSE_BOMTableOperations","RAC_Common_SummaryTabTableOperations"
			If Fn_FSOUtil_XMLFileOperations("getvalue","RAC_Common_PPM",sXMLTag,"")<>False Then
				Fn_RAC_GetXMLNodeValue=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_Common_PPM",sXMLTag,"")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "RAC_Common_MenuOperations","Fn_SMV_MenuOperations"
			If Fn_FSOUtil_XMLFileOperations("getvalue","RAC_Common_Menu",sXMLTag,"")<>False Then
				Fn_RAC_GetXMLNodeValue=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_Common_Menu",sXMLTag,"")
			ElseIf Fn_FSOUtil_XMLFileOperations("getvalue","RAC_StructureManager_Menu",sXMLTag,"")<>False Then
				Fn_RAC_GetXMLNodeValue=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_StructureManager_Menu",sXMLTag,"")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "RAC_Common_ToolbarOperations"
			If Fn_FSOUtil_XMLFileOperations("getvalue","RAC_Common_TLB",sXMLTag,"")<>False Then
				Fn_RAC_GetXMLNodeValue=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_Common_TLB",sXMLTag,"")
			ElseIf Fn_FSOUtil_XMLFileOperations("getvalue","RAC_StructureManager_TLB",sXMLTag,"")<>False Then
				Fn_RAC_GetXMLNodeValue=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_StructureManager_TLB",sXMLTag,"")
			End If
	End Select	
End Function


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_ProjectMemberSelectionTreeOperations
'
'Function Description	 :	Function used to perform operations on member selection tree
'
'Function Parameters	 :  1.sAction		: 	Acion to be performed
'							2.sNode			:	Node name on which operation is to be performed 
'							3.sPopupMenu	: 	Popup menu option to be selected
'							4.objMemberSelectionTree	: 	Tree object on which operation is to be performed. It can be either member selection tree of selected memeber tree
'
'Function Return Value	 : 	False / True
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Project perspective should be open
'
'Function Usage		     :	Call Fn_RAC_ProjectMemberSelectionTreeOperations("Select","Engineering~Designer~UserName (UserID)","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Kundan Kudale			    |  17-Nov-2016	    |	 1.0		|	  Sandeep Navghane	| 	  Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_RAC_ProjectMemberSelectionTreeOperations(sAction, sNode, sPopupMenu, objMemberSelectionTree)

	'Variables declaration
	Dim iCounter,aNode,sNodePath
	Dim iCount
	Dim bFlag
	
	'Set initial return value as False
	Fn_RAC_ProjectMemberSelectionTreeOperations=False
	
	'Get object of member selection tree from XML file
	If objMemberSelectionTree.GetItem(0) <> "Team" Then
		sNode = Replace(sNode, "Team", Trim(objMemberSelectionTree.GetItem(0)))
	End If
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select node
		Case "Select"
			aNode=Split(sNode,"~")
			
			'Expand the tree till parent node
			For iCounter = 0 To ubound(aNode)-1
				If iCounter=0 Then
					sNodePath=aNode(0)
				Else
					sNodePath=sNodePath+"~"+aNode(iCounter)
				End If
				
				objMemberSelectionTree.Expand sNodePath
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
				
				'Report any unexpected errors
				If Err.Number < 0 Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  Failed to Expand Node  " + sNodePath + "of Selected Members Tree")		
					Set objMemberSelectionTree=Nothing
					Exit function
				End If
			Next
			
			sNode = sNodePath+"~"+aNode(ubound(aNode))
			'Select the required node
			objMemberSelectionTree.Select sNode
			
			'Report any unexpected errors
            If Err.Number < 0 Then
				Fn_RAC_ProjectMemberSelectionTreeOperations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "FAIL : Failed to select Node  " + sNode + "of Selected Members Tree" )	
			Else
				Fn_RAC_ProjectMemberSelectionTreeOperations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "PASS : Successfully selected  Node  " + sNode + "of Selected Members Tree.")	
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to expand node
		Case "Expand"
			objMemberSelectionTree.Expand sNode
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
			
			'Report unexpected errors
			If Err.Number < 0 Then
				Fn_RAC_ProjectMemberSelectionTreeOperations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "FAIL : Failed to expand Node  " + sNode + "of Selected Members Tree." )	
			Else
				Fn_RAC_ProjectMemberSelectionTreeOperations = True
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "PASS : Successfully expanded  Node  " + sNode + "of Selected Members Tree.")	
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to expand node
		Case "SelectAndExpand"
			pNode=""
			If Fn_RAC_ProjectMemberSelectionTreeOperations("Verify",sNode,"") <> False Then 
				If Instr(sNode,"~") Then
					aNode=Split(sNode,"~")
				Else
					Set objMemberSelectionTree=Nothing
					Fn_RAC_ProjectMemberSelectionTreeOperations=False
				End If
				
				For iCount = 0 To Ubound(aNode)-1
					If iCount=0 Then
						pNode=aNode(iCount)
					Else
						pNode=pNode&"~"&aNode(iCount)
					End If
					If Fn_RAC_ProjectMemberSelectionTreeOperations("Expand",pNode,"")=False Then 
						Set objMemberSelectionTree=nothing
						Exit FUnction
					End If					
				Next

				If Fn_RAC_ProjectMemberSelectionTreeOperations("Select",sNode,"")=False Then 
					Set objMemberSelectionTree=nothing
					Exit FUnction
				Else
					Fn_RAC_ProjectMemberSelectionTreeOperations=True
				End If	
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Verify node
		Case "Verify"
		
			aNode=Split(sNode,"~")
			'Expand the tree till parent node
			For iCounter = 0 To ubound(aNode)-1
				If iCounter=0 Then
					sNodePath=aNode(0)
				Else
					sNodePath=sNodePath+"~"+aNode(iCounter)
				End If
				
				objMemberSelectionTree.Expand sNodePath
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
				
				'Report any unexpected errors
				If Err.Number < 0 Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  Failed to Expand Node  " + sNodePath + "of Selected Members Tree")		
					Set objMemberSelectionTree=Nothing
					Exit function
				End If
			Next
			
			For iCounter=0 to Cint(objMemberSelectionTree.GetROProperty("items count"))-1
				If trim(objMemberSelectionTree.GetItem(iCounter))=trim(sNode) Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "PASS : Node  " + sNode + " found in Selected Members Tree.")	
					Fn_RAC_ProjectMemberSelectionTreeOperations=True
					Exit for
				End If
			Next
			
			If Fn_RAC_ProjectMemberSelectionTreeOperations=False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "PASS : Node  " + sNode + " not found in Selected Members Tree.")	
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Verify node not exist
		Case "VerifyNonExist"
			
			Fn_RAC_ProjectMemberSelectionTreeOperations=False
			aNode=Split(sNode,"~")
			'Expand the tree till parent node
			For iCounter = 0 To ubound(aNode)-1
				If iCounter=0 Then
					sNodePath=aNode(0)
				Else
					sNodePath=sNodePath+"~"+aNode(iCounter)
				End If				
				bFlag=False				
				For iCount=0 to Cint(objMemberSelectionTree.GetROProperty("items count"))-1
					If trim(objMemberSelectionTree.GetItem(iCount))=trim(sNodePath) Then
						bFlag=True
						objMemberSelectionTree.Expand sNodePath
						Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
						Exit for
					Else
						bFlag=False
						Fn_RAC_ProjectMemberSelectionTreeOperations=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If				
				'Report any unexpected errors
				If Err.Number < 0 Then
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>:  Failed to Expand Node  " + sNodePath + "of Selected Members Tree")					
					Set objMemberSelectionTree=Nothing
					Exit function
				End If
			Next
			
			If bFlag=True Then							
				For iCounter=0 to Cint(objMemberSelectionTree.GetROProperty("items count"))-1
					If trim(objMemberSelectionTree.GetItem(iCounter))<>trim(sNode) Then
						Fn_RAC_ProjectMemberSelectionTreeOperations=True
					Else
						Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "Fail : Node  " + sNode + " not found in Selected Members Tree.")													
						Fn_RAC_ProjectMemberSelectionTreeOperations=False
						Exit for
					End If
				Next
			End If
			If Fn_RAC_ProjectMemberSelectionTreeOperations=False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "Fail : Node  " + sNode + " found in Selected Members Tree.")											
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "PASS : Node  " + sNode + " not found in Selected Members Tree.")											
			End If
			' - - - - - - - - - - Pop Up Menu Select
		Case "PopupMenuSelect"
			'Select node
            If Fn_RAC_ProjectMemberSelectionTreeOperations("Select",sNode,"") = False then
				Fn_RAC_ProjectMemberSelectionTreeOperations = False
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "FAIL : Failed to expand Node  " + sNode + "of Selected Members Tree." )	
			End If
			objMemberSelectionTree.OpenContextMenu sNode
			
			'Select Menu action
			Call Fn_UI_JavaMenu_Select("Fn_RAC_ProjectMemberSelectionTreeOperations",JavaWindow("Project"),sPopupMenu)
			
			If Err.number < 0 Then
				Fn_RAC_ProjectMemberSelectionTreeOperations = False
			Else
				Fn_RAC_ProjectMemberSelectionTreeOperations = True
			End If
			
	End Select
	
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_RAC_SetVisible
'
'Function Description	 :	Function used to make teamcenter Application visible\actiavte
'
'Function Parameters	 :  NA						
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Teamcneter application should be available
'
'Function Usage		     :	Call Fn_RAC_SetVisible()
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  26-Oct-2016	    |	 1.0		|	  Prasenjeet P.	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_RAC_SetVisible()
	'Declaring variables
	Dim objDefaultWindow
	Dim iCounter
   'Creating object of Teamcenter Default Window
	Set objDefaultWindow = JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0")
	
   If objDefaultWindow.Exist(2)=False Then
		Exit Function
   End IF
   For iCounter=0 to 3
		If objDefaultWindow.GetROProperty("visible") Then
		   Exit For
		End If
		objDefaultWindow.highlight
		wait 2
		objDefaultWindow.RefreshObject
   Next
End Function