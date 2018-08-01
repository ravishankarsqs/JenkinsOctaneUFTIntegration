'! @Name 			RAC_PSE_SplitBOMTableOperations
'! @Details 		Action word to perform operations on structure manager Frozen Column Table table
'! @InputParam1 	sAction 			: String to indicate what action is to be performed on structure manager Frozen Column Table table e.g. Select, Expand
'! @InputParam2 	sNodeName 			: Node name in structure manager Frozen Column Table table on which action is to be performed
'! @InputParam3 	sColumnName			: Frozen Column Table table column name
'! @InputParam4 	sValue	 			: Value to pass
'! @InputParam5 	sPopupMenu 			: Menu tag name from XML
'! @InputParam5 	iTableInstance 		: Frozen Column Table instance number
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			07 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_SplitBOMTableOperations","RAC_PSE_SplitBOMTableOperations", oneIteration, "VerifyExist","WS_P050050/01;1-AUTWSP_P-Part55511","","","",1
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_SplitBOMTableOperations","RAC_PSE_SplitBOMTableOperations", oneIteration, "Select",DataTable.Value("BOMTableObjectRevisionPath"),"","","",1

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeName,sColumnName,sValue,sPopupMenu,iTableInstance
Dim objChildObjects,objContextMenu,objNodeForRow,objDescription
Dim objFrozenColumnTable,objPSEApplet,objExpandBelow,objExpandBelow1,objNote, objInformation,objBOMTable,objBOMCompare
Dim sObjectTypeName,sTempValue,sColourCode,sColour,sTempNodeName
Dim iRowCounter,iCounter,iColumnIndex,iRows,iCount
Dim aNodeName,aPopupMenu
Dim bFlag,bBOMTableFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sColumnName = Parameter("sColumnName")
sValue = Parameter("sValue")
sPopupMenu = Parameter("sPopupMenu")
iTableInstance = Parameter("iTableInstance")

'Creating obejct of [ BOM Table ]
Set objFrozenColumnTable=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jtbl_FrozenColumnTable","")
Set objBOMTable=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jtbl_BOMTable","")
Set objPSEApplet=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","japt_PSEApplet","")

If iTableInstance="" Then
	iTableInstance=0
Else
	iTableInstance=Cint(iTableInstance)-1
End If

objFrozenColumnTable.SetTOProperty "index",Cint(iTableInstance)
objBOMTable.SetTOProperty "index",Cint(iTableInstance)
objPSEApplet.JavaObject("jobj_BOMPanel").SetTOProperty "index",Cint(iTableInstance)

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_SplitBOMTableOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of BOM table
If Fn_UI_Object_Operations("RAC_PSE_SplitBOMTableOperations","Exist",objFrozenColumnTable,"2","","")= True Then
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	objPSEApplet.JavaObject("jobj_BOMPanel").Click 1,1,"LEFT"
	bBOMTableFlag=False
ElseIf Fn_UI_Object_Operations("RAC_PSE_SplitBOMTableOperations","Exist",objBOMTable,"1","","")= True Then
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	objPSEApplet.JavaObject("jobj_BOMPanel").Click 1,1,"LEFT"
	Set objFrozenColumnTable=objBOMTable
	bBOMTableFlag=True
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & Cstr(sNodeName) & " ] as Frozen Column Table does not exist","","","","","")
	Call Fn_ExitTest()
End if
Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_PSE_SplitBOMTableOperations",sAction,"","")

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to multi select nodes from BOM table
	Case "MultiSelect"
		aNodeName = Split(sNodeName,"^")
		'Clear the already selected Nodes
		objFrozenColumnTable.Object.clearSelection
		For iCounter = 0 to UBound(aNodeName)
			If bBOMTableFlag=True Then
				If inStr(aNodeName(iCounter),"@") > 1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,aNodeName(iCounter))
				Else
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,aNodeName(iCounter))
					If iRowCounter = -1 Then
						iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,aNodeName(iCounter))
					End If
				End If
			Else
				iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,aNodeName(iCounter))		
			End If
			If iRowCounter <> -1 Then
				objFrozenColumnTable.ExtendRow iRowCounter
				sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(aNodeName(iCounter)) & " ] from BOM table","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multiselect node [ " & Cstr(aNodeName(iCounter)) & " ] from BOM table as specified node does not exist","","","","","")
				Call Fn_ExitTest()
			End If
		Next
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_SplitBOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to select node from Frozen Column Table
	Case "ActivateBOMTable"
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to activate BOM table of instance [ " & CStr(iTableInstance+1) & " ] from Split windows due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully activated BOM table of instance [ " & CStr(iTableInstance+1) & " ] from Split windows","","","",GBL_MICRO_SYNC_ITERATIONS,"")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to select node from Frozen Column Table
	Case "Select"
		If bBOMTableFlag=True Then
			If inStr(sNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
				End If
			End If
		Else
			iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)		
		End If
		
		If iRowCounter = -1 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & Cstr(sNodeName) & " ] from Frozen Column Table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		'Selecting node from table
		objFrozenColumnTable.Object.clearSelection  
		objFrozenColumnTable.SelectRow iRowCounter
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ] from Frozen Column Table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_SplitBOMTableOperations",sAction,"","")			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Frozen Column Table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Deselect node from Frozen Column Table
	Case "Deselect"
		If bBOMTableFlag=True Then
			If inStr(sNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
				End If
			End If
		Else
			iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)		
		End If
		
		If iRowCounter = -1 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Deselect node [ " & Cstr(sNodeName) & " ] from Frozen Column Table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		'Selecting node from table
		objFrozenColumnTable.Object.clearSelection  
		objFrozenColumnTable.DeselectRow iRowCounter
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Deselect node [ " & CStr(sNodeName) & " ] from Frozen Column Table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_SplitBOMTableOperations",sAction,"","")			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Deselected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Frozen Column Table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check whether specific node exist in table or not
	Case "Exist"
		If bBOMTableFlag=True Then
			If inStr(sNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
				End If
			End If
		Else
			iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)		
		End If
		
		If iRowCounter <> -1 Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_SplitBOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
		Else
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_SplitBOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to verify specific node exist in table or not
	Case "VerifyExist","VerifyNonExist"
		If bBOMTableFlag=True Then
			If inStr(sNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
				End If
			End If
		Else
			iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)		
		End If
		
		If iRowCounter <> -1 Then
			If sAction="VerifyExist" Then
				sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] is exist under Frozen Column Table","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] is exist under Frozen Column Table","","","","","") 
				Call Fn_ExitTest()
			End If
		Else
			If GBL_LOG_ADDITIONAL_INFORMATION<>"" Then
				sObjectTypeName=GBL_LOG_ADDITIONAL_INFORMATION
			Else
				sObjectTypeName="node"
			End If
			GBL_LOG_ADDITIONAL_INFORMATION=""
			
			If sAction="VerifyNonExist" Then				
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist under Frozen Column Table","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist under Frozen Column Table","","","","","") 
				Call Fn_ExitTest()
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check expand node
	Case "Expand"
		If bBOMTableFlag=True Then
			If inStr(sNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
				End If
			End If
		Else
			iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)		
		End If
		
		If iRowCounter <> -1 Then			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_SplitBOMTableOperations","SelectRow",objFrozenColumnTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Frozen Column Table","","","","","")
				Call Fn_ExitTest()
			End If		
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewExpand"
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & Cstr(sNodeName) & " ] from Frozen Column Table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_SplitBOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check expand node
	Case "ExpandBelow"
		If bBOMTableFlag=True Then
			If inStr(sNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
				End If
			End If
		Else
			iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)		
		End If
		
		If iRowCounter <> -1 Then			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_SplitBOMTableOperations","SelectRow",objFrozenColumnTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Frozen Column Table","","","","","")
				Call Fn_ExitTest()
			End If		
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewExpandBelow"
			
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_SplitBOMTableOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

			'Creating object of [ Expand Below ] dialog
			Set objExpandBelow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_ExpandBelow","")
			Set objExpandBelow1=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_ExpandBelow@2","")
			'Checking existance of [ Expand Below ] dialog
			If Fn_UI_Object_Operations("RAC_PSE_SplitBOMTableOperations","Exist",objExpandBelow,"","","") then
				If Fn_UI_JavaButton_Operations("RAC_PSE_SplitBOMTableOperations","Click",objExpandBelow,"jbtn_Yes")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Frozen Column Table as fail to click on [ Yes ] button of [ Expand Below ] dialog","","","","","")
					Call Fn_ExitTest()
				End If
			ElseIf Fn_UI_Object_Operations("RAC_PSE_SplitBOMTableOperations","Exist",objExpandBelow1,"","","") then
				If Fn_UI_JavaButton_Operations("RAC_PSE_SplitBOMTableOperations","Click",objExpandBelow1,"jbtn_Yes")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Frozen Column Table as fail to click on [ Yes ] button of [ Expand Below ] dialog","","","","","")
					Call Fn_ExitTest()
				End If
			End IF
			'Releasing object of [ Expand Below ] dialog
			Set objExpandBelow=Nothing
			Set objExpandBelow1=Nothing
			
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & Cstr(sNodeName) & " ] from Frozen Column Table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_SplitBOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to collapse node
	Case "Collapse"		
		If bBOMTableFlag=True Then
			If inStr(sNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
				End If
			End If
		Else
			iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)		
		End If
		
		If iRowCounter <> -1 Then			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_SplitBOMTableOperations","SelectRow",objFrozenColumnTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from Frozen Column Table","","","","","")
				Call Fn_ExitTest()
			End If		
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewCollapseBelow"
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse node [ " & Cstr(sNodeName) & " ] from Frozen Column Table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_SplitBOMTableOperations",sAction,"","")	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to validate foregroud\Background color of specific node cell
	Case "VerifyForegroundColour", "VerifyBackgroundColour"
		If sNodeName <> "" Then
			If bBOMTableFlag=True Then
				If inStr(sNodeName,"@") > 1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
				Else
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)
					If iRowCounter = -1 Then
						iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
					End If
				End If
			Else
				iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)		
			End If
			If iRowCounter = -1 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as fail to select node [ " & Cstr(sNodeName) & " ] while performing [ " & Cstr(sAction) & " ] operation under Frozen Column Table","","","","","")
				Call Fn_ExitTest()
			End If
			iCounter = iRowCounter
			iRowCounter = iRowCounter + 1
		Else
			iCounter = 0
			iRowCounter = objFrozenColumnTable.GetROProperty("rows")
		End If

		Do While cInt(iCounter) < cInt(iRowCounter)
			'Creating object of node
			Set objNodeForRow = objFrozenColumnTable.Object.getNodeForRow(cint(iCounter))
			'if background colour
			If sAction = "VerifyBackgroundColour" Then
				sColour = objFrozenColumnTable.Object.getBackground(objNodeForRow,False).toString()
			Else
			'if foreground colour
				sColour = objFrozenColumnTable.Object.getForeground(objNodeForRow,False).toString()
			End If

			sColour =  mid(sColour ,inStr(sColour ,"[")  ,inStr(sColour ,"]") )
			'comparing colour codes RGB
			Select Case cstr(sValue)
				Case "BLACK"
					sColourCode = "[r=0,g=0,b=0]"
				Case "WHITE"
					sColourCode =  "[r=255,g=255,b=255]"
				Case "GRAY"
					sColourCode = "[r=178,g=180,b=191]" 
				Case "DARKGRAY"
					sColourCode = "[r=128,g=128,b=128]"
				Case "DARKBLUE"
					sColourCode = "[r=0,g=0,b=255]" 
				Case "GREEN"
					sColourCode = "[r=80,g=176,b=128]"
				Case "DARKGREEN"
					sColourCode = "[r=0,g=255,b=0]"
				Case "ORANGE"
					sColourCode = "[r=255,g=200,b=0]"
				Case "RED"
					sColourCode = "[r=255,g=0,b=0]" 
				Case "YELLOW"
					sColourCode = "[r=255,g=255,b=0]"
				Case Else
					Call Fn_ExitTest()
			End Select
			if sColour = sColourCode  Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified colour [ " & Cstr(sValue) & " ] of operation [ " & Cstr(sAction) & " ]","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify colour [ " & Cstr(sValue) & " ] of operation [ " & Cstr(sAction) & " ]","","","","","") 
				Call Fn_ExitTest()
			End If
			iCounter = iCounter + 1
			'Releasing object of node
			Set objNodeForRow = nothing
		Loop
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_SplitBOMTableOperations",sAction,"","")	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Imprecise node from BOM table
	Case "Imprecise","Precise"
		sTempNodeName=sNodeName
	
		If bBOMTableFlag=True Then
			If inStr(sNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sNodeName)
				End If
			End If
		Else
			iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sNodeName)		
		End If
		
		bFlag = False
		Select Case sAction
		
			Case "Imprecise"
				If Instr(1,objFrozenColumnTable.Object.getComponentForRow(iRowCounter).getProperty("bl_config_string"),"Precise") Then
					bFlag = True
				Else
					bFlag = False
				End If
			
			Case "Precise"
			
				If Instr(1,objFrozenColumnTable.Object.getComponentForRow(iRowCounter).getProperty("bl_config_string"),"Precise") Then
					bFlag = False
				Else
					bFlag = True
				End If
			
		End Select
		
		If bFlag Then
			
			aNodeName=Split(sTempNodeName,"~")
			If Ubound(aNodeName)=0 Then
				sTempNodeName=aNodeName(0)
			Else
				sTempNodeName=aNodeName(0)
				For iCounter = 1 To Ubound(aNodeName)-1
					sTempNodeName=sTempNodeName & "~" & aNodeName(iCounter)
				Next
			End If
			
			If bBOMTableFlag=True Then
				If inStr(sTempNodeName,"@") > 1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sTempNodeName)
				Else
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objFrozenColumnTable,sTempNodeName)
					If iRowCounter = -1 Then
						iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objFrozenColumnTable,sTempNodeName)
					End If
				End If
			Else
				iRowCounter = Fn_RAC_PSEFrozenColumnTableRowOperations("getnodeindex",objFrozenColumnTable,sTempNodeName)		
			End If
			
			If iRowCounter = -1 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to " & sAction & " node [ " & Cstr(sTempNodeName) & " ] from BOM table as specified node does not exist","","","","","")
				Call Fn_ExitTest()
			End If
			'Selecting node from table
			objFrozenColumnTable.Object.clearSelection  
			objFrozenColumnTable.SelectRow iRowCounter
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditTogglePreciseImprecise"
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations", oneIteration, "Select", "FileSave"
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_SplitBOMTableOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to " & sAction & " node [ " & CStr(sNodeName) & " ] from BOM table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_SplitBOMTableOperations",sAction,"","")			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully " & sAction & " [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Update mass properties
	Case "SplitBOMWindow"
		If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations", "Click", objPSEApplet,"jbtn_SplitBOMTableWindow") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to split BOM table as fail to click on [ Split window ] button","","","","","")
			Call Fn_ExitTest()	
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully split BOM table window","","","","","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Update mass properties
	Case "BOMCompare"
		
		'Call menu operation to BOM compare
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsCompare"
		'Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_SplitBOMTableOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

		'Create object of information window
		Set objBOMCompare = Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_BOMCompare","")
		
		'Select mode
		
		if sValue<>"" Then
			If Fn_UI_JavaList_Operations("RAC_PSE_BOMTableOperations", "Select", objBOMCompare,"jlst_Mode",sValue, "", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on BOM Compare dialog as failed to select mode [ " & Cstr(sValue) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		
		If sPopupMenu<>"" Then
			If Fn_UI_JavaCheckBox_Operations("RAC_PSE_BOMTableOperations", "set", objBOMCompare, "jchk_Report", sPopupMenu)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on BOM Compare dialog as failed to set/unset report option","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		
			
		If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations", "Click", objBOMCompare,"jbtn_Apply") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [" & sAction & "] as failed to click on Apply button of BOM Compare dialog","","","","","")
			Call Fn_ExitTest()	
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations", "Click", objBOMCompare,"jbtn_Cancel") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [" & sAction & "] as failed to click on Cancel button of BOM Compare dialog","","","","","")
			Call Fn_ExitTest()	
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully performed BOM Compare with Mode [ " & Cstr(sValue) & " ]","","","","","")
		
		Set objBOMCompare = Nothing
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

If Err.number <> 0  Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on Frozen Column Table due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing Objects
Set objFrozenColumnTable =Nothing
Set objPSEApplet = Nothing

Function Fn_ExitTest()
	'Releasing Objects
	Set objFrozenColumnTable =Nothing
	Set objPSEApplet = Nothing
	ExitTest
End Function
