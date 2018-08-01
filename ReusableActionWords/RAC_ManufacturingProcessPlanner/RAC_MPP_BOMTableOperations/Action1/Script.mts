'! @Name 			RAC_MPP_BOMTableOperations
'! @Details 		Action word to perform operations on Manufacturing Process Planner bom table
'! @InputParam1 	sAction 			: String to indicate what action is to be performed on Manufacturing Process Planner bom table e.g. Select, Expand
'! @InputParam2 	sNodeName 			: Node name in Manufacturing Process Planner bom table on which action is to be performed
'! @InputParam3 	sColumnName			: BOM table column name
'! @InputParam4 	sValue	 			: Value to pass
'! @InputParam5 	sPopupMenu 			: Menu tag name from XML
'! @InputParam6 	sTabName 			: Table tab name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			07 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_MPP_BOMTableOperations","RAC_MPP_BOMTableOperations", oneIteration, "VerifyExist","WS_P050050/01;1-AUTWSP_P-Part55511","","","",""
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_MPP_BOMTableOperations","RAC_MPP_BOMTableOperations", oneIteration, "Select",DataTable.Value("BOMTableObjectRevisionPath"),"","","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeName,sColumnName,sValue,sPopupMenu,sTabName
Dim objChildObjects,objContextMenu,objNodeForRow,objDescription
Dim objBOMTable,objPSEApplet,objExpandBelow,objExpandBelow1,objMPPApplet,objManufacturingProcessPlanner
Dim sObjectTypeName,sTempValue,sColourCode,sColour
Dim iRowCounter,iCounter,iColumnIndex,iRows,iCount,iTabIndex
Dim aNodeName,aPopupMenu
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sColumnName = Parameter("sColumnName")
sValue = Parameter("sValue")
sPopupMenu = Parameter("sPopupMenu")
sTabName = Parameter("sPopupMenu")

'Creating obejct of [ BOM Table ]
Set objManufacturingProcessPlanner=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ManufacturingProcessPlanner_OR","jwnd_ManufacturingProcessPlanner","")
'Creating object of [ MPPApplet ] applet
Set objMPPApplet=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ManufacturingProcessPlanner_OR","wjapt_MPPApplet","")
'Creating obejct of [ BOM Table ]
Set objBOMTable=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jtbl_CMEBOMTreeTable","")

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

If sTabName<>"" Then	
	iTabIndex = cInt(objManufacturingProcessPlanner.JavaObject("jobj_RACTabFolderWidget").Object.getSelectedTabIndex)
	sTabName = objManufacturingProcessPlanner.JavaObject("jobj_RACTabFolderWidget").Object.getItem(iTabIndex ).text
	If sTabName=Cstr(objManufacturingProcessPlanner.JavaObject("jobj_RACTabFolderWidget").Object.getItem(iTabIndex ).text) Then
		For iCounter=iTabIndex to 0 STEP -1
			objMPPApplet.SetTOProperty "index",iTabIndex
			If objBOMTable.Exist(1) Then
				Exit for
			End If
		Next
	End If
Else
	LoadAndRunAction "RAC_ManufacturingProcessPlanner\RAC_MPP_SetIndexOfMPPAppletFromTab","RAC_MPP_SetIndexOfMPPAppletFromTab",OneIteration
	DataTable.SetCurrentRow 1		
	If DataTable.Value("ReusableActionWordReturnValue","Global")= "False" Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & Cstr(sNodeName) & " ] as BOM table does not exist","","","","","")
		Call Fn_ExitTest()
	End If
End If
	
'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_MPP_BOMTableOperations",sAction,"","")

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to select node from BOM table
	Case "Select"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		If iRowCounter = -1 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & Cstr(sNodeName) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		'Selecting node from table
		objBOMTable.Object.clearSelection  
		objBOMTable.SelectRow iRowCounter
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ] from BOM table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Deselect node from BOM table
	Case "Deselect"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		If iRowCounter = -1 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Deselect node [ " & Cstr(sNodeName) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		'Selecting node from table
		objBOMTable.Object.clearSelection  
		objBOMTable.DeselectRow iRowCounter
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Deselect node [ " & CStr(sNodeName) & " ] from BOM table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Deselected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to multi select nodes from BOM table
	Case "MultiSelect"
		aNodeName = Split(sNodeName,"^")
		'Clear the already selected Nodes
		objBOMTable.Object.clearSelection
		For iCounter = 0 to UBound(aNodeName)
			iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,Trim(aNodeName(iCounter)))
			If iRowCounter <> -1 Then
				objBOMTable.ExtendRow iRowCounter
				sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(aNodeName(iCounter)) & " ] from BOM table","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multiselect node [ " & Cstr(aNodeName(iCounter)) & " ] from BOM table as specified node does not exist","","","","","")
				Call Fn_ExitTest()
			End If
		Next
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to multi select nodes from BOM table
	Case "SelectAll"
		'Clear the already selected Nodes
		objBOMTable.Object.clearSelection
		For iCounter = 0 to cInt(objBOMTable.GetROProperty ("rows")) - 1
            objBOMTable.ExtendRow iCounter 
		Next
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select all rows\nodes from BOM table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected all rows\nodes of BOM table","","","","","")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check whether specific node exist in table or not
	Case "Exist"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		If iRowCounter <> -1 Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_BOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
		Else
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_BOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to verify specific node exist in table or not
	Case "VerifyExist","VerifyNonExist"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)			
		If iRowCounter <> -1 Then
			If sAction="VerifyExist" Then
				sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] is exist under BOM table","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] is exist under BOM table","","","","","") 
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
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist under BOM table","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist under BOM table","","","","","") 
				Call Fn_ExitTest()
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to verify specific node exist in table or not
	Case "VerifyNodeRowNumber"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)			
		If iRowCounter <> -1 Then	
			If Cstr(iRowCounter)=Cstr(sValue) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sNodeName) & " ] node is available on expected row [ " & Cstr(sValue) & " ] under BOM table","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sNodeName) & " ] node does not available on expected row [ " & Cstr(sValue) & " ] under BOM table","","","","","") 
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sNodeName) & " ] does not exist under BOM table","","","","","") 
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check whether specific node exist in table or not
	Case "ExistExt"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		If iRowCounter <> -1 Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_BOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
		Else
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_BOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check expand node
	Case "Expand"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)		
		If iRowCounter <> -1 Then			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If		
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewExpandOptionsExpand"
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & Cstr(sNodeName) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to collapse node
	Case "Collapse"		
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)		
		If iRowCounter <> -1 Then			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If		
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewCollapseBelow"
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse node [ " & Cstr(sNodeName) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to edit\modify value of specific cell
	Case "EditCellValue"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If 
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","LEFT","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Modifing value
			IF Fn_UI_JavaEdit_Operations("RAC_MPP_BOMTableOperations","Set",objMPPApplet.JavaEdit("jedt_BOMTableField"),"", sValue)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit field [ " & Cstr(sColumnName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End IF
			
			objMPPApplet.JavaEdit("jedt_BOMTableField").Activate
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited\modified field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Fail to edit\modify field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] as node [ " & Cstr(sNodeName) & " ] does not exist under BOM table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
	'Case to edit\modify value of list from specific cell
	Case "EditCellListValue"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] while editing cell list value from BOM table","","","","","")
				Call Fn_ExitTest()
			End If 
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","LEFT","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] while editing cell list value from BOM table","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
						
			Select Case sColumnName
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "Unit Of Measure"
					objMPPApplet.JavaButton("jbtn_DropDown").Click micLeftBtn
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					
					'Creating Description object of LOV tree table
					Set objDescription = Description.Create()						
					objDescription("Class Name").value = "JavaTable"
					objDescription("class_path").value = ".*LOVTreeTable.*"
					objDescription("class_path").RegularExpression = true
					Set objChildObjects = objMPPApplet.ChildObjects(objDescription)
					
					bFlag=False					
					If objChildObjects.Count > 0 Then							
						For iCounter=0 to objChildObjects(0).GetROProperty("rows")-1
							If Trim(sValue)=Trim(objChildObjects(0).Object.getValueAt(iCounter,0).getDisplayableValue()) Then
								objChildObjects(0).DoubleClickCell iCounter,0
								bFlag=True
								Exit for
							End If
						Next
					End If
					
					If bFlag=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit\modify field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
						Call Fn_ExitTest()
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited\modified field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")						
					End If
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case Else
					objMPPApplet.JavaButton("jbtn_DropDown").Click micLeftBtn
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					
					'Creating Description object of static text
					Set objDescription=Description.Create()
					objDescription("Class Name").value = "JavaStaticText"
					objDescription("label").value = sValue
					Set  objChildObjects =  objMPPApplet.ChildObjects(objDescription)
					For iCounter = 0 to objChildObjects.count - 1
						If objChildObjects(iCounter).toString()  = "[ " & sValue & "(st) ] text label" Then
							objChildObjects(iCounter).Click 1,1
							Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)							
							objMPPApplet.JavaEdit("jedt_BOMTableListField").Activate
							Exit for
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit\modify field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
							Call Fn_ExitTest()
						End If
					Next
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited\modified field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
			End Select
			'Releasing required objects	
			Set objChildObjects = nothing
			Set objDescription = nothing
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] of node [ " & Cstr(sNodeName) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellValue"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")
			sTempValue = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","GetCellData",objBOMTable,"",iRowCounter,sColumnName,"","","")))
			If sTempValue = Trim(cstr(sValue)) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(sColumnName) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
			Else
				If isNumeric(sTempValue) Then
					If cstr(Abs(sTempValue)) = Trim(cstr(sValue)) Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(sColumnName) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
						Call Fn_ExitTest()
					End  If
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
					Call Fn_ExitTest()
				End If
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellValueInStr"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		If iRowCounter <> -1 Then
			'objBOMTable.SelectRow iRowCounter 
			sTempValue = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","GetCellData",objBOMTable,"",iRowCounter,sColumnName,"","","")))
			If inStr(1,sTempValue,Trim(cstr(sValue))) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(sColumnName) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get data of specific cell
	Case "GetCellData"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_BOMTableOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		
		If iRowCounter <> -1 Then
			DataTable.Value("ReusableActionWordReturnValue","Global") = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","GetCellData",objBOMTable,"",iRowCounter,sColumnName,"","","")))
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellListValue"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
				Call Fn_ExitTest()
			End If 			
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","LEFT","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			'Clicking on drop down button
			objMPPApplet.JavaButton("jbtn_DropDown").Click micLeftBtn
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			'creating object of java static text
			Set objDescription=description.Create()
			objDescription("Class Name").value = "JavaStaticText"
			objDescription("label").value = sValue
			Set objChildObjects = objMPPApplet.ChildObjects(objDescription)
			
			For iCounter = 0 to objChildObjects.count - 1
				If objChildObjects(iCounter).toString()  = "[ " & sValue & "(st) ] text label" then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(sColumnName) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Exit For
				End If
			Next
			'Releasing objects
			Set objChildObjects = nothing
			Set objDescription = nothing
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB on specific node
	Case "PopupMenuSelect"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MPP_BOMTableOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu on node [" & Cstr(sNodeName) & "] as user passed invalid RMB menu [ False ] under BOM table" ,"","","","","")
			Call Fn_ExitTest()
		End If
		'Clear already selected nodes
		objBOMTable.Object.clearSelection
		
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)

		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")
			'Split Context menu to Build Path
			aPopupMenu = Split(sPopupMenu,":",-1,1)
			'Right click on node to open RMB menu
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Creating object of [ Context menu ]
			Set objContextMenu=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","wmnu_ContextMenu","")			
			
			Select Case Ubound(aPopupMenu)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "0"
					sPopupMenu = objContextMenu.BuildMenuPath(aMenu(0))
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
				Case "1"
					sPopupMenu = objContextMenu.BuildMenuPath(aMenu(0),aMenu(1))					
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
			End Select
			'Select RMB menu
			objContextMenu.Select sPopupMenu
			'Creating object of [ Context menu ]
			Set objContextMenu=Nothing
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on node [ " & Cstr(sNodeName) & " ] as specified node does not exist under BOM table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB on specific multiple nodes
	Case "MultiSelectPopupMenuSelect"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MPP_BOMTableOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu on node [" & Cstr(sNodeName) & "] as user passed invalid RMB menu [ False ] under BOM table" ,"","","","","")
			Call Fn_ExitTest()
		End If
		aNodeName = split(sNodeName , "^")
		'Clear the already selected Nodes
		objBOMTable.Object.clearSelection
		For iCounter = 0 to UBound(aNodeName)
			iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,aNodeName(iCounter))
			If iRowCounter <> -1 Then
				objBOMTable.ExtendRow iRowCounter 
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on node [ " & Cstr(aNodeName(iCounter)) & " ] as specified node does not exist under BOM table","","","","","")
				Call Fn_ExitTest()
			End If
		Next

		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,Trim(aNodeName(UBound(aNodeName))))
	
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("MPPBomTable","getitemtypename")
			'Split Context menu to Build Path Accordingly
			aPopupMenu = split(sPopupMenu,":",-1,1)
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			'Creating object of [ Context menu ]
			Set objContextMenu=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","wmnu_ContextMenu","")			
			
			Select Case Ubound(aPopupMenu)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "0"
					sPopupMenu = objContextMenu.BuildMenuPath(aMenu(0))
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
				Case "1"
					sPopupMenu = objContextMenu.BuildMenuPath(aMenu(0),aMenu(1))					
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
			End Select
			'Select RMB menu
			objContextMenu.Select sPopupMenu
			'Creating object of [ Context menu ]
			Set objContextMenu=Nothing
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected popup menu [ " & Cstr(sPopupMenu) & " ] on multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on multi selected nodes [ " & Cstr(sNodeName) & " ] as specified node does not exist under BOM table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify RMB menu available on specific node
	Case "VerifyPopupMenuExists","VerifyPopupMenuNonExists"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MPP_BOMTableOperations","",sPopupMenu)
		
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		
		If iRowCounter <> -1 Then
			'Split Context menu to Build Path Accordingly
			aPopupMenu = split(sPopupMenu,":",-1,1)
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as unable to right click on correct cell while performing [ " & Cstr(sAction) & " ] opearation","","","","","")					
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as unable to right click on correct cell while performing [ " & Cstr(sAction) & " ] opearation","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Creating object of [ Context menu ]
			Set objContextMenu=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","wmnu_ContextMenu","")
			
			Select Case cInt(Ubound(aPopupMenu))
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case 0
					If objContextMenu.GetItemProperty(Replace(sPopupMenu,":",";"),"Exists") <> False Then
						If sAction="VerifyPopupMenuExists" Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified popup menu [ " & Cstr(sPopupMenu) & " ] exist of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification as popup menu [ " & Cstr(sPopupMenu) & " ] exist of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
							Call Fn_ExitTest()
						End If
					Else
						If sAction="VerifyPopupMenuExists" Then								
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification as popup menu [ " & Cstr(sPopupMenu) & " ] does exist of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
							Call Fn_ExitTest()
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified popup menu [ " & Cstr(sPopupMenu) & " ] does not exist of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
						End If
					End If
					
					If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","LEFT","") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as unable to click on correct cell while performing [ " & Cstr(sAction) & " ] opearation","","","","","") 
						Call Fn_ExitTest()
					End If
			End Select
			'Releasing object of [ Context menu ]
			Set objContextMenu = nothing						
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail while performing [ " & Cstr(sAction) & " ] opearation","","","","","") 
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify RMB menu is active for specific node
	Case "VerifyPopupMenuEnabled"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MPP_BOMTableOperations","",sPopupMenu)
		
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)

		If iRowCounter <> -1 Then
			'Split Context menu to Build Path Accordingly
			aPopupMenu = split(sPopupMenu,":",-1,1)		
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as unable to right click on correct cell while performing [ " & Cstr(sAction) & " ] opearation","","","","","")					
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as unable to right click on correct cell while performing [ " & Cstr(sAction) & " ] opearation","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Creating object of [ Context menu ]
			Set objContextMenu=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","wmnu_ContextMenu","")
			
			Select Case cInt(Ubound(aPopupMenu))
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case 0
					If objContextMenu.CheckItemProperty (sPopupMenu, "Exists",true,10) Then
						If Cbool(objContextMenu.CheckItemProperty (sPopupMenu, "Enabled",true,10))<>False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified popup menu [ " & Cstr(sPopupMenu) & " ] is enabled of node [ " & Cstr(sNodeName) & " ]  under BOM table","","","","","") 
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification as popup menu [ " & Cstr(sPopupMenu) & " ] does not enabled of node [ " & Cstr(sNodeName) & " ]  under BOM table","","","","","") 
							Call Fn_ExitTest()
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : verification as popup menu [ " & Cstr(sPopupMenu) & " ] does exist of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
						Call Fn_ExitTest()
					End IF
					
					If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","LEFT","") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as unable to click on correct cell while performing [ " & Cstr(sAction) & " ] opearation","","","","","") 
						Call Fn_ExitTest()
					End If					
			End Select
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
			Set objContextMenu = nothing
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail while performing [ " & Cstr(sAction) & " ] opearation","","","","","") 
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get child\sub nodes of specific node
	Case "GetChildItems"
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		iColumnIndex=Fn_RAC_MPPBomTableColumnOperations("getcolumnindex",objBOMTable,"BOM Line")
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1        
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_BOMTableOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
			
		If iRowCounter <> -1 Then
			iRowCounter=iRowCounter+1
			sTempValue=objBOMTable.Object.getPathForRow(iRowCounter).toString()
			sNodeName=""
			For iCounter = iRowCounter To objBOMTable.GetROProperty("rows")-1
				If sTempValue=objBOMTable.Object.getPathForRow(iCounter).getParentPath().toString() Then
					If sNodeName="" Then
						sNodeName= objBOMTable.Object.getValueAt(iCounter,iColumnIndex).toString()
					Else		
						sNodeName= sNodeName & "^" & objBOMTable.Object.getValueAt(iCounter,iColumnIndex).toString()
					End If		
				End If
			Next
			If sNodeName <> "" Then
				DataTable.Value("ReusableActionWordReturnValue","Global")= sNodeName
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get child\sub nodes of specific name under specific node
	Case "GetChildItemsByName"
		aNodeName=Split(sNodeName,"^")
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,aNodeName(0))
		iColumnIndex=Fn_RAC_MPPBomTableColumnOperations("getcolumnindex",objBOMTable,"BOM Line")
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1        
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_BOMTableOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
	
		If iRowCounter <> -1 Then
			iRowCounter=iRowCounter+1
			sTempValue=objBOMTable.Object.getPathForRow(iRowCounter).getParentPath().toString()
			sNodeName =""
			For iCounter = iRowCounter To objBOMTable.GetROProperty("rows")-1
				If sTempValue=objBOMTable.Object.getPathForRow(iCounter).getParentPath().toString() Then
					If inStr(1,objBOMTable.Object.getValueAt(iCounter,iColumnIndex).toString(),aNodeName(1)) Then
						sNodeName=objBOMTable.Object.getValueAt(iCounter,iColumnIndex).toString()
						Exit for
					End If
				Else
					Exit For	
				End If
			Next
			If sNodeName <> "" Then
				DataTable.Value("ReusableActionWordReturnValue","Global")= sNodeName
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get currently selected row number\index
	Case "GetSelectedRowIndex"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1        
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_BOMTableOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		
		For iCounter = 0  to CInt(objBOMTable.Object.getRowCount) - 1 
			If objBOMTable.Object.isRowSelected(iCounter) Then 
				DataTable.Value("ReusableActionWordReturnValue","Global")= iCounter
				Exit for
			End if
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get all selected row path
	Case "GetSelectedRowPaths"
		sTempValue = ""
		sNodeName =""
		For iCounter =0  to CInt(objBOMTable.Object.getRowCount) - 1 
			If objBOMTable.Object.isRowSelected(iCounter) Then
				sNodeName =objBOMTable.Object.getPathForRow(iCounter).toString()
				sNodeName = Right(sNodeName, (Len(sNodeName)-inStr(1, sNodeName, ",", 1)))					
				sNodeName = Left(sNodeName, Len(sNodeName)-1)
				
				If inStr(sNodeName, "@BOM::") > 0 Then
					sNodeName = Trim(replace(sNodeName,"""",""))
					aNodeName = split(sNodeName,",")
					sNodeName = ""
					For iCount = 0 to uBound(aNodeName)
						aNodeName(iCount) = Left(aNodeName(iCount), inStr(aNodeName(iCount),"@")-1)
						If sNodeName = "" Then
							sNodeName = Trim(aNodeName(iCount))
						Else
							sNodeName = sNodeName & ", " & Trim(aNodeName(iCount))
						End If
					Next
				End If
				sNodeName = Trim(replace(sNodeName,", ",":"))
				If sTempValue = "" Then
					sTempValue = sNodeName
				Else
					sTempValue = sTempValue & "~" & sNodeName
				End If
			End if
		Next

		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1        
		DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_BOMTableOperations"
		
		If sTempValue = ""  Then
			DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		Else
			DataTable.Value("ReusableActionWordReturnValue","Global")= sTempValue
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to validate foregroud\Background color of specific node cell
	Case "VerifyForegroundColour", "VerifyBackgroundColour"
		If sNodeName <> "" Then
			iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as fail to select node [ " & Cstr(sNodeName) & " ] while performing [ " & Cstr(sAction) & " ] operation under BOM table","","","","","")
				Call Fn_ExitTest()
			End If
			iCounter = iRowCounter
			iRowCounter = iRowCounter + 1
		Else
			iCounter = 0
			iRowCounter = objBOMTable.GetROProperty("rows")
		End If

		Do While cInt(iCounter) < cInt(iRowCounter)
			'Creating object of node
			Set objNodeForRow = objBOMTable.Object.getNodeForRow(cint(iCounter))
			'if background colour
			If sAction = "VerifyBackgroundColour" Then
				sColour = objBOMTable.Object.getBackground(objNodeForRow,False).toString()
			Else
			'if foreground colour
				sColour = objBOMTable.Object.getForeground(objNodeForRow,False).toString()
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
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to validate foregroud\Background color of specific node cell
	Case "PopupMenuSelectExt"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_MPP_BOMTableOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu on node [" & Cstr(sNodeName) & "] as user passed invalid RMB menu [ False ] under BOM table" ,"","","","","")
			Call Fn_ExitTest()
		End If
		'Clear already selected nodes
		objBOMTable.Object.clearSelection
		
		iRowCounter = Fn_RAC_MPPBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)

		If iRowCounter <> -1 Then
			'Right click on node to open RMB menu
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_MPP_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Selecting java menu
			If Fn_UI_JavaMenu_Operations("RAC_MPP_BOMTableOperations","Select",JavaWindow("jwnd_StructureManager"),sPopupMenu)=False Then				
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on node [ " & Cstr(sNodeName) & " ] as specified node does not exist under BOM table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MPP_BOMTableOperations",sAction,"","")
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

If Err.number <> 0  Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on BOM table due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing Objects
Set objBOMTable =Nothing
Set objMPPApplet = Nothing

Function Fn_ExitTest()
	'Releasing Objects
	Set objBOMTable =Nothing
	Set objMPPApplet = Nothing
	ExitTest
End Function

