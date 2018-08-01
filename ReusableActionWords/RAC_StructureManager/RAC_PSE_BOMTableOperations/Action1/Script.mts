'! @Name 			RAC_PSE_BOMTableOperations
'! @Details 		Action word to perform operations on structure manager bom table
'! @InputParam1 	sAction 			: String to indicate what action is to be performed on structure manager bom table e.g. Select, Expand
'! @InputParam2 	sNodeName 			: Node name in structure manager bom table on which action is to be performed
'! @InputParam3 	sColumnName			: BOM table column name
'! @InputParam4 	sValue	 			: Value to pass
'! @InputParam5 	sPopupMenu 			: Menu tag name from XML
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			07 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "VerifyExist","WS_P050050/01;1-AUTWSP_P-Part55511","","",""
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "Select",DataTable.Value("BOMTableObjectRevisionPath"),"","",""
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "ExpandBelowToLevelAndCollapseLowerLevel",DataTable.Value("BOMTableObjectRevisionPath"),"","2",""
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "SetSortingOrderAndVerifySortOrder",DataTable.Value("BOMTableObjectRevisionPath"),"","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeName,sColumnName,sValue,sPopupMenu
Dim objChildObjects,objContextMenu,objNodeForRow,objDescription, objSnapShot
Dim objBOMTable,objPSEApplet,objExpandBelow,objExpandBelow1,objNote, objInformation
Dim sObjectTypeName,sTempValue,sColourCode,sColour,sTempNodeName
Dim iRowCounter,iCounter,iColumnIndex,iRows,iCount
Dim aNodeName,aPopupMenu
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sColumnName = Parameter("sColumnName")
sValue = Parameter("sValue")
sPopupMenu = Parameter("sPopupMenu")

'Creating obejct of [ BOM Table ]
Set objBOMTable=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jtbl_BOMTable","")
Set objPSEApplet=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","japt_PSEApplet","")
Set objNote=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_Note","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of BOM table
If Fn_UI_Object_Operations("RAC_PSE_BOMTableOperations","Exist",objBOMTable,"","","")= True Then
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	objPSEApplet.JavaObject("jobj_BOMPanel").Click 1,1,"LEFT"
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & Cstr(sNodeName) & " ] as BOM table does not exist","","","","","")
	Call Fn_ExitTest()
End if
Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_PSE_BOMTableOperations",sAction,"","")

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to select node from BOM table
	Case "Select"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
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
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","DONOTSYNC","")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Deselect node from BOM table
	Case "Deselect"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
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
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Deselected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","DONOTSYNC","")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to multi select nodes from BOM table
	Case "MultiSelect"
		aNodeName = Split(sNodeName,"^")
		'Clear the already selected Nodes
		objBOMTable.Object.clearSelection
		For iCounter = 0 to UBound(aNodeName)
			If inStr(trim(aNodeName(iCounter)),"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,Trim(aNodeName(iCounter)))
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,Trim(aNodeName(iCounter)))
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,Trim(aNodeName(iCounter)))
				End If
			End If
			If iRowCounter <> -1 Then
				objBOMTable.ExtendRow iRowCounter
				sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(aNodeName(iCounter)) & " ] from BOM table","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multiselect node [ " & Cstr(aNodeName(iCounter)) & " ] from BOM table as specified node does not exist","","","","","")
				Call Fn_ExitTest()
			End If
		Next
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
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
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected all rows\nodes of BOM table","","","","DONOTSYNC","")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check whether specific node exist in table or not
	Case "Exist"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		If iRowCounter <> -1 Then
			Call Fn_CommonUtil_DataTableOperations("\AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
		Else
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to verify specific node exist in table or not
	Case "VerifyExist","VerifyNonExist"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If			
		
		If iRowCounter <> -1 Then
			If sAction="VerifyExist" Then
				sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] is exist under BOM table","","","","DONOTSYNC","") 
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
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist under BOM table","","","","DONOTSYNC","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist under BOM table","","","","","") 
				Call Fn_ExitTest()
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check whether specific node exist in table or not
	Case "ExistExt"
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
		If iRowCounter <> -1 Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
		Else
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check expand node
	Case "Expand"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		
		If iRowCounter <> -1 Then			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If		
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewExpand"
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & Cstr(sNodeName) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check expand node
	Case "ExpandBelow"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		
		If iRowCounter <> -1 Then			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If		
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewExpandBelow"
			
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

			'Creating object of [ Expand Below ] dialog
			Set objExpandBelow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_ExpandBelow","")
			Set objExpandBelow1=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_ExpandBelow@2","")
			'Checking existance of [ Expand Below ] dialog
			If Fn_UI_Object_Operations("RAC_PSE_BOMTableOperations","Exist",objExpandBelow,"","","") then
				If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations","Click",objExpandBelow,"jbtn_Yes")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table as fail to click on [ Yes ] button of [ Expand Below ] dialog","","","","","")
					Call Fn_ExitTest()
				End If
			ElseIf Fn_UI_Object_Operations("RAC_PSE_BOMTableOperations","Exist",objExpandBelow1,"","","") then
				If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations","Click",objExpandBelow1,"jbtn_Yes")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table as fail to click on [ Yes ] button of [ Expand Below ] dialog","","","","","")
					Call Fn_ExitTest()
				End If
			End IF
			'Releasing object of [ Expand Below ] dialog
			Set objExpandBelow=Nothing
			Set objExpandBelow1=Nothing
			
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & Cstr(sNodeName) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to collapse node
	Case "Collapse"		
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		
		If iRowCounter <> -1 Then			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If		
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewCollapseBelow"
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to collapse node [ " & Cstr(sNodeName) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to edit\modify value of specific cell
	Case "EditCellValue"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
			
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If 
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","LEFT","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Modifing value
			IF Fn_UI_JavaEdit_Operations("RAC_PSE_BOMTableOperations","Set",objPSEApplet.JavaEdit("jedt_BOMTableField"),"", sValue)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit field [ " & Cstr(sColumnName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End IF
			
			objPSEApplet.JavaEdit("jedt_BOMTableField").Activate
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited\modified field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Fail to edit\modify field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] as node [ " & Cstr(sNodeName) & " ] does not exist under BOM table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to edit\modify noe of specific cell
	Case "EditCellNoteValue"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
			
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If 
			Call Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","clickcell",objBOMTable,"",iRowCounter,sColumnName,"","LEFT","")
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","doubleclickcell",objBOMTable,"",iRowCounter,sColumnName,"","LEFT","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Modifing value
			IF Fn_UI_JavaEdit_Operations("RAC_PSE_BOMTableOperations","Set",objNote.JavaEdit("jedt_Note"),"", sValue)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit field [ " & Cstr(sColumnName) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End IF
			
			objNote.JavaButton("jbtn_OK").Click micLeftBtn
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited\modified field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Fail to edit\modify field [ " & Cstr(sColumnName) & " ] to value [ " & Cstr(sValue) & " ] as node [ " & Cstr(sNodeName) & " ] does not exist under BOM table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
	'Case to edit\modify value of list from specific cell
	Case "EditCellListValue"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
			
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] while editing cell list value from BOM table","","","","","")
				Call Fn_ExitTest()
			End If 
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","LEFT","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] while editing cell list value from BOM table","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
						
			Select Case sColumnName
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "Unit Of Measure","WBS Code"
					objPSEApplet.JavaButton("jbtn_DropDown").Click micLeftBtn
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					
					'Creating Description object of LOV tree table
					Set objDescription = Description.Create()						
					objDescription("Class Name").value = "JavaTable"
					objDescription("class_path").value = ".*LOVTreeTable.*"
					objDescription("class_path").RegularExpression = true
					Set objChildObjects = objPSEApplet.ChildObjects(objDescription)
					
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
					objPSEApplet.JavaButton("jbtn_DropDown").Click micLeftBtn
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					
					'Creating Description object of static text
					Set objDescription=Description.Create()
					objDescription("Class Name").value = "JavaStaticText"
					objDescription("label").value = sValue
					Set  objChildObjects =  objPSEApplet.ChildObjects(objDescription)
					For iCounter = 0 to objChildObjects.count - 1
						If objChildObjects(iCounter).toString()  = "[ " & sValue & "(st) ] text label" Then
							objChildObjects(iCounter).Click 1,1
							Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)							
							objPSEApplet.JavaEdit("jedt_BOMTableListField").Activate
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
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellValue"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		If iRowCounter <> -1 Then
		
			'Add the column for which cell value is to be verified
			LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "AddColumn","",sColumnName,"",""
			
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")
			sTempValue = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","GetCellData",objBOMTable,"",iRowCounter,sColumnName,"","","")))
			If sTempValue = Trim(cstr(sValue)) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(Parameter("sColumnName")) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","DONOTSYNC","")
			Else
				If isNumeric(sTempValue) Then
					If cstr(Abs(sTempValue)) = Trim(cstr(sValue)) Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(Parameter("sColumnName")) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","DONOTSYNC","")
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(Parameter("sColumnName")) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
						Call Fn_ExitTest()
					End  If
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(Parameter("sColumnName")) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
					Call Fn_ExitTest()
				End If
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(Parameter("sColumnName")) & " ] does not contain value [ " & Cstr(sValue) & " ] of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellNotEmpty"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")
			sTempValue = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","GetCellData",objBOMTable,"",iRowCounter,sColumnName,"","","")))
			If sTempValue <> "" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(Parameter("sColumnName")) & " ] contains value [ " & Cstr(sTempValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(Parameter("sColumnName")) & " ] does not contain any value of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(Parameter("sColumnName")) & " ] does not contain any value of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellValueInStr"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		If iRowCounter <> -1 Then
			'objBOMTable.SelectRow iRowCounter 
			'Add the column for which cell value is to be verified
			LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "AddColumn","",sColumnName,"",""
			sTempValue = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","GetCellData",objBOMTable,"",iRowCounter,sColumnName,"","","")))
			If inStr(1,sTempValue,Trim(cstr(sValue))) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(sColumnName) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get data of specific cell
	Case "GetCellData"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1
		DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		
		If iRowCounter <> -1 Then
			DataTable.Value("ReusableActionWordReturnValue","Global") = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","GetCellData",objBOMTable,"",iRowCounter,sColumnName,"","","")))
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellListValue"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
				Call Fn_ExitTest()
			End If 			
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","LEFT","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			'Clicking on drop down button
			objPSEApplet.JavaButton("jbtn_DropDown").Click micLeftBtn
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			'creating object of java static text
			Set objDescription=description.Create()
			objDescription("Class Name").value = "JavaStaticText"
			objDescription("label").value = sValue
			Set objChildObjects = objPSEApplet.ChildObjects(objDescription)
			
			For iCounter = 0 to objChildObjects.count - 1
				If objChildObjects(iCounter).toString()  = "[ " & sValue & "(st) ] text label" then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(sColumnName) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","DONOTSYNC","")
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
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB on specific node
	Case "PopupMenuSelect"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_PSE_BOMTableOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu on node [" & Cstr(sNodeName) & "] as user passed invalid RMB menu [ False ] under BOM table" ,"","","","","")
			Call Fn_ExitTest()
		End If
		'Clear already selected nodes
		objBOMTable.Object.clearSelection
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If

		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")
			'Split Context menu to Build Path
			aPopupMenu = Split(sPopupMenu,":",-1,1)
			'Right click on node to open RMB menu
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
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
					sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0))
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
				Case "1"
					sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0),aPopupMenu(1))					
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
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB on specific multiple nodes
	Case "MultiSelectPopupMenuSelect"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_PSE_BOMTableOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu on node [" & Cstr(sNodeName) & "] as user passed invalid RMB menu [ False ] under BOM table" ,"","","","","")
			Call Fn_ExitTest()
		End If
		aNodeName = split(sNodeName , "^")
		'Clear the already selected Nodes
		objBOMTable.Object.clearSelection
		For iCounter = 0 to UBound(aNodeName)
			If inStr(aNodeName(iCounter),"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,aNodeName(iCounter))
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,aNodeName(iCounter))
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,aNodeName(iCounter))
				End If
			End If
			If iRowCounter <> -1 Then
				objBOMTable.ExtendRow iRowCounter 
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on node [ " & Cstr(aNodeName(iCounter)) & " ] as specified node does not exist under BOM table","","","","","")
				Call Fn_ExitTest()
			End If
		Next
		If inStr(Trim(aNodeName(UBound(aNodeName))),"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,Trim(aNodeName(UBound(aNodeName))))
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,Trim(aNodeName(UBound(aNodeName))))
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,Trim(aNodeName(UBound(aNodeName))))
			End If
		End If
	
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")
			'Split Context menu to Build Path Accordingly
			aPopupMenu = split(sPopupMenu,":",-1,1)
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
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
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify RMB menu available on specific node
	Case "VerifyPopupMenuExists","VerifyPopupMenuNonExists"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_PSE_BOMTableOperations","",sPopupMenu)
		
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If

		If iRowCounter <> -1 Then
			'Split Context menu to Build Path Accordingly
			aPopupMenu = split(sPopupMenu,":",-1,1)
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as unable to right click on correct cell while performing [ " & Cstr(sAction) & " ] opearation","","","","","")					
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
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
					
					If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","LEFT","") = False Then
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
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify RMB menu is active for specific node
	Case "VerifyPopupMenuEnabled"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_PSE_BOMTableOperations","",sPopupMenu)
		
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If

		If iRowCounter <> -1 Then
			'Split Context menu to Build Path Accordingly
			aPopupMenu = split(sPopupMenu,":",-1,1)		
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as unable to right click on correct cell while performing [ " & Cstr(sAction) & " ] opearation","","","","","")					
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
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
					
					If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","LEFT","") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as unable to click on correct cell while performing [ " & Cstr(sAction) & " ] opearation","","","","","") 
						Call Fn_ExitTest()
					End If					
			End Select
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
			Set objContextMenu = nothing
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail while performing [ " & Cstr(sAction) & " ] opearation","","","","","") 
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get child\sub nodes of specific node
	Case "GetChildItems"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		iColumnIndex=Fn_RAC_PSEBOMTableColumnOperations("getcolumnindex",objBOMTable,"BOM Line")
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1        
		DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
			
		If iRowCounter <> -1 Then
			iRowCounter=iRowCounter+1
			sTempValue=objBOMTable.Object.getPathForRow(iRowCounter).getParentPath().toString()
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
		If inStr(aNodeName(0),"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,aNodeName(0))
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,aNodeName(0))
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,aNodeName(0))
			End If
		End If
		iColumnIndex=Fn_RAC_PSEBOMTableColumnOperations("getcolumnindex",objBOMTable,"BOM Line")
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1        
		DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
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
		DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
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
		DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
		
		If sTempValue = ""  Then
			DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		Else
			DataTable.Value("ReusableActionWordReturnValue","Global")= sTempValue
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to validate foregroud\Background color of specific node cell
	Case "VerifyForegroundColour", "VerifyBackgroundColour"
		If sNodeName <> "" Then
			If inStr(sNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
				End If
			End If
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
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified colour [ " & Cstr(sValue) & " ] of operation [ " & Cstr(sAction) & " ]","","","","DONOTSYNC","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify colour [ " & Cstr(sValue) & " ] of operation [ " & Cstr(sAction) & " ]","","","","","") 
				Call Fn_ExitTest()
			End If
			iCounter = iCounter + 1
			'Releasing object of node
			Set objNodeForRow = nothing
		Loop
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to validate foregroud\Background color of specific node cell
	Case "PopupMenuSelectExt"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_PSE_BOMTableOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu on node [" & Cstr(sNodeName) & "] as user passed invalid RMB menu [ False ] under BOM table" ,"","","","","")
			Call Fn_ExitTest()
		End If
		'Clear already selected nodes
		objBOMTable.Object.clearSelection
		
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If

		If iRowCounter <> -1 Then
			'Right click on node to open RMB menu
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","ClickCell",objBOMTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Selecting java menu
			If Fn_UI_JavaMenu_Operations("RAC_PSE_BOMTableOperations","Select",JavaWindow("jwnd_StructureManager"),sPopupMenu)=False Then				
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under BOM table","","","","DONOTSYNC","")
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on node [ " & Cstr(sNodeName) & " ] as specified node does not exist under BOM table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to validate foregroud\Background color of specific node cell			
	Case "PopulateObjectDetailsInDataTable"
		DataTable.SetCurrentRow 1
		Call Fn_CommonUtil_DataTableOperations("AddColumn","BOMTableObjectCount","","")
		iCount=Fn_CommonUtil_DataTableOperations("GetValue","BOMTableObjectCount","","")
		If iCount="" Then
			iCount=1
		Else
			iCount= iCount+1
		End If
		Call Fn_CommonUtil_DataTableOperations("SetValue","BOMTableObjectCount",iCount,"")
		DataTable.SetCurrentRow iCount
		Call Fn_CommonUtil_DataTableOperations("AddColumn","BOMTableObjectRevisionName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","BOMTableObjectRevisionPath","","")
				
		iRowCounter = objBOMTable.GetROProperty("rows")
		
		For iCounter = 0 To iRowCounter -1 Step 1
			If inStr(1,objBOMTable.object.getPathForRow(iCounter).getLastPathComponent().toString(),sValue) Then
				Exit For
			End If
		Next
		
		If cInt(iCounter) = cInt(iRowCounter) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to populate path details of item with id [ " & Cstr(sValue) & " ] in data table from BOM table","","","","","")
			Call Fn_ExitTest()
		Else			
			Datatable.Value("BOMTableObjectRevisionName","Global") = objBOMTable.object.getPathForRow(iCounter).getLastPathComponent().toString()
			sTempValue = Split(objBOMTable.object.getPathForRow(iCounter).toString(),", ",2)
			sTempValue(1) = Replace(sTempValue(1),"]","")
			sTempValue(1) = Replace(sTempValue(1),", ","~")
			Datatable.Value("BOMTableObjectRevisionPath","Global") = sTempValue(1)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check expand node
	Case "ExpandBelowToLevelAndCollapseLowerLevel"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		
		If iRowCounter <> -1 Then			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")	
			If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","SelectRow",objBOMTable,"",iRowCounter,"","","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] to level [ " & Cstr(sValue) & " ] from BOM table","","","","","")
				Call Fn_ExitTest()
			End If		
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewExpandBelowToLevel"
			
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

			'Creating object of [ Expand To Level ] dialog
			Set objExpandBelow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_ExpandToLevel","")
			
			'Checking existance of [ Expand To Level ] dialog
			If Fn_UI_Object_Operations("RAC_PSE_BOMTableOperations","Exist",objExpandBelow,"","","")=False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] to level [ " & Cstr(sValue) & " ] from BOM table as [ Expand To Level ] dialog does not exist","","","","","")
				Call Fn_ExitTest()
			End If
			
			'settin Expand level
			IF Fn_UI_JavaEdit_Operations("RAC_PSE_BOMTableOperations","Set",objExpandBelow,"jedt_ExpandLevel",sValue)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] to level [ " & Cstr(sValue) & " ] from BOM table as failed to set exand level in in field","","","","","")
				Call Fn_ExitTest()
			End IF
			
			'Setting [ Collapse Lower Level ] option
			If Fn_UI_JavaCheckBox_Operations("RAC_PSE_BOMTableOperations", "Set", objExpandBelow, "jckb_CollapseLowerLevel", "ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] to level [ " & Cstr(sValue) & " ] from BOM table as failed to select [ Collapse Lower Level ] option","","","","","")
				Call Fn_ExitTest()
			End If
	
			If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations","Click",objExpandBelow,"jbtn_OK")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] to level [ " & Cstr(sValue) & " ] from BOM table as fail to click on [ Yes ] button of [ Expand Below ] dialog","","","","","")
				Call Fn_ExitTest()
			End If

			'Releasing object of [ Expand To Level ] dialog
			Set objExpandBelow=Nothing			
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & Cstr(sNodeName) & " ] to level [ " & Cstr(sValue) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
		
	Case "GetAllItemIDs"

		'Select the top node
		If Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","SelectRow",objBOMTable,"",0,"","","","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to get all item ids as failed to select top node in BOM table","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Select menu View - Expand Below
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewExpandBelow"
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

		'Creating object of [ Expand Below ] dialog
		Set objExpandBelow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_ExpandBelow","")
		Set objExpandBelow1=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_ExpandBelow@2","")
		
		'Checking existance of [ Expand Below ] dialog
		If Fn_UI_Object_Operations("RAC_PSE_BOMTableOperations","Exist",objExpandBelow,"","","") then
			If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations","Click",objExpandBelow,"jbtn_Yes")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table as fail to click on [ Yes ] button of [ Expand Below ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
		ElseIf Fn_UI_Object_Operations("RAC_PSE_BOMTableOperations","Exist",objExpandBelow1,"","","") then
			If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations","Click",objExpandBelow1,"jbtn_Yes")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table as fail to click on [ Yes ] button of [ Expand Below ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End IF
		
		'Releasing object of [ Expand Below ] dialog
		Set objExpandBelow=Nothing
		Set objExpandBelow1=Nothing
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Get all the item ids
		For iCounter =0  to CInt(objBOMTable.Object.getRowCount) - 1 
			If iCounter = 0 Then
				sTempValue = objBOMTable.GetCellData(iCounter, "Item Id")
			Else
				sTempValue = sTempValue & "~" & objBOMTable.GetCellData(iCounter, "Item Id")
			End If
		Next
		
		'Store details in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1        
		DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= sTempValue
		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellValueNotEditable","VerifyCellValueEditable"
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		If iRowCounter <> -1 Then
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")
			
			'Get the column index
			sTempValue = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_PSE_BOMTableOperations","getcolumnindex",objBOMTable,"","",sColumnName,"","","")))
			
			'Get editable value of the cell
			sTempValue = objBOMTable.Object.isCellEditable(Cint(iRowCounter), Cint(sTempValue))
			
			If sAction = "VerifyCellValueNotEditable" Then
				If Lcase(Cstr(sTempValue)) = "false" Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified column [ " & Cstr(sColumnName) & " ] is non editable for node [ " & Cstr(sNodeName) & " ] under BOM table","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Failed to verify column [ " & Cstr(sColumnName) & " ] is non editable for node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			If sAction = "VerifyCellValueEditable" Then
				If Lcase(Cstr(sTempValue)) = "true" Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified column [ " & Cstr(sColumnName) & " ] is editable for node [ " & Cstr(sNodeName) & " ] under BOM table","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Failed to verify column [ " & Cstr(sColumnName) & " ] is editable for node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of node [ " & Cstr(sNodeName) & " ] under BOM table","","","","","") 
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to verify specific column exist on table or not
	Case "VerifyColumnExist","VerifyColumnNonExist"
		iColumnIndex=Fn_RAC_PSEBOMTableColumnOperations("getcolumnindex",objBOMTable,sColumnName)				
		If iColumnIndex <> -1 Then
			If sAction="VerifyColumnExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified column [ " & Cstr(sColumnName) & " ] is exist\available on BOM table","","","","DONOTSYNC","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as column [ " & Cstr(sColumnName) & " ] is exist\available on BOM table","","","","","") 
				Call Fn_ExitTest()
			End If
		Else
			If sAction="VerifyColumnNonExist" Then				
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified column [ " & Cstr(sColumnName) & " ] does not exist\available on BOM table","","","","DONOTSYNC","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as column [ " & Cstr(sColumnName) & " ] does not exist\available on BOM table","","","","","") 
				Call Fn_ExitTest()
			End If
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Add column to table
	Case "AddColumn"
		objBOMTable.SelectColumnHeader "#1","RIGHT"
		If objPSEApplet.JavaMenu("label:=Insert column\(s\) ...").Exist(6) Then
			objPSEApplet.JavaMenu("label:=Insert column\(s\) ...").Select
		End If
		LoadAndRunAction "RAC_Common\RAC_Common_ChangeColumnsOperations","RAC_Common_ChangeColumnsOperations",OneIteration,"Add","nooption","psebomtable",sColumnName,"Close",""
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Imprecise node from BOM table
	Case "Imprecise","Precise"
		sTempNodeName=sNodeName
'		If sAction="Imprecise" Then
'			sTempValue="Precise"
'		ElseIf sAction="Precise" Then
'			sTempValue="Working"
'		End If
		
		If inStr(sNodeName,"@") > 1 Then
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
		Else
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sNodeName)
			If iRowCounter = -1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sNodeName)
			End If
		End If
		If iRowCounter = -1 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to " & sAction & " node [ " & Cstr(sNodeName) & " ] from BOM table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		bFlag = False
		Select Case sAction
		
			Case "Imprecise"
				If Instr(1,objBOMTable.Object.getComponentForRow(iRowCounter).getProperty("bl_config_string"),"Precise") Then
					bFlag = True
				Else
					bFlag = False
				End If
			
			Case "Precise"
			
				If Instr(1,objBOMTable.Object.getComponentForRow(iRowCounter).getProperty("bl_config_string"),"Precise") Then
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
			
			If inStr(sTempNodeName,"@") > 1 Then
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sTempNodeName)
			Else
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sTempNodeName)
				If iRowCounter = -1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sTempNodeName)
				End If
			End If
			
			If iRowCounter = -1 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to " & sAction & " node [ " & Cstr(sTempNodeName) & " ] from BOM table as specified node does not exist","","","","","")
				Call Fn_ExitTest()
			End If
			'Selecting node from table
			objBOMTable.Object.clearSelection  
			objBOMTable.SelectRow iRowCounter
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","TogglePreciseImprecise"
			'Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations", oneIteration, "Select", "FileSave"
			'Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to " & sAction & " node [ " & CStr(sNodeName) & " ] from BOM table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_BOMTableOperations",sAction,"","")			
			sObjectTypeName=Fn_RAC_GetTreeNodeType("PSEBomTable","getitemtypename")			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully " & sAction & " [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from BOM table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
		End If
		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Update mass properties
	Case "UpdateMassProperties"
		
		'Call menu operation to update mass properties
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"KeyPress","CATIAV5UpdateAssemblyMassProperties"
'		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

		'Create object of information window
		Set objInformation = Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jwnd_Information","")
		
		'Verify existence of Information dialog
		bFlag = False
		For iCounter = 1 To 30 Step 1
			If Fn_UI_Object_Operations("RAC_PSE_BOMTableOperations","Exist",objInformation,"2","","")= True Then
				bFlag = True
				Exit For
			Else
				Wait GBL_MIN_MICRO_TIMEOUT
			End If
		Next
		
		'Click on OK button of Information dialog
		If bFlag = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on BOM table as Information dialog was not displayed","","","","","")
			Call Fn_ExitTest()
		Else
			If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations","Click",objInformation,"jbtn_OK") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [" & sAction & "] as failed to click on OK button of Information dialog","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully performed Update Mass Properties operation","","","","","")
			End If
		End If
		Set objInformation = Nothing		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to view sorted order of cloumns
	Case "SetSortingOrderAndVerifySortOrder"
	
		LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "AddColumn","",sColumnName,"",""
		If sColumnName<>"" Then
			iColumnIndex=Fn_RAC_PSEBOMTableColumnOperations("getcolumnindex",objBOMTable,sColumnName)
		Else
			iColumnIndex=Fn_RAC_PSEBOMTableColumnOperations("getcolumnindex",objBOMTable,"BOM Line")
			sColumnName="BOM Line"
		End If
		If sValue="" or sValue="Ascending" Then
			objBOMTable.Object.getColumnModel().setSortOrder iColumnIndex,true
			sValue="Ascending"
		Else
			objBOMTable.Object.getColumnModel().setSortOrder iColumnIndex,false
			sValue="Descending"
		End If
		
		'Get child items of top node
		LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "GetChildItems",sNodeName,"","",""
		Datatable.SetCurrentRow 1
		aNodeName = Split(Datatable.Value("ReusableActionWordReturnValue","Global"), "^")		
				
		If sColumnName<>"BOM Line" Then
			For iCounter = 0 To Ubound(aNodeName)
				sTempNodeName=sNodeName & "~" & aNodeName(iCounter)
				If inStr(sTempNodeName,"@") > 1 Then
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sTempNodeName)
				Else
					iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindex",objBOMTable,sTempNodeName)
					If iRowCounter = -1 Then
						iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getnodeindexext",objBOMTable,sTempNodeName)
					End If
				End If
				If iRowCounter = -1 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & Cstr(sTempNodeName) & " ] from BOM table as specified node does not exist","","","","","")
					Call Fn_ExitTest()
				End If
				If sTempValue="" Then
					sTempValue=objBOMTable.GetCellData(iRowCounter, sColumnName)
				Else
					sTempValue=sTempValue & "~" & objBOMTable.GetCellData(iRowCounter, sColumnName)				
				End If				
			Next
			aNodeName=Split(sTempValue,"~")
		Else
			'Get Item IDs of the child items
			For iCounter = 0 To Ubound(aNodeName)
				aNodeName(iCounter) = Split(aNodeName(iCounter),"/")(0)
			Next		
		End If
				
		If Fn_CommonUtil_StringArrayOperations("VerifyOrder",aNodeName,sValue) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as order of item id in BOM Table [ " & Join(aNodeName,"~") & " ] is not sorted in [ " & Cstr(sValue) & " ] order","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified order of item id in BOM Table [ " & Join(aNodeName,"~") & " ] is sorted in [ " & Cstr(sValue) & " ] order","","","","DONOTSYNC","")
		End If
	Case "VerifyNodeIsSelected"
		'Creating object of node
		sTempValue = ""
		sTempNodeName =""
		For iCounter =0  to CInt(objBOMTable.Object.getRowCount) - 1 
			If objBOMTable.Object.isRowSelected(iCounter) Then
				sTempNodeName =objBOMTable.Object.getPathForRow(iCounter).toString()
				sTempNodeName = Right(sTempNodeName, (Len(sTempNodeName)-inStr(1, sTempNodeName, ",", 1)))					
				sTempNodeName = Left(sTempNodeName, Len(sTempNodeName)-1)
				
				If inStr(sTempNodeName, "@BOM::") > 0 Then
					sTempNodeName = Trim(replace(sTempNodeName,"""",""))
					aNodeName = split(sTempNodeName,",")
					sTempNodeName = ""
					For iCount = 0 to uBound(aNodeName)
						aNodeName(iCount) = Left(aNodeName(iCount), inStr(aNodeName(iCount),"@")-1)
						If sTempNodeName = "" Then
							sTempNodeName = Trim(aNodeName(iCount))
						Else
							sTempNodeName = sTempNodeName & ", " & Trim(aNodeName(iCount))
						End If
					Next
				End If
				sTempNodeName = Trim(replace(sTempNodeName,", ","~"))
				If sTempValue = "" Then
					sTempValue = sTempNodeName
				Else
					sTempValue = sTempValue & "~" & sTempNodeName
				End If
			End if
		Next
	
		If sTempValue = sNodeName  Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified node [ " & Cstr(sNodeName) & " ] is selected under BOM table","","","","DONOTSYNC","") 
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] is not selected under BOM table","","","","","") 
			Call Fn_ExitTest()
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to create new snapshot
	Case "CreateSnapshot"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileNewSnapshot"
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

		Set objSnapShot = Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_CreateSnapshot","")
		If Fn_UI_Object_Operations("RAC_PSE_BOMTableOperations","Exist",objSnapShot,"","","")= False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on BOM table as fail to verify existence of snapshot window","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Enter name
		sTempValue = Fn_Setup_GenerateObjectInformation("getname","Snapshot")
		IF Fn_UI_JavaEdit_Operations("RAC_PSE_BOMTableOperations","Set",objSnapShot.JavaEdit("jedt_Name"),"", sTempValue)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter value in name field of snapshot dialog","","","","","")
			Call Fn_ExitTest()
		End IF		
		
		'Click on ok button
		If Fn_UI_JavaButton_Operations("RAC_PSE_BOMTableOperations","Click",objSnapShot,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create snapshot as failed to click on OK button.","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Store the name in data table
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1        
		DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_BOMTableOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= sTempValue
		Set objSnapShot = Nothing
		
	Case "SaveWithConfirmation"	
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations", oneIteration, "Select", "FileSave"
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_BOMTableOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		If JavaWindow("jwnd_StructureManager").Dialog("dlg_ConfirmationDialog").Exist(20) Then
			JavaWindow("jwnd_StructureManager").Dialog("dlg_ConfirmationDialog").WinButton("wbtn_Yes").Click
			If Err.Number <> 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save assembly from BOM table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully saved assembly from BOM table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save assembly from BOM table as [ Confirmation Dialog ] does not appears after clicking on Save button","","","","","") 
			Call Fn_ExitTest()
		End If	
	Case "VerifyPanelHeader"
		If sValue="" Then
			sValue="INCORRECT INPUT"
		End If
		
		If Instr(1,objPSEApplet.JavaStaticText("jstx_BOMTablePanelHeader").GetROProperty("label"),sValue) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified BOM table panel header contains [ " & Cstr(sValue) & " ] value","","","","DONOTSYNC","") 
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as BOM table panel header does not contains [ " & Cstr(sValue) & " ] value","","","","","") 
			Call Fn_ExitTest()	
		End If
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

If Err.number <> 0  Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on BOM table due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing Objects
Set objBOMTable =Nothing
Set objPSEApplet = Nothing

Function Fn_ExitTest()
	'Releasing Objects
	Set objBOMTable =Nothing
	Set objPSEApplet = Nothing
	ExitTest
End Function

