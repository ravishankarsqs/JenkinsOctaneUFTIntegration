'! @Name 			RAC_PSE_AttachmentTableOperations
'! @Details 		Action word to perform operations on structure manager attachment table
'! @InputParam1 	sAction 			: String to indicate what action is to be performed on structure manager attachment table e.g. Select, Expand
'! @InputParam2 	sNodeName 			: Node name in structure manager attachment table on which action is to be performed
'! @InputParam3 	sColumnName			: attachment table column name
'! @InputParam4 	sValue	 			: Value to pass
'! @InputParam5 	sPopupMenu 			: Menu tag name from XML
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			07 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_AttachmentTableOperations","RAC_PSE_AttachmentTableOperations", oneIteration, "VerifyExist","WS_P050050/01;1-AUTWSP_P-Part55511","","",""
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_AttachmentTableOperations","RAC_PSE_AttachmentTableOperations", oneIteration, "Select",DataTable.Value("BOMTableObjectRevisionPath"),"","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sNodeName,sColumnName,sValue,sPopupMenu
Dim iRowCounter
Dim aNodeName
Dim sDatasetType,sChildNodePath,sNodePath,sTempValue,sObjectTypeName
Dim iInstanceHandler,iCount
Dim bFlag
Dim aPopupMenu
Dim objAttachmentsTable,objPSEApplet,objContextMenu

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sColumnName = Parameter("sColumnName")
sValue = Parameter("sValue")
sPopupMenu = Parameter("sPopupMenu")

'Creating obejct of [ attachment table ]
Set objAttachmentsTable=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jtbl_AttachmentsTable","")
Set objPSEApplet=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","japt_PSEApplet","")

LoadAndRunAction "RAC_StructureManager\RAC_PSE_DataPanelTabOperations","RAC_PSE_DataPanelTabOperations",OneIteration,"Activate","Attachments"

objPSEApplet.JavaObject("jobj_AttachmentsPanel").Click 0,1
Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_AttachmentTableOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME=sAction

'Checking existance of attachment table
If Fn_UI_Object_Operations("RAC_PSE_AttachmentTableOperations","Exist",objAttachmentsTable,"","","")= False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on node [ " & Cstr(sNodeName) & " ] as attachment table does not exist","","","","","")
	Call Fn_ExitTest()
End if

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_PSE_AttachmentTableOperations",sAction,"","")

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to select node from attachment table
	Case "Select"
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)
		If iRowCounter = -1 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & Cstr(sNodeName) & " ] from attachment table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		'Selecting node from table
		objAttachmentsTable.SelectRow Cint(iRowCounter)
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select node [ " & CStr(sNodeName) & " ] from attachment table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")			
			sObjectTypeName="node"			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from attachment table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to Deselect node from attachment table
	Case "Deselect"
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)
		If iRowCounter = -1 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Deselect node [ " & Cstr(sNodeName) & " ] from attachment table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		'Selecting node from table
		objAttachmentsTable.DeselectRow Cint(iRowCounter)
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Deselect node [ " & CStr(sNodeName) & " ] from attachment table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")			
			sObjectTypeName="node"			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Deselected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from attachment table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to multi select nodes from attachment table
	Case "MultiSelect"
		aNodeName = Split(sNodeName,"^")
		'Clear the already selected Nodes
		objAttachmentsTable.Object.clearSelection
		For iCounter = 0 to UBound(aNodeName)
			iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)
			If iRowCounter <> -1 Then
				objAttachmentsTable.Object.expandNode objAttachmentsTable.Object.getNodeForRow(cint(iRowCounter))
				sObjectTypeName="node"
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully multi selected [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(aNodeName(iCounter)) & " ] from attachment table","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to multiselect node [ " & Cstr(aNodeName(iCounter)) & " ] from attachment table as specified node does not exist","","","","","")
				Call Fn_ExitTest()
			End If
		Next
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to multi select nodes from attachment table
	Case "SelectAll"
		For iCounter = 0 to cInt(objAttachmentsTable.GetROProperty ("rows")) - 1
            objAttachmentsTable.Object.expandNode objAttachmentsTable.Object.getNodeForRow(cint(iRowCounter))
		Next
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select all rows\nodes from attachment table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Selected all rows\nodes of attachment table","","","","","")			
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check whether specific node exist in table or not
	Case "Exist"
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)
		If iRowCounter <> -1 Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_AttachmentTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
		Else
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1        
			DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_AttachmentTableOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to verify specific node exist in table or not
	Case "VerifyExist","VerifyNonExist"
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)			
		If iRowCounter <> -1 Then
			If sAction="VerifyExist" Then
				sObjectTypeName="node"
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] is exist under attachment table","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodeName) & " ] is exist under attachment table","","","","","") 
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
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist under attachment table","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] does not exist under attachment table","","","","","") 
				Call Fn_ExitTest()
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to verify specific node exist in table or not
	Case "VerifyDatasetOfSpecificTypeExist"
		aNodeName=Split(sNodeName,"^")
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,aNodeName(0))
		
		If iRowCounter <> -1 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(aNodeName(0)) & " ] does not exist under attachment table","","","","","")
			Call Fn_ExitTest()
		Else
			sDatasetType = sPopupMenu
			If sPopupMenu <> "" Then
				sPopupMenu = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewDatasetValues_APL",sPopupMenu,"")
			End If
			
			iInstanceHandler=1		
			iCount=Cint(objAttachmentsTable.GetROProperty("rows"))-1
			sChildNodePath=aNodeName(0) & "~" & aNodeName(1)
			bFlag=False
			For iCounter=iRowCounter to iCount-1
				sNodePath = sChildNodePath & "@" & iInstanceHandler
				iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodePath)
				If iRowCounter <> -1 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as node [ " & Cstr(sNodePath) & " ] does not exist under attachment table","","","","","")
					Call Fn_ExitTest()
				End If
				If objAttachmentsTable.Object.getNodeForRow(cint(iRowCounter)).getProperty("me_cl_object_type")=sPopupMenu Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
					bFlag=True
					Exit For
				End If
				iInstanceHandler=iInstanceHandler+1
			Next
		End If
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify that dataset [ " & Cstr(aNodeName(1)) & " ] of type [ " & Cstr(sDatasetType) & " ] does not exist under node [ " & Cstr(aNodeName(0)) & " ]","","","","DONOTSYNC","")
			Call Fn_ExitTest()
		End IF
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to check expand node
	Case "Expand"
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)		
		If iRowCounter <> -1 Then			
			sObjectTypeName="node"	
			objAttachmentsTable.Object.expandNode objAttachmentsTable.Object.getNodeForRow(cint(iRowCounter))
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & Cstr(sNodeName) & " ] from attachment table as specified node does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand node [ " & CStr(sNodeName) & " ] from attachment table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")			
			sObjectTypeName="node"			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully expanded [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] from attachment table","","","",GBL_MICRO_SYNC_ITERATIONS,"")			
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellValue"
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)	
		If iRowCounter <> -1 Then
			sObjectTypeName="node"
			sTempValue = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_PSE_AttachmentTableOperations","GetCellData",objAttachmentsTable,"",iRowCounter,sColumnName,"","","")))
			If sTempValue = Trim(cstr(sValue)) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(Parameter("sColumnName")) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
			Else
				If isNumeric(sTempValue) Then
					If cstr(Abs(sTempValue)) = Trim(cstr(sValue)) Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(Parameter("sColumnName")) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(Parameter("sColumnName")) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
						Call Fn_ExitTest()
					End  If
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(Parameter("sColumnName")) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","") 
					Call Fn_ExitTest()
				End If
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(Parameter("sColumnName")) & " ] does not contain value [ " & Cstr(sValue) & " ] of node [ " & Cstr(sNodeName) & " ] under attachment table","","","","","") 
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify value of specific cell
	Case "VerifyCellValueInStr"
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)
		If iRowCounter <> -1 Then
			'objAttachmentsTable.SelectRow iRowCounter 
			sTempValue = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_PSE_AttachmentTableOperations","GetCellData",objAttachmentsTable,"",iRowCounter,sColumnName,"","","")))
			If inStr(1,sTempValue,Trim(cstr(sValue))) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified field [ " & Cstr(sColumnName) & " ] contains value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","") 
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] of node [ " & Cstr(sNodeName) & " ] under attachment table","","","","","") 
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to get data of specific cell
	Case "GetCellData"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1
		DataTable.Value("ReusableActionWordName","Global")= "RAC_PSE_AttachmentTableOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)
		
		If iRowCounter <> -1 Then
			DataTable.Value("ReusableActionWordReturnValue","Global") = Trim(cstr(Fn_UI_JavaTable_Operations("RAC_PSE_AttachmentTableOperations","GetCellData",objAttachmentsTable,"",iRowCounter,sColumnName,"","","")))
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB on specific node
	Case "PopupMenuSelect"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_PSE_BOMTableOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu on node [" & Cstr(sNodeName) & "] as user passed invalid RMB menu [ False ] under attachment table" ,"","","","","")
			Call Fn_ExitTest()
		End If
		
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)

		If iRowCounter <> -1 Then
			sObjectTypeName="node"
			'Split Context menu to Build Path
			aPopupMenu = Split(sPopupMenu,":",-1,1)
			'Right click on node to open RMB menu
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_PSE_AttachmentTableOperations","ClickCell",objAttachmentsTable,"",Cint(iRowCounter),"Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_PSE_AttachmentTableOperations","ClickCell",objAttachmentsTable,"",Cint(iRowCounter),sColumnName,"","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
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
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
					Call Fn_ExitTest()
			End Select
			'Select RMB menu
			objContextMenu.Select sPopupMenu
			'Creating object of [ Context menu ]
			Set objContextMenu=Nothing
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on node [ " & Cstr(sNodeName) & " ] as specified node does not exist under attachment table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to validate foregroud\Background color of specific node cell
	Case "PopupMenuSelectExt"
		'Retrive popup menu
		sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_PSE_BOMTableOperations","",sPopupMenu)
		
		If sPopupMenu = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu on node [" & Cstr(sNodeName) & "] as user passed invalid RMB menu [ False ] under attachment table" ,"","","","","")
			Call Fn_ExitTest()
		End If
		'Clear already selected nodes
		iRowCounter = Fn_RAC_PSEBOMTableRowOperations("getattachmentstablenodeindex",objAttachmentsTable,sNodeName)

		If iRowCounter <> -1 Then
			'Right click on node to open RMB menu
			If sColumnName = "" Then
				If Fn_UI_JavaTable_Operations("RAC_PSE_AttachmentTableOperations","ClickCell",objAttachmentsTable,"",iRowCounter,"BOM Line","","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				If Fn_UI_JavaTable_Operations("RAC_PSE_AttachmentTableOperations","ClickCell",objAttachmentsTable,"",iRowCounter,sColumnName,"","RIGHT","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
					Call Fn_ExitTest()
				End If
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Selecting java menu
			If Fn_UI_JavaMenu_Operations("RAC_PSE_AttachmentTableOperations","Select",JavaWindow("jwnd_StructureManager"),sPopupMenu)=False Then				
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected popup menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sObjectTypeName) & " ] [ " & Cstr(sNodeName) & " ] under attachment table","","","","","")
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenu) & " ] on node [ " & Cstr(sNodeName) & " ] as specified node does not exist under attachment table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_PSE_AttachmentTableOperations",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Case to verify specific column exist on table or not
	Case "VerifyColumnExist","VerifyColumnNonExist"
		iColumnIndex=Fn_RAC_PSEBOMTableColumnOperations("getcolumnindex",objAttachmentsTable,sColumnName)				
		If iColumnIndex <> -1 Then
			If sAction="VerifyColumnExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified column [ " & Cstr(sColumnName) & " ] is exist\available on attachment table","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as column [ " & Cstr(sColumnName) & " ] is exist\available on attachment table","","","","","") 
				Call Fn_ExitTest()
			End If
		Else
			If sAction="VerifyColumnNonExist" Then				
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified column [ " & Cstr(sColumnName) & " ] does not exist\available on attachment table","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as column [ " & Cstr(sColumnName) & " ] does not exist\available on attachment table","","","","","") 
				Call Fn_ExitTest()
			End If
		End If	
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

If Err.number <> 0  Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on attachment table due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing Objects
Set objAttachmentsTable =Nothing
Set objPSEApplet = Nothing

Function Fn_ExitTest()
	'Releasing Objects
	Set objAttachmentsTable =Nothing
	Set objPSEApplet = Nothing
	ExitTest
End Function


