'! @Name 			RAC_Organization_OrganizationTreeOperations
'! @Details 		To perform operations on Organnization's JV tree
'! @InputParam1 	sAction 		: String to indicate what action is to be performed on JV tree e.g. Select, Expand
'! @InputParam2 	sNodeName 		: Node name in JV tree on which action is to be performed
'! @InputParam3 	sPopupMenu 		: Menu tag name from XML
'! @Author 			Mohinni Deshmukh mohini.deshmukh@sqs.com
'! @Reviewer 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Date 			19 Dec 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Organization\RAC_Organization_OrganizationTreeOperations", "RAC_Organization_JV_Operations", oneIteration, "ExpandAndVerifyExist", "Organization~CFAA",""

Dim sAction,sNodeName,sPopupMenu
Dim objOrganizationTree
Dim aNodeName
Dim iCounter
Dim bFlag

'Get parameter values in local variables
sAction = Parameter("sAction")
sNodeName = Parameter("sNodeName")
sPopupMenu = Parameter("sPopupMenu")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Organization_OrganizationTreeOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'creating object of [ Navigation Tree ]
Set objOrganizationTree=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Organization_OR","jtree_OrganizationTree","")

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify JV from organization tree
	Case "ExpandAndVerifyExist"
		aNodeName=Split(sNodeName,"~")
		sTempNode=aNodeName(0)
		For iCounter = 1 To Ubound(aNodeName)-1
			sTempNode=sTempNode & "~" & aNodeName(iCounter)
			objOrganizationTree.Expand sTempNode
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			If Err.Number <> 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existance of node [ " & Cstr(sNodeName) & "] as failed to expand parent node [ " & Cstr(sTempNode) & "] from organization tree due to error number [ " & Cstr(Err.Number) & " ] and error description as [ " & Cstr(Err.Description) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		Next
		
		bFlag=False
		For iCounter = 1 To objOrganizationTree.GetROProperty("items count")-1
			If objOrganizationTree.GetItem(iCounter)=sNodeName Then
				bFlag=True
				Exit For
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify organization node [" & Cstr(sNodeName) & "] in organization tree as specified node does not exist in organization tree.","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sNodeName) & " ]  exist in organization tree","","","","DONOTSYNC","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	Case Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Invalid operation [ " & Cstr(sAction) & " ]","","","","","")	
End Select

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] operation on organization tree due to error number as [ " & Cstr(Err.Number) & " ] and error description as [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If



Function Fn_ExitTest()
	'Releasing all objects
	Set objOrganizationTree=Nothing
	ExitTest
End Function
