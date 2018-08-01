'! @Name 			RAC_Common_ChangeColumnsOperations
'! @Details 		This Action word to perform operations on Change Columns dialog
'! @InputParam1 	sAction 			: Action to be performed
'! @InputParam2 	sInvokeOption 		: Invoke Option (menu or nooption)
'! @InputParam3 	sTableName 			: Table name on which user wants to perform operation
'! @InputParam4 	sColumnName 		: Column name
'! @InputParam5 	sButton 			: Button Name
'! @InputParam6 	bShowInternalNames 	: Show Internal Names options
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com 
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			25 Apr 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ChangeColumnsOperations","RAC_Common_ChangeColumnsOperations",OneIteration, "Add","nooption", "psebomtable", "Item Type","Close",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ChangeColumnsOperations","RAC_Common_ChangeColumnsOperations",OneIteration, "Remove","nooption", "psebomtable", "BOM Line","Close",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption, sTableName,sColumnName,sButton,bShowInternalNames
Dim sPerspective
Dim objChangeColumns
Dim aColumnName
Dim iCounter

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters value in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sTableName =  Parameter("sTableName")
sColumnName = Parameter("sColumnName")
sButton = Parameter("sButton")
bShowInternalNames = Parameter("bShowInternalNames")

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Get [ Change Columns ] object from xml file
Select Case lcase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter","my teamcenter"
		Set objChangeColumns=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ChangeColumns","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "structuremanager","structure manager"
		Set objChangeColumns=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ChangeColumns@2","")
End Select

Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ChangeColumnsOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'checking Change Ownership dialogs existence
If Not Fn_UI_Object_Operations("RAC_Common_ChangeColumnsOperations","Exist", objChangeColumns, GBL_DEFAULT_TIMEOUT,"","") Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on Change Columns dialog as [ Change Columns ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Change Columns",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "add"
		aColumnName=Split(sColumnName,"~")
		For iCounter=0 to ubound(aColumnName)
			If Fn_UI_JavaList_Operations("RAC_Common_ChangeColumnOperations", "Select", objChangeColumns, "jlst_DisplayedColumns", aColumnName(iCounter), "","")=False Then
				If Fn_UI_JavaList_Operations("RAC_Common_ChangeColumnOperations", "Select", objChangeColumns, "jlst_AvailableColumns", aColumnName(iCounter), "","")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to add column [ " & Cstr(aColumnName(iCounter)) & " ] from Change Columns dialog as specific column is not available in [ Available Columns ] list","","","","","")
					Call Fn_ExitTest()
				Else
					'Click on add button
					If Fn_UI_JavaButton_Operations("RAC_Common_ChangeColumnsOperations", "Click", objChangeColumns,"jbtn_Add") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail add column [ " & Cstr(aColumnName(iCounter)) & " ] from Change Columns dialog as fail to click on [ Add ] button","","","","","")
						Call Fn_ExitTest()
					End If			
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added column [ " & Cstr(aColumnName(iCounter)) & " ] from Change Columns dialogs [ Available Columns ] list to [ Displayed Columns ] list","","","","","")
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added column [ " & Cstr(aColumnName(iCounter)) & " ] from Change Columns dialogs [ Available Columns ] list to [ Displayed Columns ] list","","","","","")
			End If
		Next
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on Apply button
		If Fn_UI_JavaButton_Operations("RAC_Common_ChangeColumnsOperations", "Click", objChangeColumns,"jbtn_Apply") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail add columns from Change Columns dialog as fail to click on [ Apply ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		If sButton<>"" Then
			'Click on button
			If Fn_UI_JavaButton_Operations("RAC_Common_ChangeColumnsOperations", "Click", objChangeColumns,"jbtn_" & sButton) = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail add columns from Change Columns dialog as fail to click on [ " & Cstr(sButton) & " ] button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Change Columns",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "remove"
		aColumnName=Split(sColumnName,"~")
		For iCounter=0 to ubound(aColumnName)
			If Fn_UI_JavaList_Operations("RAC_Common_ChangeColumnOperations", "Select", objChangeColumns, "jlst_DisplayedColumns", aColumnName(iCounter), "","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remove column [ " & Cstr(aColumnName(iCounter)) & " ] from Change Columns dialog as specific column is not available in [ Displayed Columns ] list","","","","","")
				Call Fn_ExitTest()
			Else
				'Click on Remove button
				If Fn_UI_JavaButton_Operations("RAC_Common_ChangeColumnsOperations", "Click", objChangeColumns,"jbtn_Remove") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail remove column [ " & Cstr(aColumnName(iCounter)) & " ] from Change Columns dialog as fail to click on [ Remove ] button","","","","","")
					Call Fn_ExitTest()
				End If			
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully removed column [ " & Cstr(aColumnName(iCounter)) & " ] from Change Columns dialogs [ Displayed Columns ] list","","","","","")
			End If
		Next
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on Apply button
		If Fn_UI_JavaButton_Operations("RAC_Common_ChangeColumnsOperations", "Click", objChangeColumns,"jbtn_Apply") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail add columns from Change Columns dialog as fail to click on [ Apply ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		If sButton<>"" Then
			'Click on button
			If Fn_UI_JavaButton_Operations("RAC_Common_ChangeColumnsOperations", "Click", objChangeColumns,"jbtn_" & sButton) = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail add columns from Change Columns dialog as fail to click on [ " & Cstr(sButton) & " ] button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If				
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Change Columns",sAction,"","")	
End Select

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on Change Columns dialog due to error number [" & Cstr(Err.Number) & "] and error description  [" & Cstr(Err.Description) & "]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing object of Change Columns dialog
Set objChangeColumns = Nothing
	
Function Fn_ExitTest()
	'Releasing object of Change Columns dialog
	Set objChangeColumns = Nothing
	ExitTest
End Function
