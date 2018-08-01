'! @Name 			RAC_Common_ColumnManagementOperations
'! @Details 		This Action word to perform operations on Column Management dialog
'! @InputParam1 	sAction 			: Action to be performed
'! @InputParam2 	sInvokeOption 		: Invoke Option (menu or nooption)
'! @InputParam3 	sColumnName 		: Column name
'! @InputParam4 	sButton 			: Button Name
'! @InputParam5 	bShowInternalNames 	: Show Internal Names options
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com 
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			25 Apr 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ColumnManagementOperations","RAC_Common_ColumnManagementOperations",OneIteration, "Add","detailstableviewmenu","Item Type","Close",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption, sColumnName,sButton,bShowInternalNames
Dim sPerspective
Dim objColumnManagement
Dim aColumnName
Dim iCounter,iCount
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters value in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sColumnName = Parameter("sColumnName")
sButton = Parameter("sButton")
bShowInternalNames = Parameter("bShowInternalNames")

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Get [ Column Management ] object from xml file
Select Case lcase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter","my teamcenter"
		Set objColumnManagement=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_ColumnManagement","")
End Select

Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "detailstableviewmenu"
		LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"PopupMenuSelect", "ViewMenu", "Column...","2"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ColumnManagementOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'checking Change Ownership dialogs existence
If Not Fn_UI_Object_Operations("RAC_Common_ColumnManagementOperations","Exist", objColumnManagement, GBL_DEFAULT_TIMEOUT,"","") Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on Column Management dialog as [ Column Management ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Column Management",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "add"
		aColumnName=Split(sColumnName,"~")
		For iCounter=0 to ubound(aColumnName)
			bFlag=False
			For iCount=0 to objColumnManagement.JavaTable("jtbl_DisplayedColumns").GetROProperty("rows")-1
				If Trim(objColumnManagement.JavaTable("jtbl_DisplayedColumns").GetCellData(iCount,"Property"))= Trim(aColumnName(iCounter)) Then
					bFlag=True
					Exit For	
				End If
			Next
			If bFlag=True Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added column [ " & Cstr(aColumnName(iCounter)) & " ] from Column Management dialogs [ Available Columns ] list to [ Displayed Columns ] list","","","","","")
			Else
				bFlag=False
				For iCount=0 to objColumnManagement.JavaTable("jtbl_AvailableProperties").GetROProperty("rows")-1
					If Trim(objColumnManagement.JavaTable("jtbl_AvailableProperties").GetCellData(iCount,"Property"))= Trim(aColumnName(iCounter)) Then
						objColumnManagement.JavaTable("jtbl_AvailableProperties").SelectCell iCount,0
						bFlag=True
						Exit For	
					End If
				Next
				If bFlag=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to add column [ " & Cstr(aColumnName(iCounter)) & " ] from Column Management dialog as specific column is not available in [ Available Columns ] list","","","","","")
					Call Fn_ExitTest()
				End If	
				'Click on add button
				If Fn_UI_JavaButton_Operations("RAC_Common_ColumnManagementOperations", "Click", objColumnManagement,"jbtn_Add") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail add column [ " & Cstr(aColumnName(iCounter)) & " ] from Column Management dialog as fail to click on [ Add ] button","","","","","")
					Call Fn_ExitTest()
				End If			
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added column [ " & Cstr(aColumnName(iCounter)) & " ] from Column Management dialogs [ Available Columns ] list to [ Displayed Columns ] list","","","","","")	
			End IF
		Next
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on Apply button
		If Fn_UI_JavaButton_Operations("RAC_Common_ColumnManagementOperations", "Click", objColumnManagement,"jbtn_Apply") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail add columns from Column Management dialog as fail to click on [ Apply ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		If sButton<>"" Then
			'Click on button
			If Fn_UI_JavaButton_Operations("RAC_Common_ColumnManagementOperations", "Click", objColumnManagement,"jbtn_" & sButton) = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail add columns from Column Management dialog as fail to click on [ " & Cstr(sButton) & " ] button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Column Management",sAction,"","")
End Select

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on Column Management dialog due to error number [" & Cstr(Err.Number) & "] and error description  [" & Cstr(Err.Description) & "]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing object of Column Management dialog
Set objColumnManagement = Nothing
	
Function Fn_ExitTest()
	'Releasing object of Column Management dialog
	Set objColumnManagement = Nothing
	ExitTest
End Function
