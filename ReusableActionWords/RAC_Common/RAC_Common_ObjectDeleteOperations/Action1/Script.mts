'! @Name 			RAC_Common_ObjectDeleteOperations
'! @Details 		Action word to Delete object.
'! @InputParam1 	sAction 			: Action to be performed e.g. Delete
'! @InputParam2 	sInvokeOption 		: Method to invoke Remove dialog e.g. menu
'! @InputParam3 	sNodePath 			: Node path
'! @InputParam4 	sNodeContainer 		: Node container name
'! @InputParam5 	sDeleteAllSequences : Keep sub tree option
'! @InputParam6 	sButton		 		: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			17 March 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ObjectDeleteOperations","RAC_Common_ObjectDeleteOperations",OneIteration,"basicdeletewitherror","toolbar","0000313/AA-Asm Cockpit~0000316/AA-Asm Cockpit","navigationtree","ON",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ObjectDeleteOperations","RAC_Common_ObjectDeleteOperations",OneIteration,"basicdelete","toolbar","0000313/AA-Asm Cockpit~0000316/AA-Asm Cockpit","navigationtree","ON",""

Option Explicit
Err.Clear

'Declaring varaibles	
Dim sAction,sInvokeOption,sNodePath,sNodeContainer,sDeleteAllSequences,sButton
Dim objDelete

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sNodePath = Parameter("sNodePath")
sNodeContainer = Parameter("sNodeContainer")
sDeleteAllSequences = Parameter("sDeleteAllSequences")
sButton= Parameter("sButton")

'Selecting node from table
If sNodePath<>"" Then
	Select Case LCase(sNodeContainer)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "navigationtree",""
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Select", sNodePath,""
	End Select
End If

'inoke Delete dialog
Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditDelete"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "toolbar"	
		LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","Delete", "",""
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'Use this invoke option when user wants to invoke Delete dialog from outside function
End Select

'Creating object of [ Delete ] dialog
Set objDelete=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_Delete","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ObjectDeleteOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of Delete dialog
If Fn_UI_Object_Operations("RAC_Common_ObjectDeleteOperations", "Exist", objDelete, GBL_DEFAULT_TIMEOUT,"","")=True Then
	sButton="jbtn_OK"
ElseIf Fn_UI_Object_Operations("RAC_Common_ObjectDeleteOperations", "Exist", JavaDialog("jdlg_Delete"), GBL_DEFAULT_TIMEOUT,"","")=True Then
	Set objDelete=JavaDialog("jdlg_Delete")
	sButton="jbtn_Yes"
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Delete ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If


'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Object Delete Operations",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Delete obejct
	Case "basicdelete"
		If sDeleteAllSequences<>"" Then
			objDelete.JavaCheckBox("jckb_DeleteAllSequences").Set sDeleteAllSequences
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		
		'Click on [ OK ] button
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectDeleteOperations", "Click", objDelete,sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Delete selected object as fail to click on [ " & Cstr(sButton) & " ] button","","","","","")	
			Call Fn_ExitTest()
		End If			
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Delete Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Delete selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Deleted selected object from assembly","","","","","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Delete obejct
	Case "basicdeletewitherror"
		If sDeleteAllSequences<>"" Then
			objDelete.JavaCheckBox("jckb_DeleteAllSequences").Set sDeleteAllSequences
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		
		'Click on [ OK ] button
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectDeleteOperations", "Click", objDelete,sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Delete selected object as fail to click on [ " & Cstr(sButton) & " ] button","","","","","")	
			Call Fn_ExitTest()
		End If			
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Delete Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of Delete dialog for selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on [ " & Cstr(sButton) & " ] button of Delete dialog for selected object","","","","","")
		End If
End Select

'Releasing object
Set objDelete=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objDelete=Nothing
	ExitTest
End Function

