'! @Name 			RAC_PSE_ObjectRemoveOperations
'! @Details 		Action word to remove object from BOM Table
'! @InputParam1 	sAction 		: Action to be performed e.g. AutoRemoveBasic
'! @InputParam2 	sInvokeOption 	: Method to invoke Remove dialog e.g. menu
'! @InputParam3 	sNodePath 		: Table node path
'! @InputParam4 	sNodeContainer 	: Table node container name
'! @InputParam5 	sKeepSubTree 	: Keep sub tree option
'! @InputParam6 	sButton		 	: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			13 Dec 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_ObjectRemoveOperations","RAC_PSE_ObjectRemoveOperations",OneIteration,"basicremoveandsave","toolbar","0000313/AA-Asm Cockpit~0000316/AA-Asm Cockpit","psebomtable","",""

Option Explicit
Err.Clear

'Declaring varaibles	
Dim sAction,sInvokeOption,sNodePath,sNodeContainer,sKeepSubTree,sButton
Dim sID,sName,sRevision,sPerspective
Dim iRemoveObjectCount
Dim objRemove

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sNodePath = Parameter("sNodePath")
sNodeContainer = Parameter("sNodeContainer")
sKeepSubTree = Parameter("sKeepSubTree")
sButton= Parameter("sButton")

'Selecting node from table
If sNodePath<>"" Then
	Select Case LCase(sNodeContainer)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "psebomtable",""
			LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "Select",sNodePath,"","",""
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "psebomtable_multiselect",""
			LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "MultiSelect",sNodePath,"","",""
	End Select
End If

'inoke Remove dialog
Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditRemove"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "toolbar"	
		LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","RemoveTheReferenceWithoutDeletingTheObject", "",""
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'Use this invoke option when user wants to invoke Remove dialog from outside function
End Select

'Creating object of [ Remove ] dialog
Set objRemove=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_Remove","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ObjectRemoveOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of Remove dialog
If Fn_UI_Object_Operations("RAC_PSE_ObjectRemoveOperations", "Exist", objRemove, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Remove ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Object Remove Operations",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Remove obejct
	Case "basicremoveandsave"
		'Click on [ Yes ] button
		If Fn_UI_JavaButton_Operations("RAC_PSE_ObjectRemoveOperations", "Click", objRemove,"jbtn_Yes")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remove selected object as fail to click on [ Yes ] button","","","","","")	
			Call Fn_ExitTest()
		End If			
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		'Save the changes
		LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","Savethecurrentcontent", "",""
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ObjectRemoveOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

		If JavaWindow("jwnd_StructureManager").Dialog("dlg_ConfirmationDialog").Exist(20) Then
			JavaWindow("jwnd_StructureManager").Dialog("dlg_ConfirmationDialog").WinButton("wbtn_Yes").Click
			If Err.Number <> 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save assembly from BOM table due to error [ " & Cstr(Err.Description) & " ]","","","","","") 
				Call Fn_ExitTest()
			End If
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Remove Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remove selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Removed selected object from assembly","","","","","")
		End If
End Select

'Releasing object
Set objRemove=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objRemove=Nothing
	ExitTest
End Function

