'! @Name 			RAC_PSE_RemoveDesignFromProductOperations
'! @Details 		Action word to remove Design From Product In BOM Table
'! @InputParam1 	sAction 		: Action to be performed e.g. AutoRemoveBasic
'! @InputParam2 	sInvokeOption 	: Method to invoke Remove dialog e.g. menu
'! @InputParam3 	sNodePath 		: Table node path
'! @InputParam4 	sNodeContainer 	: Table node container name
'! @InputParam5 	sButton		 	: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			11 Apr 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_PSE_RemoveDesignFromProductOperations","RAC_PSE_RemoveDesignFromProductOperations",OneIteration,"basicremoveandsave","toolbar","0000313/AA-Asm Cockpit~0000316/AA-Asm Cockpit","psebomtable",""

Option Explicit
Err.Clear

'Declaring varaibles	
Dim sAction,sInvokeOption,sNodePath,sNodeContainer,sButton
Dim objRemoveDesignFromProduct

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sNodePath = Parameter("sNodePath")
sNodeContainer = Parameter("sNodeContainer")
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
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditRemoveDesignFromProduct"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "toolbar"	
		LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","RemoveDesignfromProduct", "",""
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'Use this invoke option when user wants to invoke Remove dialog from outside function
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_RemoveDesignFromProductOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Creating object of [ Remove ] dialog
Set objRemoveDesignFromProduct=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_RemoveDesignFromProduct","")

'Checking existance of Remove dialog
If Fn_UI_Object_Operations("RAC_PSE_ObjectRemoveOperations", "Exist", objRemoveDesignFromProduct, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Remove Design(s) from Product ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Object Remove Design from Product Operations",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Remove Design From Product obejct
	Case "basicremovedesignfromproduct"
		'Click on [ Yes ] button
		If Fn_UI_JavaButton_Operations("RAC_PSE_ObjectRemoveOperations", "Click", objRemoveDesignFromProduct,"jbtn_Yes")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remove design from product as fail to click on [ Yes ] button","","","","","")	
			Call Fn_ExitTest()
		End If			
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on [ OK ] button
		If Fn_UI_Object_Operations("RAC_PSE_ObjectRemoveOperations", "Exist", objRemoveDesignFromProduct, GBL_DEFAULT_TIMEOUT,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remove design from product as [ Remove Design(s) from Product ] confirmation dialog does not exist","","","","","")
			Call Fn_ExitTest()
		End If

		If Fn_UI_JavaButton_Operations("RAC_PSE_ObjectRemoveOperations", "Click", objRemoveDesignFromProduct,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remove design from product as fail to click on [ OK ] button","","","","","")	
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Remove design from product Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remove selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully removed selected design from product","","","","","")
		End If
End Select

'Releasing object
Set objRemoveDesignFromProduct=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objRemoveDesignFromProduct=Nothing
	ExitTest
End Function

