'! @Name 			RAC_Common_RemoteImportOperations
'! @Details 		Action word to perform remote import opertations
'! @InputParam1 	sAction 						: String to indicate what action is to be performed
'! @InputParam2 	sInvokeOption					: Remote Import Option dialog invoke option
'! @InputParam3 	sButton 						: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Mohini Deshmukh Mohini.Deshmukh@sqs.com
'! @Date 			22 Aug 2017
'! @Version 		1.0
'! @Example 		dictRemoteExportInfo("RemoteImportOptions")="true"
'! @Example 		dictRemoteExportInfo("TransferOwnership")="ON"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_RemoteImportOperations","RAC_Common_RemoteImportOperations",OneIteration,"Import","Menu",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sButton
Dim objRemoteImportOptions,objImportFromPLM,objImportRemoteOptions
Dim sOptionSettingsHeader,sOptionSettingsValue
Dim iCounter

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sButton = Parameter("sButton")

'Invoking [ Remote ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsImportRemote"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Creating object of [ Remote Import ] dialog
Set objImportFromPLM = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ImportFromPLM","")
Set objImportRemoteOptions = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ImportRemoteOptions","")
Set objRemoteImportOptions =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_RemoteImportOptions","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_RemoteImportOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of [  Remote ] dialog
If Fn_UI_Object_Operations("RAC_Common_RemoteImportOperations","Exist", objImportFromPLM,"","","")=False Then

	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Remote Import ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_RemoteImportOperations",sAction,"","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Remote
	Case "Import"				
		If LCase(dictRemoteExportInfo("RemoteImportOptions"))="true" Then
			'Click on Set Remote Import Options button
			If Fn_UI_JavaButton_Operations("RAC_Common_RemoteImportOperations","Click",objImportFromPLM,"jbtn_ImportRemote")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remote Import selected objects as fail to click on [ Set Remote Import Options ] button from Remote Import dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			
			sOptionSettingsHeader=""
			sOptionSettingsValue=""
			
			If dictRemoteExportInfo("TransferOwnership")<>"" Then
				sOptionSettingsHeader="Transfer Options"
				sOptionSettingsValue="Transfer ownership"
				
				objRemoteImportOptions.JavaCheckBox("jckb_ImportOption").SetTOProperty "attached text","Transfer ownership"
				If Fn_UI_JavaCheckBox_Operations("RAC_Common_RemoteImportOperations", "Set", objRemoteImportOptions, "jckb_ImportOption", "ON") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remote Import selected objects as fail to select option [ Transfer ownership ] from Remote Import Options dialog","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			End If
						
			'Click on OK button
			If Fn_UI_JavaButton_Operations("RAC_Common_RemoteImportOperations","Click",objRemoteImportOptions,"jbtn_OK")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remote Import selected objects as fail to click on [ OK ] button from Remote Import Options dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)			
		End If
		
		'Click on Yes button
		If Fn_UI_JavaButton_Operations("RAC_Common_RemoteImportOperations","Click",objImportFromPLM,"jbtn_Yes")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remote Import selected objects as fail to click on [ Yes ] button from Remote Import dialog","","","","","")
			Call Fn_ExitTest()
		End If
'		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		If objImportRemoteOptions.Exist(30) Then			
			'Click on Yes button
			If Fn_UI_JavaButton_Operations("RAC_Common_RemoteImportOperations","Click",objImportRemoteOptions,"jbtn_Yes")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remote Import selected objects as fail to click on [ Yes ] button from Options Settings dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
		
		If objImportFromPLM.Exist(6) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remote Import selected objects","","","","","")
			Call Fn_ExitTest()
		End If

		'Capturing execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_RemoteImportOperations",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Remote Import selected objects due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Remote Imported selected objects","","","","","")
		End If		
End Select

'Creating object of [ Remote Import ] dialog
Set objImportFromPLM=Nothing
Set objImportRemoteOptions = Nothing
Set objRemoteImportOptions =Nothing

Function Fn_ExitTest()
	'Creating object of [ Remote Import ] dialog
	Set objImportFromPLM=Nothing
	Set objImportRemoteOptions = Nothing
	Set objRemoteImportOptions =Nothing
	ExitTest
End Function


