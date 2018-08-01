'! @Name 			RAC_Common_ImportFromOtherSite
'! @Details 		This Action word to perform import from other site operations
'! @InputParam1 	sInvokeOption 		: Action to be performed
'! @InputParam2 	sPerspective 	: automation id
'! @InputParam3 	sMenuLabel 	: Invoke Option (menu or nooption)
'! @Author 			Kundan Kudale kundan.kudale@sqs.com
'! @Reviewer 		Sandeep Navghane sandeep.navghane@sqs.com 
'! @Date 			03 Apr 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ImportFromOtherSite","RAC_Common_ImportFromOtherSite",OneIteration,"menu","","EditCopy"

Option Explicit
Err.Clear

'Declaring variables
Dim sInvokeOption, sPerspective, sMenuLabel
Dim objImportRemote, objImportFromPLM, objImportRemoteOptions
Dim iCounter

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters value in local variables
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sMenuLabel = Parameter("sMenuLabel")

'Get [ User Settings ] object from xml file
Select Case lcase(sPerspective)
	Case "myteamcenter","","my teamcenter"
		Set objImportRemote = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ImportRemote","")
		Set objImportFromPLM = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ImportFromPLM","")
		Set objImportRemoteOptions = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ImportRemoteOptions","")
End Select

Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select", sMenuLabel
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ImportFromOtherSite"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'checking existence of Import Remote Dialog
If Not Fn_UI_Object_Operations("RAC_Common_ImportFromOtherSite","Exist", objImportRemote, GBL_DEFAULT_TIMEOUT,"","") Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence on [Import Remote] dialog","","","","","")
	Call Fn_ExitTest()
End If

'Click on yess button on Import Remote Dialog
If Fn_UI_JavaButton_Operations("RAC_Common_ImportFromOtherSite", "Click", objImportRemote,"jbtn_Yes") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [Yes] button on Import Remote dialog","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'checking existence of Import from PLM Dialog
If Not Fn_UI_Object_Operations("RAC_Common_ImportFromOtherSite","Exist", objImportFromPLM, GBL_DEFAULT_TIMEOUT,"","") Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence on [Import From PLM] dialog","","","","","")
	Call Fn_ExitTest()
End If

'Click on yess button on Import from PLM Dialog
If Fn_UI_JavaButton_Operations("RAC_Common_ImportFromOtherSite", "Click", objImportFromPLM,"jbtn_Yes") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [Yes] button on Import From PLM dialog","","","","","")
	Call Fn_ExitTest()
End If

'checking existence of Import Remote Options Dialog
If Not Fn_UI_Object_Operations("RAC_Common_ImportFromOtherSite","Exist", objImportRemoteOptions, GBL_DEFAULT_TIMEOUT,"","") Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence on [Import Remote Options] dialog","","","","","")
	Call Fn_ExitTest()
End If

'Click on yes button on Import Remote Options Dialog
If Fn_UI_JavaButton_Operations("RAC_Common_ImportFromOtherSite", "Click", objImportRemoteOptions,"jbtn_Yes") = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [Yes] button on Import Remote Options dialog","","","","","")
	Call Fn_ExitTest()
End If

'Wait until Import From PLM dialog is displayed
For iCounter = 1 To 20 Step 1
	If Fn_UI_Object_Operations("RAC_Common_ImportFromOtherSite","Exist", objImportFromPLM, GBL_MIN_MICRO_TIMEOUT,"","") Then
		Wait GBL_MIN_MICRO_TIMEOUT
	Else
		Exit For
	End If
Next

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform import from remote operation due to error number [" & Cstr(Err.Number) & "] and error description  [" & Cstr(Err.Description) & "]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing object 
Set objImportRemote = Nothing
Set objImportFromPLM = Nothing
Set objImportRemoteOptions = Nothing
	
Function Fn_ExitTest()
	'Releasing object 
	Set objImportRemote = Nothing
	Set objImportFromPLM = Nothing
	Set objImportRemoteOptions = Nothing
	ExitTest
End Function
