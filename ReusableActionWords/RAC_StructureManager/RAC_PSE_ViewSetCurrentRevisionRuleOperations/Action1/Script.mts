'! @Name 			RAC_PSE_ViewSetCurrentRevisionRuleOperations
'! @Details 		Action word to perform operation on View\Set Revision Rule dialog
'! @InputParam1 	sAction 			: Action to be performed e.g. AutoReplaceBasic
'! @InputParam2 	sInvokeOption 		: Method to invoke Replace dialog e.g. menu
'! @InputParam3 	sRevisionRule		: Closure rule name
'! @InputParam4 	sDetails			: Rule details
'! @InputParam5 	sButton			 	: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			21 Jun 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_ViewSetCurrentRevisionRuleOperations","RAC_PSE_ViewSetCurrentRevisionRuleOperations",OneIteration,"Set","","BOM_View","",""
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_ViewSetCurrentRevisionRuleOperations","RAC_PSE_ViewSetCurrentRevisionRuleOperations",OneIteration,"Unset","","BOM_View","",""

Option Explicit
Err.Clear

'Declaring varaibles	
Dim sAction,sInvokeOption,sRevisionRule,sDetails,sButton
Dim objViewSetCurrentRevisionRule
Dim sTempValue

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sRevisionRule = Parameter("sRevisionRule")
sDetails = Parameter("sDetails")
sButton = Parameter("sButton")

'invoke ViewSetCurrentRevisionRule dialog
Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsRevisionRuleViewSetCurrent"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'Use this invoke option when user wants to invoke Replace dialog from outside function
End Select

'Creating object of [ ViewSetCurrentRevisionRule ] dialog
Set objViewSetCurrentRevisionRule=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_ViewSetCurrentRevisionRule","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ViewSetCurrentRevisionRuleOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of ViewSetCurrentRevisionRule dialog
If Fn_UI_Object_Operations("RAC_PSE_ViewSetCurrentRevisionRuleOperations", "Exist", objViewSetCurrentRevisionRule, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ View\Set Current Revsion Rule ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If	

		
'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","ViewSetCurrentRevisionRule",sAction,"","")

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Search and Replace obejct
	Case "Set"
		If sRevisionRule<>"" Then
			sTempValue=sRevisionRule
			sRevisionRule=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_RevisionRuleValues_APL",sRevisionRule,"")
			
			If Cstr(sRevisionRule)="False" Then
				sRevisionRule=sTempValue
			End If
			
			'Selecting closure rule
			If Fn_UI_JavaList_Operations("RAC_PSE_ViewSetCurrentRevisionRuleOperations", "Select", objViewSetCurrentRevisionRule,"jlst_Rules",sRevisionRule, "", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to " & sAction & " current revision rule [ " & Cstr(sRevisionRule) & " ] as fail to select revision rule [ " & Cstr(sRevisionRule) & " ] from [ Rules ] list","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		'Clicking on button
		If Fn_UI_JavaButton_Operations("RAC_PSE_ViewSetCurrentRevisionRuleOperations","Click",objViewSetCurrentRevisionRule,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button while performing [ " & Cstr(sAction) & " ] operation on [ View\Set Closure Rule For Expansion ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","ViewSetCurrentRevisionRule",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to " & sAction & " revision rule [ " & Cstr(sRevisionRule) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully " & sAction & " revision rule [ " & Cstr(sRevisionRule) & " ]","","","","","")
		End If		
End Select

'Releasing object
Set objViewSetCurrentRevisionRule=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objViewSetCurrentRevisionRule=Nothing
	ExitTest
End Function
