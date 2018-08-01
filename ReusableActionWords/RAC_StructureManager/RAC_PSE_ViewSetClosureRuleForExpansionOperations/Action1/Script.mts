'! @Name 			RAC_PSE_ViewSetClosureRuleForExpansionOperations
'! @Details 		Action word to perform operation on View\Set Closure Rule For Expansion dialog
'! @InputParam1 	sAction 			: Action to be performed e.g. AutoReplaceBasic
'! @InputParam2 	sInvokeOption 		: Method to invoke Replace dialog e.g. menu
'! @InputParam3 	sClosureRule		: Closure rule name
'! @InputParam4 	sPrimaryObject		: Primary object name
'! @InputParam5 	sColumn 			: Column name
'! @InputParam6 	sValue			 	: Value
'! @InputParam7 	sButton			 	: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			17 Jun 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_ViewSetClosureRuleForExpansionOperations","RAC_PSE_ViewSetClosureRuleForExpansionOperations",OneIteration,"Set","","BOM_View","","","",""
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_ViewSetClosureRuleForExpansionOperations","RAC_PSE_ViewSetClosureRuleForExpansionOperations",OneIteration,"Unset","","EBOM_View","","","",""

Option Explicit
Err.Clear

'Declaring varaibles	
Dim sAction,sInvokeOption,sClosureRule,sPrimaryObject,sColumn,sValue,sButton
Dim objViewSetClosureRuleForExpansion

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sClosureRule = Parameter("sClosureRule")
sPrimaryObject = Parameter("sPrimaryObject")
sColumn = Parameter("sColumn")
sValue = Parameter("sValue")
sButton = Parameter("sButton")

'inoke ViewSetClosureRuleForExpansion dialog
Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsViewSetClosureRuleForExpansion"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'Use this invoke option when user wants to invoke Replace dialog from outside function
End Select

'Creating object of [ ViewSetClosureRuleForExpansion ] dialog
Set objViewSetClosureRuleForExpansion=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jwnd_ViewSetClosureRuleForExpansion","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ViewSetClosureRuleForExpansionOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of ViewSetClosureRuleForExpansion dialog
If Fn_UI_Object_Operations("RAC_PSE_ViewSetClosureRuleForExpansionOperations", "Exist", objViewSetClosureRuleForExpansion, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ View\Set Closure Rule For Expansion ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If	
		
'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","ViewSetClosureRuleForExpansion",sAction,"","")

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Search and Replace obejct
	Case "Set","Unset"
		If sClosureRule<>"" Then
			'Selecting closure rule
			If Fn_UI_JavaList_Operations("RAC_PSE_ViewSetClosureRuleForExpansionOperations", "Select", objViewSetClosureRuleForExpansion,"jlst_ClosureRules",sClosureRule, "", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to " & sAction & " closure rule [ " & Cstr(sClosureRule) & " ] as fail to select closure rule [ " & Cstr(sClosureRule) & " ] from [ Closure Rules ] list","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		
		If sButton="" Then
			If sAction="Set" Then
				sButton="OK"
			ElseIf sAction="Unset" Then
				sButton="UnsetRule"
			End If
		End If
		'Clicking on button
		If Fn_UI_JavaButton_Operations("RAC_PSE_ViewSetClosureRuleForExpansionOperations","Click",objViewSetClosureRuleForExpansion,"jbtn_" & sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button while performing [ " & Cstr(sAction) & " ] operation on [ View\Set Closure Rule For Expansion ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","ViewSetClosureRuleForExpansion",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to " & sAction & " closure rule [ " & Cstr(sClosureRule) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully " & sAction & " closure rule [ " & Cstr(sClosureRule) & " ]","","","","","")
		End If		
End Select

'Releasing object
Set objViewSetClosureRuleForExpansion=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objViewSetClosureRuleForExpansion=Nothing
	ExitTest
End Function
