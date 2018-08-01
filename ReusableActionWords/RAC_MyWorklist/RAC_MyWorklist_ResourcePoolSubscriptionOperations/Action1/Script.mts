'! @Name 			RAC_MyWorklist_ResourcePoolSubscriptionOperations
'! @Details 		This Action word to perform operations on Resource Pool Subscription dialog
'! @InputParam1 	sAction						: Action name
'! @InputParam2		sInvokeOption				: Resource Pool Subscription dialog invoke option
'! @InputParam3		sPerspective				: Perspective name
'! @InputParam4		sTreeNode					: MyWorklist tree node
'! @InputParam5		sGroupRoleFilterOption		: Group Role Filter Option
'! @InputParam6		sGroupRoleUserAutomationID	: Group Role user automation id
'! @InputParam7		sButton						: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			05 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_ResourcePoolSubscriptionOperations","RAC_MyWorklist_ResourcePoolSubscriptionOperations",OneIteration,"Add","Menu","MyTeamcenter","My Worklist","All","TestUser3draftingOkcAlDrafter","Cancel"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sPerspective,sTreeNode,sGroupRoleFilterOption,sGroupRoleUserAutomationID,sButton
Dim objResourcePoolSubscription
Dim sGroup,sRole

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sTreeNode = Parameter("sTreeNode")
sGroupRoleFilterOption = Parameter("sGroupRoleFilterOption")
sGroupRoleUserAutomationID = Parameter("sGroupRoleUserAutomationID")
sButton = Parameter("sButton")

'Invoking [ Resource Pool Subscription ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","Action1",OneIteration,"Select","ToolsResourcePoolSubscription"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Creating object of [ Resource Pool Subscription ] dialog
Select Case LCase(sPerspective)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter",""
		'Creating object of [ Resource Pool Subscription ] dialog
		Set objResourcePoolSubscription=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","jdlg_ResourcePoolSubscription","")
End Select

'Checking existance of [ Resource Pool Subscription ] dialog
If Fn_UI_Object_Operations("RAC_MyWorklist_ResourcePoolSubscriptionOperations","Exist", objResourcePoolSubscription, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Resource Pool Subscription ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capture business functionality start time	
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_MyWorklist_ResourcePoolSubscriptionOperations",sAction,"","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "Add"
		If sTreeNode="" Then
			sTreeNode="My Worklist"
		End IF
		'Selecting node from my worklist tree
		If Fn_UI_JavaTree_Operations("RAC_MyWorklist_ResourcePoolSubscriptionOperations", "Select",objResourcePoolSubscription,"jtree_MyWorklistTree",sTreeNode,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select mywork list  Tree node [ " & Cstr(sTreeNode) & " ] while performing operation [ " & Cstr(sAction) & " ] on Resource Pool Subscription dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		If sGroupRoleFilterOption="" Then
			sGroupRoleFilterOption="All"
		End IF
		
		'Selecting group role filter option
		objResourcePoolSubscription.JavaRadioButton("GroupRoleFilterOption").SetTOProperty "attached text",sGroupRoleFilterOption
		If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_ResourcePoolSubscriptionOperations", "Set", objResourcePoolSubscription, "jrdb_GroupRoleFilterOption", "ON")=False THen
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Group Role Filter ] Option value as [ " & Cstr(sGroupRoleFilterOption) & " ] while performing operation [ " & Cstr(sAction) & " ] on Resource Pool Subscription dialog","","","","","")
			Call Fn_ExitTest()
		End IF	
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		sGroup = Fn_Setup_GetTestUserDetailsFromExcelOperations("getgroup","",sGroupRoleUserAutomationID)
		sRole =  Fn_Setup_GetTestUserDetailsFromExcelOperations("getrole","",sGroupRoleUserAutomationID)	
		
		'Setting group
		If Fn_UI_JavaList_Operations("RAC_MyWorklist_ResourcePoolSubscriptionOperations", "Select", objResourcePoolSubscription,"jlst_Group",sGroup, "", "") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Group value as [ " & Cstr(sGroup) & " ] while performing operation [ " & Cstr(sAction) & " ] on Resource Pool Subscription dialog","","","","","")
			Call Fn_ExitTest()	
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Setting role
		If Fn_UI_JavaList_Operations("RAC_MyWorklist_ResourcePoolSubscriptionOperations", "Select", objResourcePoolSubscription,"jlst_Role",sRole, "", "") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Role value as [ " & Cstr(sRole) & " ] while performing operation [ " & Cstr(sAction) & " ] on Resource Pool Subscription dialog","","","","","")
			Call Fn_ExitTest()	
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on Add button
		If Fn_UI_JavaButton_Operations("RAC_MyWorklist_ResourcePoolSubscriptionOperations", "Click", objResourcePoolSubscription,"jbtn_Add")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Add ] button while performing operation [ " & Cstr(sAction) & " ] on Resource Pool Subscription dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)		
		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_MyWorklist_ResourcePoolSubscriptionOperations", "Click", objResourcePoolSubscription,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button while performing operation [ " & Cstr(sAction) & " ] on Resource Pool Subscription dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
		End IF
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MyWorklist_ResourcePoolSubscriptionOperations",sAction,"","")
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added Resource Pool Subscription of Group [ " & Cstr(sGroup) & " ] and Role [ " & Cstr(sRole) & " ]","","","","","")
End Select

'Releasing object
Set objResourcePoolSubscription=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objResourcePoolSubscription=Nothing
	ExitTest
End Function

