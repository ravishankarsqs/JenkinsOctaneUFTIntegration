'! @Name 			RAC_MyWorklist_AssignResponsibleParty
'! @Details 		Action word to perform operations on Assign Responsible Party dialog
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam3 	sInvokeOption		: Assign Responsible Party creation dialog invoke option
'! @InputParam4 	sPerspective	 	: Perspective name in which user wants to perform operations on Assign Responsible Party dialog
'! @InputParam5 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			28 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_AssignResponsibleParty","RAC_MyWorklist_AssignResponsibleParty",OneIteration,"Assign","Menu","MyTeamcenter","OK"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sPerspective,sButton
Dim objAssignResponsibleParty
Dim iCounter
Dim aNode
Dim sNode

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sButton = Parameter("sButton")

'Invoking [ Assign Responsible Party ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ActionsAssign"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Invoke dialog From Summary tab
	Case "summarytablink"
		LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Summary","Participants"		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
		JavaWindow("jwnd_DefaultWindow").JavaObject("to_class:=JavaObject","text:=Reassign\.\.\.").Click 2,2,"LEFT"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_AssignResponsibleParty"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Creating object of [ Assign Responsible Party ] dialog
Select Case LCase(sPerspective)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter",""
		Set objAssignResponsibleParty =Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyWorklist_OR","jdlg_AssignResponsibleParty","")		
End Select

'Checking existance of [  Assign Responsible Party ] dialog
If Fn_UI_Object_Operations("RAC_MyWorklist_AssignResponsibleParty","Exist", objAssignResponsibleParty, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Assign Responsible Partys ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_MyWorklist_AssignResponsibleParty",sAction,"","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Assign Responsible Partys
	Case "Assign"
		'Selecting Organization
		If dictAssignResponsiblePartyInfo("OrganizationNode")<>"" Then
			dictAssignResponsiblePartyInfo("OrganizationNode")=Fn_Setup_GetTestUserDetailsFromExcelOperations("getorganizationtreeusernodepath","",dictAssignResponsiblePartyInfo("OrganizationNode"))
			aNode = Split(dictAssignResponsiblePartyInfo("OrganizationNode"),"~")
			If Ubound(aNode) > 1 Then
				For iCounter = 0 to Ubound(aNode) - 1
					If iCounter = 0 Then
						sNode = aNode(0)
					Else
						sNode = sNode & "~" & aNode(iCounter)
					End If
					'expanding node
					Call Fn_UI_JavaTree_Operations("RAC_MyWorklist_AssignResponsibleParty", "Expand",objAssignResponsibleParty, "jtree_OrganizationTree",sNode,"","")
					Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)					
				Next
			End If
			
			If Fn_UI_JavaTree_Operations("RAC_MyWorklist_AssignResponsibleParty", "Select",objAssignResponsibleParty, "jtree_OrganizationTree",dictAssignResponsiblePartyInfo("OrganizationNode"),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select organization tree node [ " & Cstr(dictAssignResponsiblePartyInfo("OrganizationNode")) & " ] while performing [ " & Cstr(sAction) & " ] operation on Assign Responsible Party dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)			
		End If
		
		'Selecting Project
		If dictAssignResponsiblePartyInfo("ProjectName")<>"" Then
			If Fn_UI_JavaTab_Operations("RAC_MyWorklist_AssignResponsibleParty","Select",objAssignResponsibleParty,"jtab_MainTab","Project Teams") Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project [ " & Cstr(dictAssignResponsiblePartyInfo("ProjectName")) & " ] from project list while performing [ " & Cstr(sAction) & " ] operation on Assign Responsible Party dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)		
				
			'Selecting Project from List
			If Fn_UI_JavaList_Operations("RAC_MyWorklist_AssignResponsibleParty", "Select", objAssignResponsibleParty,"jlst_ProjectsList",dictAssignResponsiblePartyInfo("ProjectName"), "", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project [ " & Cstr(dictAssignResponsiblePartyInfo("ProjectName")) & " ] from project list while performing [ " & Cstr(sAction) & " ] operation on Assign Responsible Party dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		'Selecting Project  
		If dictAssignResponsiblePartyInfo("ProjectNode")<>"" Then
			If Fn_UI_JavaTree_Operations("RAC_MyWorklist_AssignResponsibleParty", "Select",objAssignResponsibleParty, "jtree_ProjectsTree",dictAssignResponsiblePartyInfo("ProjectNode"),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project tree node [ " & Cstr(dictAssignResponsiblePartyInfo("ProjectNode")) & " ] while performing [ " & Cstr(sAction) & " ] operation on Assign Responsible Party dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)			
		End If
		'Selecting group option
		If dictAssignResponsiblePartyInfo("GroupOption")<>"" Then
			Call Fn_UI_Object_Operations("RAC_MyWorklist_AssignResponsibleParty","SetTOProperty",objAssignResponsibleParty.JavaRadioButton("jrdb_Group"),"","attached text",dictAssignResponsiblePartyInfo("GroupOption"))
			If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_AssignResponsibleParty","Set",objAssignResponsibleParty,"jrdb_Group","ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select resource pool option [ " & Cstr(dictAssignResponsiblePartyInfo("GroupOption")) & " ] while performing [ " & Cstr(sAction) & " ] operation on Assign Responsible Party dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If		
		'Click on Add button
		If Fn_UI_JavaButton_Operations("RAC_MyWorklist_AssignResponsibleParty","Click",objAssignResponsibleParty,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ OK ] button while performing [ " & Cstr(sAction) & " ] operation on Assign Responsible Party dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)		
		
		'Capturing execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MyWorklist_AssignResponsibleParty",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ Assign ] operation on [ Assign Responsible Party ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully assigned [ Assign Responsible Party ] to user : [  " & Cstr(dictAssignResponsiblePartyInfo("OrganizationNode")) & " ] for selected task","","","","","")
		End If		
End Select

'Creating object of [ Assign Responsible Party ] dialog
Set objAssignResponsibleParty=Nothing

Function Fn_ExitTest()
	'Creating object of [ Assign Responsible Party ] dialog
	Set objAssignResponsibleParty=Nothing
	ExitTest
End Function

