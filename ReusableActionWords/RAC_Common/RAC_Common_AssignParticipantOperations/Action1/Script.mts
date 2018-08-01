'! @Name 			RAC_Common_AssignParticipantOperations
'! @Details 		Action word to perform operations on Assign participant dialog
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam3 	sInvokeOption		: Assign participant creation dialog invoke option
'! @InputParam4 	sPerspective	 	: Perspective name in which user wants to perform operations on Assign participant dialog
'! @InputParam5 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			28 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_AssignParticipantOperations","RAC_Common_AssignParticipantOperations",OneIteration,"Add","Menu","MyTeamcenter","OK"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sPerspective,sButton
Dim objAssignParticipants
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

'Invoking [ Assign participant ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsAssignParticipants"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Invoke dialog From Summary tab
	Case "summarytablink"
		LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Summary","Participants"		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
		JavaWindow("jwnd_DefaultWindow").JavaObject("to_class:=JavaObject","text:=Assign Participants\.\.\.").Click 2,2,"LEFT"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_AssignParticipantOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Creating object of [ Assign participant ] dialog
Select Case LCase(sPerspective)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter",""
		Set objAssignParticipants =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_AssignParticipants","")		
End Select

'Checking existance of [  Assign participant ] dialog
If Fn_UI_Object_Operations("RAC_Common_AssignParticipantOperations","Exist", objAssignParticipants, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Assign Participants ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_AssignParticipantOperations",sAction,"","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Add Assign participants
	Case "Add"
		'Selecting Participants            
		If dictAssignParticipantInfo("ParticipantNode")<>"" Then
			If Fn_UI_JavaTree_Operations("RAC_Common_AssignParticipantOperations","Select",objAssignParticipants,"jtree_ParticipantsTree",dictAssignParticipantInfo("ParticipantNode"),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select participant tree node [ " & Cstr(dictAssignParticipantInfo("ParticipantNode")) & " ] while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
		'Selecting Organization
		If dictAssignParticipantInfo("OrganizationNode")<>"" Then
			aNode = Split(dictAssignParticipantInfo("OrganizationNode"),"~")
			If Ubound(aNode) > 1 Then
				For iCounter = 0 to Ubound(aNode) - 1
					If iCounter = 0 Then
						sNode = aNode(0)
					Else
						sNode = sNode & "~" & aNode(iCounter)
					End If
					'expanding node
					Call Fn_UI_JavaTree_Operations("RAC_Common_AssignParticipantOperations", "Expand",objAssignParticipants, "jtree_OrganizationTree",sNode,"","")
					Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)					
				Next
			End If
			
			If Fn_UI_JavaTree_Operations("RAC_Common_AssignParticipantOperations", "Select",objAssignParticipants, "jtree_OrganizationTree",dictAssignParticipantInfo("OrganizationNode"),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select organization tree node [ " & Cstr(dictAssignParticipantInfo("OrganizationNode")) & " ] while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)			
		End If
		'Selecting Project
		If dictAssignParticipantInfo("ProjectName")<>"" Then
			If Fn_UI_JavaTab_Operations("RAC_Common_AssignParticipantOperations","Select",objAssignParticipants,"jtab_MainTab","Project Teams")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project [ " & Cstr(dictAssignParticipantInfo("ProjectName")) & " ] from project list while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)		
				
			'Selecting Project from List
			If Fn_UI_JavaList_Operations("RAC_Common_AssignParticipantOperations", "Select", objAssignParticipants,"jlst_ProjectsList",dictAssignParticipantInfo("ProjectName"), "", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project [ " & Cstr(dictAssignParticipantInfo("ProjectName")) & " ] from project list while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		'Selecting Project  
		If dictAssignParticipantInfo("ProjectNode")<>"" Then
			If Fn_UI_JavaTab_Operations("RAC_Common_AssignParticipantOperations","Select",objAssignParticipants,"jtab_MainTab","Project Teams")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project [ " & Cstr(dictAssignParticipantInfo("ProjectNode")) & " ] from project list while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTree_Operations("RAC_Common_AssignParticipantOperations", "Select",objAssignParticipants, "jtree_ProjectsTree",dictAssignParticipantInfo("ProjectNode"),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project tree node [ " & Cstr(dictAssignParticipantInfo("ProjectNode")) & " ] while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)			
		End If
		'Selecting member option
		If dictAssignParticipantInfo("MemberOption")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_AssignParticipantOperations","SetTOProperty",objAssignParticipants.JavaRadioButton("jrdb_Member"),"","attached text",dictAssignParticipantInfo("MemberOption"))
			If Fn_UI_JavaRadioButton_Operations("RAC_Common_AssignParticipantOperations","Set",objAssignParticipants,"jrdb_Member","ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select resource pool option [ " & Cstr(dictAssignParticipantInfo("MemberOption")) & " ] while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		'Selecting group option
		If dictAssignParticipantInfo("GroupOption")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_AssignParticipantOperations","SetTOProperty",objAssignParticipants.JavaRadioButton("jrdb_Group"),"","attached text",dictAssignParticipantInfo("GroupOption"))
			If Fn_UI_JavaRadioButton_Operations("RAC_Common_AssignParticipantOperations","Set",objAssignParticipants,"jrdb_Group","ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select resource pool option [ " & Cstr(dictAssignParticipantInfo("GroupOption")) & " ] while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If		
		'Click on Add button
		If Fn_UI_JavaButton_Operations("RAC_Common_AssignParticipantOperations","Click",objAssignParticipants,"jbtn_Add")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Add ] button while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_AssignParticipantOperations", "Click", objAssignParticipants,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End IF
		
		'Capturing execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_AssignParticipantOperations",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ Add ] operation on [ Assign Participant ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added assign Participant : [  " & Cstr(dictAssignParticipantInfo("ParticipantNode")) & " ] and User : [  " & Cstr(dictAssignParticipantInfo("OrganizationNode")) & " ] ","","","","","")
		End If	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Remove Assign participants
	Case "Remove","RemoveEXT"
		If sAction="Remove" Then
			'Selecting Participants            
			dictAssignParticipantInfo("ParticipantNode")= Fn_Setup_GetTestUserDetailsFromExcelOperations("getparticipantstreeusernodepath","",dictAssignParticipantInfo("ParticipantNode"))
		ElseIf sAction="RemoveEXT" Then
			dictAssignParticipantInfo("ParticipantNode")= Fn_Setup_GetTestUserDetailsFromExcelOperations("getparticipantstreeusernodepathext","",dictAssignParticipantInfo("ParticipantNode"))
		End If	
		
		If dictAssignParticipantInfo("ParticipantNode")<>"" Then
			aNode = Split(dictAssignParticipantInfo("ParticipantNode"),"~")
			If Ubound(aNode) > 1 Then
				For iCounter = 0 to Ubound(aNode) - 1
					If iCounter = 0 Then
						sNode = aNode(0)
					Else
						sNode = sNode & "~" & aNode(iCounter)
					End If
					'expanding node
					Call Fn_UI_JavaTree_Operations("RAC_Common_AssignParticipantOperations", "Expand",objAssignParticipants, "jtree_ParticipantsTree",sNode,"","")
					Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)					
				Next
			End If
			
			If Fn_UI_JavaTree_Operations("RAC_Common_AssignParticipantOperations", "Select",objAssignParticipants, "jtree_ParticipantsTree",dictAssignParticipantInfo("ParticipantNode"),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Participants tree node [ " & Cstr(dictAssignParticipantInfo("ParticipantNode")) & " ] while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)			
		End If
		
		'Click on Reomve button
		If Fn_UI_JavaButton_Operations("RAC_Common_AssignParticipantOperations","Click",objAssignParticipants,"jbtn_Remove")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Remove ] button while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)		
		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_AssignParticipantOperations", "Click", objAssignParticipants,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button while performing [ " & Cstr(sAction) & " ] operation on Assign participant dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End IF
		
		'Capturing execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_AssignParticipantOperations",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ Remove ] operation on [ Assign Participant ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Remove assign Participant : [  " & Cstr(dictAssignParticipantInfo("ParticipantNode")) & " ] from Participant list","","","","","")
		End If		
End Select

'Creating object of [ Assign participant ] dialog
Set objAssignParticipants=Nothing

Function Fn_ExitTest()
	'Creating object of [ Assign participant ] dialog
	Set objAssignParticipants=Nothing
	ExitTest
End Function

