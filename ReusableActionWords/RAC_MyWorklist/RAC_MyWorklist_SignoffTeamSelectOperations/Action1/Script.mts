'! @Name RAC_MyWorklist_SignoffTeamSelectOperations
'! @Details Action word to perform operations on Select Signoff team dialog
'! @InputParam1. sAction: Action name
'! @InputParam2. sMode : Mode from where operations perform
'! @Author Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer Kundan Kudale kundan.kudale@sqs.com
'! @Date 16 Jan 2016
'! @Version 1.0
'! @Example dictSignoffTeamInfo("SignoffTeam")="Signoff Team~Users"
'! @Example dictSignoffTeamInfo("OrganizationUser")="Organization~FORD MOTOR COMPANY~FNA~AAM : Audio Amplifier Module"
'! @Example dictSignoffTeamInfo("Button")="OK"
'! @Example LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_SignoffTeamSelectOperations","RAC_MyWorklist_SignoffTeamSelectOperations",OneIteration,"Add","SelectSignoffTeam"

'Declaring variables
Dim objSignoffDialog
Dim aNode,aMulNode
Dim iCount,iCounter,iRowNumber
Dim sNode
Dim sAction,sMode

GBL_CURRENT_EXECUTABLE_APP="RAC"

sAction=Parameter("sAction")
sMode=Parameter("sMode")

'Selecting mode
Select Case sMode
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "SelectSignoffTeam"
		Set objSignoffDialog =Fn_Setup_GetObjectFromXML("RAC_MyWorklist", "SelectSignoffTeam")
		If Not objSignoffDialog.Exist(20) Then
			'Calling menu
			LoadAndRunAction "RAC_Common\RAC_MenuOperations","RAC_MenuOperations",OneIteration,"Select","ActionsPerform","RAC_Menu"	
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
End Select	

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_SignoffTeamSelectOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

Call Fn_Setup_CaptureFunctionExecutionTime("CaptureStartTime","RAC_MyWorklist_SignoffTeamSelectOperations",sMode,"","")
				
Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "Add"
		'Selecting Signoff team
		If dictSignoffTeamInfo("SignoffTeam")<>"" Then
			iRowNumber=Fn_MyWorkList_SignoffTeamTreeGetItemPath(dictSignoffTeamInfo("SignoffTeam"))
			
			If iRowNumber=-1 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Sign Off Team Tree node [ " & Cstr(dictSignoffTeamInfo("SignoffTeam")) & " ] while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			objSignoffDialog.JavaTree("SignOffTeamTree").Object.setSelectionRow iRowNumber
			Call Fn_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
		'Selecting Organization User
		If dictSignoffTeamInfo("OrganizationUser")<>"" Then
			aMulNode=Split(dictSignoffTeamInfo("OrganizationUser"),"^")
			For iCount = 0 To ubound(aMulNode)
				aNode=Split(aMulNode(iCount),"~")
				sNode=aNode(0)
				For iCounter = 1 To ubound(aNode)-1
					sNode=sNode+"~"+aNode(iCounter)
					objSignoffDialog.JavaTree("OrganizationTree").Expand sNode
					Call Fn_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
				Next
				If iCount=0 Then
					dictSignoffTeamInfo("OrganizationUser")=aMulNode(iCount)
				Else
					dictSignoffTeamInfo("OrganizationUser")=dictSignoffTeamInfo("OrganizationUser")+"^"+aMulNode(iCount)
				End If
				Call Fn_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Next
			aMulNode=Split(dictSignoffTeamInfo("OrganizationUser"),"^")
			If Fn_UI_JavaTree_Operations("RAC_MyWorklist_SignoffTeamSelectOperations", "Select",objSignoffDialog, "OrganizationTree",aMulNode(0),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Organization Tree node [ " & Cstr(aMulNode(0)) & " ] while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_ReadyStatusSync(GBL_MAX_SYNC_ITERATIONS)
			For iCount=1 to ubound(aMulNode)
				objSignoffDialog.JavaTree("OrganizationTree").ExtendSelect aMulNode(iCount)
				Call Fn_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			Next
		End If
		'Selecting Resource Pool Member Option
		If dictSignoffTeamInfo("ResourcePoolMemberOption")<>"" Then
			objSignoffDialog.JavaRadioButton("ResourcePoolMemberOption").SetTOProperty "attached text",dictSignoffTeamInfo("ResourcePoolMemberOption")
			If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_SignoffTeamSelectOperations", "Set", objSignoffDialog, "ResourcePoolMemberOption", "ON")=False THen
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Resource Pool Member ] Option while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End IF	
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		'Selecting Resource Pool Group Option
		If dictSignoffTeamInfo("ResourcePoolGroupOption")<>"" Then
			objSignoffDialog.JavaRadioButton("ResourcePoolGroupOption").SetTOProperty "attached text",dictSignoffTeamInfo("ResourcePoolGroupOption")
			If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_SignoffTeamSelectOperations", "Set", objSignoffDialog, "ResourcePoolGroupOption", "ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Resource Group Member ] Option while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		'Selecting Review Quorum Option
		If dictSignoffTeamInfo("ReviewQuorumOption")<>"" Then
			objSignoffDialog.JavaRadioButton("ReviewQuorumOption").SetTOProperty "attached text",dictSignoffTeamInfo("ReviewQuorumOption")
			If Fn_UI_JavaRadioButton_Operations("RAC_MyWorklist_SignoffTeamSelectOperations", "Set", objSignoffDialog, "ReviewQuorumOption", "ON")=False Then					
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Review Quorum ] Option while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End If	
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			'Setting Numeric value				
			If dictSignoffTeamInfo("ReviewQuorumOption")="Numeric" Then
				If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_SignoffTeamSelectOperations", "Set",  objSignoffDialog, "ReviewQuorumNumeric", dictSignoffTeamInfo("Numeric") )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Review Quorum Numeric ] Option while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
					Call Fn_ExitTest()
				End If	
				Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Elseif dictSignoffTeamInfo("ReviewQuorumOption")="Percent" Then
				'Setting percent
				If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_SignoffTeamSelectOperations", "Set",  objSignoffDialog, "ReviewQuorumPercentage", dictSignoffTeamInfo("Percent") )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Review Quorum Percentage ] Option while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
					Call Fn_ExitTest()
				End IF	
				Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End If
		End If
		'Setting Process Description
		If dictSignoffTeamInfo("ProcessDescription")<>"" Then
			If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_SignoffTeamSelectOperations", "Set",  objSignoffDialog, "ProcessDescription", dictSignoffTeamInfo("ProcessDescription") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set [ Process Description ] value as [ " & Cstr(dictSignoffTeamInfo("ProcessDescription")) & " ] while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End If	
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		
		'Setting Process Description
		If dictSignoffTeamInfo("Comments") = "" Then
			dictSignoffTeamInfo("Comments") = "Test comments"			
		End If
		If dictSignoffTeamInfo("Comments")<>"" Then
			If Fn_UI_JavaEdit_Operations("RAC_MyWorklist_SignoffTeamSelectOperations", "Set",  objSignoffDialog, "Comments", dictSignoffTeamInfo("Comments") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set [ Comments ] value as [ " & Cstr(dictSignoffTeamInfo("ProcessDescription")) & " ] while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End IF	
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		
		'Setting ad hoc done checkbox value
		If dictSignoffTeamInfo("Ad-hoc done")<>"" Then 
			Call Fn_UI_Object_SetTOProperty_Operations("RAC_MyWorklist_SignoffTeamSelectOperations","Set",objSignoffDialog.JavaCheckBox("CompleteUserSelection"),"attached text","Ad-hoc done")
			If Fn_UI_JavaCheckBox_Operations("RAC_MyWorklist_SignoffTeamSelectOperations", "Set", objSignoffDialog, "CompleteUserSelection", dictSignoffTeamInfo("Ad-hoc done")) = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set [ Ad-hoc done ] checkbox value as [ " & Cstr(dictSignoffTeamInfo("Ad-hoc done")) & " ] while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End IF	
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If

		'Clciking on [ Add ] button
		if dictSignoffTeamInfo("AddButton") <> "False" then 
		
			If Fn_UI_JavaButton_Operations("RAC_MyWorklist_SignoffTeamSelectOperations","Click",objSignoffDialog, "Add")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ Add ] button while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End IF
		End if
		'dictSignoffTeamInfo.Remove("AddButton")
		Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		If dictSignoffTeamInfo("Button")<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_MyWorklist_SignoffTeamSelectOperations","Click",objSignoffDialog,dictSignoffTeamInfo("Button"))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click [ " & Cstr(dictSignoffTeamInfo("Button")) & " ] button while performing operation [ " & Cstr(sAction) & " ] on Select Signoff team dialog","","","","","")
				Call Fn_ExitTest()
			End IF	
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			
'			'Close warning message if displayed 
'			If objSignoffDialog.JavaDialog("Warning").Exist(20) Then
'				objSignoffDialog.JavaDialog("Warning").JavaButton("OK").Click
'				If objSignoffDialog.Exist(5) Then
'					objSignoffDialog.Close
'				End If
'			End If
'			
'			'Close warning message if displayed 
			If JavaWindow("MyWorkListWindow").JavaWindow("TcDefaultApplet").JavaDialog("Warning").Exist(20) Then
				JavaWindow("MyWorkListWindow").JavaWindow("TcDefaultApplet").JavaDialog("Warning").JavaButton("OK").Click
'				If objSignoffDialog.Exist(5) Then
'					objSignoffDialog.Close
'				End If
			End If
		End if		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Fail to select signoff team member [ " & Cstr(dictSignoffTeamInfo("OrganizationUser")) & " ] to [ " & Cstr(dictSignoffTeamInfo("SignoffTeam")) & " ] ","","","","","")	
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected signoff team member [ " & Cstr(dictSignoffTeamInfo("OrganizationUser")) & " ] to [ " & Cstr(dictSignoffTeamInfo("SignoffTeam")) & " ] ","","","","","")	
		End If
		Call Fn_Setup_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MyWorklist_SignoffTeamSelectOperations",sMode,"","")
End Select
'Releasing object
Set objSignoffDialog=Nothing

Function Fn_ExitTest()
	Set objSignoffDialog=Nothing
	ExitTest
End Function


Function Fn_MyWorkList_SignoffTeamTreeGetItemPath(sNodeName)
	Dim iItemsCount,iCounter
	Dim sTempNode
	Dim objSignOffTree
	
	Fn_MyWorkList_SignoffTeamTreeGetItemPath=-1
	
	Set objSignOffTree =Fn_Setup_GetObjectFromXML("RAC_MyWorklist", "SelectSignoffTeam")
	Set objSignOffTree=objSignOffTree.JavaTree("SignOffTeamTree")
	
	iItemsCount = objSignOffTree.GetROProperty("items count")
	
	For iCounter = 0 to iItemsCount-1
		sTempNode = objSignOffTree.Object.getPathForRow(iCounter).tostring()
		sTempNode=Replace(sTempNode,"[","")
		sTempNode=Replace(sTempNode,"]","")
		sTempNode=Replace(sTempNode,", ","~")
		
		If sTempNode=sNodeName Then
			Fn_MyWorkList_SignoffTeamTreeGetItemPath = iCounter
			iRowCounter=iCounter
			Exit for
		End If
	Next	

	Set objSignOffTree=Nothing	
End Function


