'! @Name 			RAC_Common_WorkflowProcessOperations
'! @Details 		To perform operations on New Process Dialog
'! @InputParam1		sAction 			: Action to be performed
'! @InputParam2 	sInvokeOption 		: New Process dialog invoke option
'! @InputParam3 	sPerspective 		: Perspective name in which user wants to perform operations
'! @InputParam4 	sWorkflowXMLName 	: Workflow xml name
'! @InputParam5 	sButtonName 		: ButtonName
'! @InputParam6 	dictWorkflowProcessInfo : External dictionary parameter for workflow information
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com 
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			25 May 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_WorkflowProcessOperations","RAC_Common_WorkflowProcessOperations",OneIteration,"ModifySession","TestUser4TechnologyNWCSWSTechnician","",""
'! @Example 		dictWorkflowProcessInfo("ProcessName")="Process1"
'! @Example 		dictWorkflowProcessInfo("Description")="Process description"
'! @Example 		dictWorkflowProcessInfo("ProcessTemplate")="GOG4ALCADPartRelease"
'! @Example 		dictWorkflowProcessInfo("ResourceReviewerNode")="Reviewer~Profiles~*/drafter/1"
'! @Example 		dictWorkflowProcessInfo("OrganizationReviewerNode")="Organization~al~okc~drafting~drafter~Sunny Ruparel (502425666)"
'! @Example 		dictWorkflowProcessInfo("ResourceApproverNode")="Approver~Profiles~*/engineer/1"
'! @Example 		dictWorkflowProcessInfo("OrganizationApproverNode")="Organization~al~okc~engineering~engineer~Sunny Ruparel (502425666)"
'! @Example 		dictWorkflowProcessInfo("ResourceIssuerNode")="Issuer~Profiles~*/doc-controller/1"
'! @Example 		dictWorkflowProcessInfo("OrganizationIssuerNode")="Organization~al~okc~drafting~doc-controller~Sunny Ruparel (502425666)"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_WorkflowProcessOperations","RAC_Common_WorkflowProcessOperations",OneIteration,"AssignAllTasks","menu","myteamcenter","RAC_ProcessTemplateName_WF","OK"
'! @Example 		dictWorkflowProcessInfo("ProcessName")="Process1"
'! @Example 		dictWorkflowProcessInfo("Description")="Process description"
'! @Example 		dictWorkflowProcessInfo("ProcessTemplate")="GOG4ALCADPartRelease"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_WorkflowProcessOperations","RAC_Common_WorkflowProcessOperations",OneIteration,"Assign","menu","myteamcenter","RAC_ProcessTemplateName_WF","OK"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_WorkflowProcessOperations","RAC_Common_WorkflowProcessOperations",OneIteration,"ClickButton","nooption","myteamcenter","","Cancel"

Option Explicit

'Declaring variables
Dim sAction,sInvokeOption,sPerspective,sWorkflowXMLName,sButtonName
Dim objStaticTexts,objNewProcessDialog
Dim bFlag
Dim sNode,sNodeName
Dim aMulNode,aNode
Dim iCounter,iCount
Dim objDescription,objChild,objComboBox

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get parameter values in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sWorkflowXMLName = Parameter("sWorkflowXMLName")	
sButtonName = Parameter("sButtonName")

'Retrive workflow template
If sWorkflowXMLName<>"" Then
	dictWorkflowProcessInfo("ProcessTemplate")=Fn_FSOUtil_XMLFileOperations("getvalue",sWorkflowXMLName,dictWorkflowProcessInfo("ProcessTemplate"),"")
End If

Select Case lCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Select menu [File -> New - > Workflow Process...]
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileNewWorkflowProcess"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
	Case "summarytablink"
		LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Summary","Overview"		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
		JavaWindow("jwnd_DefaultWindow").JavaObject("to_class:=JavaObject","text:=New Workflow Process\.\.\.").Click 2,2,"LEFT"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	   
  	 Case "nooption"
		'do nothing
End Select

'creating object of [ New Process ] dialog
Select Case lcase(sPerspective)
	Case "","myteamcenter"
		Set objNewProcessDialog = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR", "jdlg_NewProcessDialog","")
	Case "structure manager","structuremanager"
	    Set objNewProcessDialog = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR", "jdlg_NewProcessDialog@2","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_WorkflowProcessOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of [ New Process Dialog ]
If Fn_UI_Object_Operations("RAC_Common_WorkflowProcessOperations", "Exist", objNewProcessDialog, "","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ New Process ] dialog as [ New Process ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Captire execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Workflow Process Operations",sAction,"","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to click on [ New Process ] dialog button
	Case "ClickButton"
		If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_" & sButtonName)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on  button [ " & Cstr(sButtonName) & " ] of [ New Process Dialog ]","","","","","")
			Call Fn_ExitTest()
		End If	
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Workflow Process Operations",sAction,"","")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to assign workflow
	Case "Assign","AssignEXT","AssignEXT2"
		If  dictWorkflowProcessInfo("ProcessName") <>"" Then
			'Set Process Name
			If Fn_UI_JavaEdit_Operations("RAC_Common_WorkflowProcessOperations", "Set", objNewProcessDialog, "jedt_ProcessName", dictWorkflowProcessInfo("ProcessName"))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set process name value as [ " & Cstr(dictWorkflowProcessInfo("ProcessName")) & " ] while performing assign workflow operation","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		If dictWorkflowProcessInfo("Description")<> "" Then
			'Set Process description
			If Fn_UI_JavaEdit_Operations("RAC_Common_WorkflowProcessOperations", "Set", objNewProcessDialog, "jedt_Description", dictWorkflowProcessInfo("Description"))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Description value as [ " & Cstr(dictWorkflowProcessInfo("Description")) & " ] while performing assign workflow operation","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		'Selecting process temaplte filter 
		If dictWorkflowProcessInfo("ProcessTemplateFilter")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_WorkflowProcessOperations","settoproperty", objNewProcessDialog.JavaRadioButton("jrdb_ProcessTemplateFilter"),"","attached text", dictWorkflowProcessInfo("ProcessTemplateFilter"))
			If Fn_UI_JavaRadioButton_Operations("RAC_Common_WorkflowProcessOperations", "Set", objNewProcessDialog.JavaRadioButton("jrdb_ProcessTemplateFilter"), "", "ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Process Template Filter option [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplateFilter")) & " ] while performing assign workflow operation","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		'Set Process Template from Drop down list
		If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_ProcessTemplateDropDownButton")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on process Template drop down button while performing assign workflow operation","","","","","")
			Call Fn_ExitTest()
		End IF
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)

		bFlag=False
		'selecting process template
		Set  objStaticTexts = Fn_UI_Object_GetChildObjects("RAC_Common_WorkflowProcessOperations", objNewProcessDialog, "Class Name", "JavaStaticText")
		For iCounter = 0 to objStaticTexts.count-1
			If  objStaticTexts(iCounter).getROProperty("label") = dictWorkflowProcessInfo("ProcessTemplate") Then
				objStaticTexts(iCounter).Click 1,1
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
				bFlag=True
				Exit for
			End If
		Next
		Set objStaticTexts = Nothing
		
		If bFlag=False then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Process Template [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ] while performing assign workflow operation","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on "OK" button
		'objNewProcessDialog.highlight
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		If sButtonName<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_" & sButtonName) = False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButtonName) & " ] while performing assign workflow operation","","","","","")
				Call Fn_ExitTest()
			End if
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If sAction <> "AssignEXT" Then
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
				bFlag=False
				For iCounter = 0 To 2
					If objNewProcessDialog.Exist(2) Then
						Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Else
						bFlag=True
						Exit For
					End If								
				Next
			Else
				bFlag=True
			End If
			
			If sAction = "AssignEXT2" Then
				If objNewProcessDialog.Exist(0) Then
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					bFlag=False
					For iCounter = 1 To 10
						If objNewProcessDialog.Exist(2) Then
							Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
						Else
							bFlag=True
							Exit For
						End If								
					Next
				End If
			End If
			
			If sAction <> "AssignEXT" Then
				If sButtonName="OK" or sButtonName="jbtn_OK" Then
					If JavaWindow("jwnd_DefaultWindow").JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_Warning").Exist(10) Then
						If Instr(1,JavaWindow("jwnd_DefaultWindow").JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_Warning").JavaEdit("jedt_WarningMessage").GetROProperty("value"),"successful") Then
							JavaWindow("jwnd_DefaultWindow").JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_Warning").JavaButton("jbtn_OK").Click
						End If
					End If
				End IF
			End IF
		Else
			bFlag=True
		End If
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to assign workflow [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Workflow Process Operations",sAction,"","")
			If sAction = "AssignEXT" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully select workflow [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ] and click on [ OK ] button","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully assigned workflow [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ]","","","","","")
			End If	
		End If
		dictWorkflowProcessInfo.RemoveAll
				
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 							
	Case "AssignAllTasks"
		'Select AssignAllTasks Tab	
		If Fn_UI_JavaTab_Operations("RAC_Common_WorkflowProcessOperations", "Select", objNewProcessDialog, "jtab_Tab", "Assign All Tasks") =False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select tab [ Assign All Tasks ] while performing assign workflow with all task operation","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		'Set  Process Name
		If  dictWorkflowProcessInfo("ProcessName") <>"" Then
			'Set Process Name
			If Fn_UI_JavaEdit_Operations("RAC_Common_WorkflowProcessOperations", "Set", objNewProcessDialog, "jedt_ProcessName", dictWorkflowProcessInfo("ProcessName"))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set process name value as [ " & Cstr(dictWorkflowProcessInfo("ProcessName")) & " ] while performing assign workflow with all task operation","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		If dictWorkflowProcessInfo("Description") <> "" Then
			'Set Process description
			If Fn_UI_JavaEdit_Operations("RAC_Common_WorkflowProcessOperations", "Set", objNewProcessDialog, "jedt_Description", dictWorkflowProcessInfo("Description"))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set process Description value as [ " & Cstr(dictWorkflowProcessInfo("Description")) & " ] while performing assign workflow with all task operation","","","","","")
				Call Fn_ExitTest()										 
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If				
		'Set Process Template from Drop down list
		If dictWorkflowProcessInfo("ProcessTemplate") <> "" Then
			'Set Process Template from Drop down list
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_ProcessTemplateDropDwonButton")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on process Template drop down button while performing assign workflow with assign all task operation","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			bFlag=False	
			Set  objStaticTexts = Fn_UI_Object_GetChildObjects("RAC_Common_WorkflowProcessOperations", objNewProcessDialog, "Class Name", "JavaStaticText")
			For iCounter = 0 to objStaticTexts.count-1
				If  objStaticTexts(iCounter).getROProperty("label") = dictWorkflowProcessInfo("ProcessTemplate") Then
					objStaticTexts(iCounter).Click 1,1
					Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
					bFlag=True
					Exit for
				End If
			Next
			Set objStaticTexts = Nothing
			If bFlag=False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Process Template [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ] while performing assign workflow with all task operation","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		'Selecting Resource Reviewer
		If dictWorkflowProcessInfo("ResourceEngineeringConfiguratorNode")<>"" Then 'ResourceDocControllerNode = ResourceEngineeringConfiguratorNode
			If dictWorkflowProcessInfo("ProcessTemplate") <> "" Then
				dictWorkflowProcessInfo("ResourceEngineeringConfiguratorNode")=dictWorkflowProcessInfo("ProcessTemplate") & "~" & dictWorkflowProcessInfo("ResourceEngineeringConfiguratorNode")
			End If
			If Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode(objNewProcessDialog.JavaTree("jtree_Resources"),dictWorkflowProcessInfo("ResourceEngineeringConfiguratorNode"), "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Resource Reviewer Node [ " & Cstr(dictWorkflowProcessInfo("ResourceEngineeringConfiguratorNode")) & " ] while performing assign workflow with all task operation","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			aMulNode=Split(dictWorkflowProcessInfo("OrganizationEngineeringConfiguratorNode"),"^") 'OrganizationDocControllerNode = OrganizationEngineeringConfiguratorNode
			For iCount = 0 To ubound(aMulNode)
				aNode=Split(aMulNode(iCount),"~")
				sNode=aNode(0)
				For iCounter = 1 To ubound(aNode)-1
					sNode=sNode & "~" & aNode(iCounter)
					Call Fn_UI_JavaTree_Operations("RAC_Common_WorkflowProcessOperations","Expand",objNewProcessDialog,"jtree_OrganizationTree",sNode,"","")
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
				Next
				If iCount=0 Then
					sNodeName = aMulNode(iCount)
				Else
					sNodeName  =sNodeName & "^" & aMulNode(iCount)
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Next
			aMulNode=Split(sNodeName,"^")
			Call Fn_UI_JavaTree_Operations("RAC_Common_WorkflowProcessOperations","Select",objNewProcessDialog,"jtree_OrganizationTree",aMulNode(0),"","")
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			For iCount=1 to ubound(aMulNode)
				objNewProcessDialog.JavaTree("jtree_OrganizationTree").ExtendSelect aMulNode(iCount)
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Next
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_Add")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Add ] button while performing assign workflow with all task operation","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		'Selecting Approver Reviewer
		If dictWorkflowProcessInfo("ResourceManufacturingConfiguratorNode")<>"" Then 'ResourceApproverNode = ResourceManufacturingConfiguratorNode
			If dictWorkflowProcessInfo("ProcessTemplate") <> "" Then
				dictWorkflowProcessInfo("ResourceManufacturingConfiguratorNode")=dictWorkflowProcessInfo("ProcessTemplate") & "~" & dictWorkflowProcessInfo("ResourceManufacturingConfiguratorNode")
			End If
			If Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode(objNewProcessDialog.JavaTree("jtree_Resources"),dictWorkflowProcessInfo("ResourceManufacturingConfiguratorNode"), "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Resource Approver Node [ " & Cstr(dictWorkflowProcessInfo("ResourceManufacturingConfiguratorNode")) & " ] while performing assign workflow with all task operation","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			aMulNode=Split(dictWorkflowProcessInfo("OrganizationManufacturingConfiguratorNode"),"^") 'OrganizationApproverNode = OrganizationManufacturingConfiguratorNode
			For iCount = 0 To ubound(aMulNode)
				aNode=Split(aMulNode(iCount),"~")
				sNode=aNode(0)
				For iCounter = 1 To ubound(aNode)-1
					sNode=sNode & "~" & aNode(iCounter)
					Call Fn_UI_JavaTree_Operations("RAC_Common_WorkflowProcessOperations","Expand",objNewProcessDialog,"jtree_OrganizationTree",sNode,"","")
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
				Next
				If iCount=0 Then
					sNodeName = aMulNode(iCount)
				Else
					sNodeName  =sNodeName & "^" & aMulNode(iCount)
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Next
			aMulNode=Split(sNodeName,"^")
			Call Fn_UI_JavaTree_Operations("RAC_Common_WorkflowProcessOperations","Select",objNewProcessDialog,"jtree_OrganizationTree",aMulNode(0),"","")
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			For iCount=1 to ubound(aMulNode)
				objNewProcessDialog.JavaTree("jtree_OrganizationTree").ExtendSelect aMulNode(iCount)
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Next
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_Add")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Add ] button while performing assign workflow with all task operation","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		'Selecting Approver Reviewer
		If dictWorkflowProcessInfo("ResourcePlanningNode")<>"" Then 'ResourceIssuerNode = ResourcePlanningNode
			If dictWorkflowProcessInfo("ProcessTemplate") <> "" Then
				dictWorkflowProcessInfo("ResourcePlanningNode")=dictWorkflowProcessInfo("ProcessTemplate") & "~" & dictWorkflowProcessInfo("ResourcePlanningNode")
			End If
			If Fn_RAC_SelectNewProcessAssignAllTaskResourceTreeNode(objNewProcessDialog.JavaTree("jtree_Resources"),dictWorkflowProcessInfo("ResourcePlanningNode"), "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Resource Issuer Node [ " & Cstr(dictWorkflowProcessInfo("ResourcePlanningNode")) & " ] while performing assign workflow with all task operation","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			aMulNode=Split(dictWorkflowProcessInfo("OrganizationPlanningConfiguratorNode"),"^")
			For iCount = 0 To ubound(aMulNode)
				aNode=Split(aMulNode(iCount),"~")
				sNode=aNode(0)
				For iCounter = 1 To ubound(aNode)-1
					sNode=sNode & "~" & aNode(iCounter)
					Call Fn_UI_JavaTree_Operations("RAC_Common_WorkflowProcessOperations","Expand",objNewProcessDialog,"jtree_OrganizationTree",sNode,"","")
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
				Next
				If iCount=0 Then
					sNodeName = aMulNode(iCount)
				Else
					sNodeName  =sNodeName & "^" & aMulNode(iCount)
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Next
			aMulNode=Split(sNodeName,"^")
			Call Fn_UI_JavaTree_Operations("RAC_Common_WorkflowProcessOperations","Select",objNewProcessDialog,"jtree_OrganizationTree",aMulNode(0),"","")
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			For iCount=1 to ubound(aMulNode)
				objNewProcessDialog.JavaTree("jtree_OrganizationTree").ExtendSelect aMulNode(iCount)
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			Next
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_Add")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Add ] button while performing assign workflow with all task operation","","","","","")
				Call Fn_ExitTest()
			End IF
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		'Click on button
		If sButtonName<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_" & sButtonName) = False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButtonName) & " ] while performing assign workflow operation","","","","","")
				Call Fn_ExitTest()
			End if				
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			bFlag=False
			For iCounter = 0 To 2
				If objNewProcessDialog.Exist(2) Then
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				Else
					bFlag=True
					Exit For
				End If								
			Next
		Else
			bFlag=True
		End If
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to assign workflow [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Workflow Process Operations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully assigned workflow [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ]","","","","","")
		End If	
		dictWorkflowProcessInfo.RemoveAll
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "VerifyWorkflowTemplateNotAvailable"
		'Set Process Template from Drop down list
		If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_ProcessTemplateDropDownButton")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on process Template drop down button from assign workflow dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)

		bFlag=False
		'selecting process template
		Set  objStaticTexts = Fn_UI_Object_GetChildObjects("RAC_Common_WorkflowProcessOperations", objNewProcessDialog, "Class Name", "JavaStaticText")
		For iCounter = 0 to objStaticTexts.count-1
			If  objStaticTexts(iCounter).getROProperty("label") = dictWorkflowProcessInfo("ProcessTemplate") Then
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
				If objStaticTexts(iCounter).getROProperty("abs_x")<>"" or objStaticTexts(iCounter).getROProperty("abs_y")<>"" Then
					bFlag=True
					Exit for
				End If				
			End If
		Next
		Set objStaticTexts = Nothing
		
'		If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_ProcessTemplateDropDownButton")=False Then
'			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on process Template drop down button from assign workflow dialog","","","","","")
'			Call Fn_ExitTest()
'		End IF
		objNewProcessDialog.JavaEdit("jedt_ProcessName").Click 1,1,"LEFT"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		If bFlag=False then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ] workflow template does not exist\available for selected object","","","","DONOTSYNC","")	
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ] workflow template is exist\available for selected object","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sButtonName<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_" & sButtonName) = False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButtonName) & " ] from assign workflow dialog","","","","","")
				Call Fn_ExitTest()
			End if
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "VerifyDefaultSelectedTemplate"
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		If Trim(dictWorkflowProcessInfo("ProcessTemplate"))=Trim(Fn_UI_JavaEdit_Operations("RAC_Common_WorkflowProcessOperations", "gettext", objNewProcessDialog, "jedt_ProcessTemplate","")) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ] workflow template is selected selected by default on [ New Process Dialog ]","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplate")) & " ] workflow template is not selected selected by default on [ New Process Dialog ]","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sButtonName<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_" & sButtonName) = False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButtonName) & " ] from assign workflow dialog","","","","","")
				Call Fn_ExitTest()
			End if
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "VerifyWorkflowTemplateAvailable"
		'Set Process Template from Drop down list
		If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_ProcessTemplateDropDownButton")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on process Template drop down button from assign workflow dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		aMulNode=Split(dictWorkflowProcessInfo("ProcessTemplateName"),"~")
		
		'selecting process template
		Set  objStaticTexts = Fn_UI_Object_GetChildObjects("RAC_Common_WorkflowProcessOperations", objNewProcessDialog, "Class Name", "JavaStaticText")
		For iCount = 0 to Ubound(aMulNode)
			bFlag=False
			For iCounter = 0 to objStaticTexts.count-1
				If  objStaticTexts(iCounter).getROProperty("label") = aMulNode(iCount) Then
					Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
					If objStaticTexts(iCounter).getROProperty("abs_x")<>"" or objStaticTexts(iCounter).getROProperty("abs_y")<>"" Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aMulNode(iCount)) & " ] workflow template is exist\available for selected object","","","","DONOTSYNC","")	
						bFlag=True
						Exit for
					End If				
				End If
			Next
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aMulNode(iCount)) & " ] workflow template does not exist\available for selected object","","","","","")
				Call Fn_ExitTest()
			End If
		Next	
		
		objStaticTexts(0).Click 1,1
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		Set  objStaticTexts =Nothing
		
		If sButtonName<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_" & sButtonName) = False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButtonName) & " ] from assign workflow dialog","","","","","")
				Call Fn_ExitTest()
			End if
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "VerifyWorkflowTemplateCount"
		'Set Process Template from Drop down list
		If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_ProcessTemplateDropDownButton")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on process Template drop down button from assign workflow dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		Set objDescription=Description.Create
		objDescription("toolkit class").Value="com.teamcenter.rac.util.combobox.iComboBox.*"
		objDescription("toolkit class").RegularExpression=True
		Set objChild=objNewProcessDialog.ChildObjects(objDescription)
		Set objComboBox=objChild(0)
		Set objChild=Nothing
		Set objDescription=Nothing
		
		iCount=0
		'selecting process template
		Set  objStaticTexts = Fn_UI_Object_GetChildObjects("RAC_Common_WorkflowProcessOperations", objComboBox, "Class Name", "JavaStaticText")
		For iCounter = 0 to objStaticTexts.count-1
			If objStaticTexts(iCounter).getROProperty("abs_x")<>"" or objStaticTexts(iCounter).getROProperty("abs_y")<>"" Then
				iCount=iCount+1
			End If
		Next
				
		If Cint(dictWorkflowProcessInfo("ProcessTemplateCount"))=Cint(iCount) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified current workflow template count [ " & Cstr(iCount) & " ] match with expected workflow template count [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplateCount")) & " ] for selected object","","","","DONOTSYNC","")	
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as current workflow template count [ " & Cstr(iCount) & " ] does not match with expected workflow template count [ " & Cstr(dictWorkflowProcessInfo("ProcessTemplateCount")) & " ] for selected object","","","","","")
			Call Fn_ExitTest()
		End If
		
		objStaticTexts(0).Click 1,1
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		Set  objStaticTexts =Nothing
		Set objComboBox=Nothing
		
		If sButtonName<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_WorkflowProcessOperations", "Click", objNewProcessDialog,"jbtn_" & sButtonName) = False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButtonName) & " ] from assign workflow dialog","","","","","")
				Call Fn_ExitTest()
			End if
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Workflow operation Fail due to invalid case\action name [ " & Cstr(sAction) & " ]","","","","","")
		Call Fn_ExitTest()
End Select			

If Err.Number<0 then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on new process dialog due to Error [" & Cstr(Err.Description) & "]","","","","","")
	Set objNewProcessDialog =Nothing
	Call Fn_ExitTest()
End If

'Releasing new process dialog object
Set objNewProcessDialog = Nothing

Function Fn_ExitTest()
	'Releasing new process dialog object
	Set objNewProcessDialog =Nothing
	ExitTest
End Function

