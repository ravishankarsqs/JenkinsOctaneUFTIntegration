'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ActionWord Name			:	RAC_MyWorklist_DelegateSignoffOperations

'Module\Functionality Name	:	MyWorklist functionality
'
'Description				:	Action word to perform operations on Delegate Signoff dialog
'
'Input Parameters			:   1.sAction : Action to perform
'								2.sNode: Myworklist tree node
'								3.sMode: ViewerTab/PerformDoTask
'								4.sAutomationID : Automation ID
'								5.sOrganizationUser : Organization tree user path
'								6.sProjectTeam : Project Team details
'								7.sButton : Button name
'
'Output Parameter			: 	NA
'
'Examples					:  
'								LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_DelegateSignoffOperations","RAC_MyWorklist_DelegateSignoffOperations", oneIteration,"DelegateSignoff","","PerformSignoff","TestUser1EngineeringNWCSWSEngineer","Organization~WS~NWCS~Engineering~Engineer~Sandeep Navghane (502426858)","","Close"
'                       
'History			:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Sandeep Navghane		 	| 09-Oct-2015	|	1.0			|	Sandeep Navghane 			| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

'Declaring variables
Dim sAction,sNode,sMode,sAutomationID,sOrganizationUser,sProjectTeam,sButton
Dim sMenuPath,sMenu,sGroup,sRole,sUserGroupRole,sUserName,sUserID,sUser
Dim objPerformSignoffDialog,objDelegateSignoffDialog
Dim iCounter,iCount
Dim bFlag
Dim aItem

sAction = Parameter("sAction")
sNode = Parameter("sNode")
sMode = Parameter("sMode")
sAutomationID = Parameter("sAutomationID")
sOrganizationUser = Parameter("sOrganizationUser")
sProjectTeam = Parameter("sProjectTeam")
sButton = Parameter("sButton")

sUserGroupRole=""
If sAutomationID<>"" Then
	sUserName = Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",sAutomationID)
	sUserID= Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",sAutomationID)
	sGroup = Fn_Setup_GetTestUserDetailsFromExcelOperations("getgroup","",sAutomationID)
	sRole = Fn_Setup_GetTestUserDetailsFromExcelOperations("getrole","",sAutomationID)
	sUserGroupRole = sUserName & " (" & sUserID & ")-" & sGroup & "/" & sRole
End If

GBL_CURRENT_EXECUTABLE_APP="RAC"

'to select specific node from [ My Worklist ] tree
If sNode<>"" Then
	LoadAndRunAction "RAC_MyWorklist\RAC_MyWorklist_TreeNodeOperations","Action1",OneIteration,"Select",sNode,"",""
	Call Fn_ReadyStatusSync(GBL_MAX_SYNC_ITERATIONS)
End If

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_DelegateSignoffOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Selecting mode
Select Case sMode
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "ViewerTab"
		Set objPerformSignoffDialog=Fn_Setup_GetObjectFromXML("RAC_MyWorklist", "MyWorkListApplet")
		'Selecting [ Viewer ] tab
		LoadAndRunAction "RAC_Common\RAC_TabFolderWidgetOperations","Action1",OneIteration,"Select", "Viewer", ""	
		Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "SignOffButtonClick"
			'Creating object of [ Delegate Signoff ] dialog
			Set objDelegateSignoffDialog=Fn_Setup_GetObjectFromXML("RAC_MyWorklist", "SignoffDecision")
			If Fn_UI_JavaButton_Operations("RAC_MyWorklist_DelegateSignoffOperations","Click",objDelegateSignoffDialog,sButton)=False Then
				Set objPerofrmTask=Nothing
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on button [ " & Cstr(sButton) & " ] while performing perform Delegate Signoff operation","","","","","")
				Call Fn_ExitTest()
			End If	
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			ExitAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "PerformSignoff"
		Set objPerformSignoffDialog=Fn_Setup_GetObjectFromXML("RAC_MyWorklist", "PerformSignoff")
		If Not objPerformSignoffDialog.Exist(6) Then
			'Calling menu
			LoadAndRunAction "RAC_Common\RAC_MenuOperations","Action1",OneIteration,"Select","ActionsPerform","RAC_Menu"	
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
End Select	

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyWorklist_DelegateSignoffOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

Call Fn_Setup_CaptureFunctionExecutionTime("CaptureStartTime","RAC_MyWorklist_DelegateSignoffOperations",sAction,"","")
'Creating object of [ Delegate Signoff ] dialog
Set objDelegateSignoffDialog=Fn_Setup_GetObjectFromXML("RAC_MyWorklist", "DelegateSignoff")

Select Case sAction
	Case "DelegateSignoff"
		'Seleting decision link
		If sUserGroupRole="" Then
			iCounter=0
		Else
			bFlag=False
			For iCounter = 0 To Cint(objPerformSignoffDialog.JavaTable("SignoffTable").GetROProperty("rows"))-1
				If trim(sUserGroupRole)=Trim(objPerformSignoffDialog.JavaTable("SignoffTable").GetCellData(iCounter,"User-Group/Role")) Then
					bFlag=True
					Exit for				
				End If			
			Next
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to get value [ " & Cstr(sUserGroupRole) & " ] from sign off table while performing perform do task operation","","","","","")
				Set objDelegateSignoffDialog=Nothing
				Set objPerformSignoffDialog=Nothing
				Call Fn_ExitTest()
			End If
		End If
		objPerformSignoffDialog.JavaTable("SignoffTable").ClickCell Cstr(iCounter),"User-Group/Role"
		Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		'Checking existance of [ Delegate Signoff ] dialog
		If objDelegateSignoffDialog.Exist(20) Then
			aItem = Split(sOrganizationUser, "~")
			sUser = aItem(0)
			For iCounter = 0 to Ubound(aItem)
				If iCounter > 0 Then
					sUser = sUser & "~" & aItem(iCounter)
				End If
				
				For iCount = 0 to objDelegateSignoffDialog.JavaTree("UserForSelectionTree").GetROProperty("items count")-1
					If objDelegateSignoffDialog.JavaTree("UserForSelectionTree").GetItem(iCount) = sUser Then								
						objDelegateSignoffDialog.JavaTree("UserForSelectionTree").Expand sUser									
						Call Fn_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
						Exit For
					End If								
				Next
			Next 
			objDelegateSignoffDialog.JavaTree("UserForSelectionTree").Select sOrganizationUser
			Call Fn_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)		
			
			If Fn_UI_JavaButton_Operations("RAC_MyWorklist_DelegateSignoffOperations","Click",objDelegateSignoffDialog,"OK")=False THen
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ OK ] button while performing delegate signoff operation","","","","","")
				Set objDelegateSignoffDialog=Nothing
				Set objPerformSignoffDialog=Nothing
				Call Fn_ExitTest()
			End If
			Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			If sButton<>"" Then
				If Fn_UI_JavaButton_Operations("RAC_MyWorklist_DelegateSignoffOperations","Click",objPerformSignoffDialog,sButton)=False THen
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button while performing delegate signoff operation","","","","","")
					Set objDelegateSignoffDialog=Nothing
					Set objPerformSignoffDialog=Nothing
					Call Fn_ExitTest()
				End If
				Call Fn_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to performing delegate signoff operation as delegate signoff dialog does not exist","","","","","")
			Set objDelegateSignoffDialog=Nothing
			Set objPerformSignoffDialog=Nothing
			Call Fn_ExitTest()
		End If
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Perform Delegate Signoff operation due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Set objDelegateSignoffDialog=Nothing
			Set objPerformSignoffDialog=Nothing
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Performed Delegate Signoff operation from [" & Cstr(sMode) & "]","","","","","")	
		End If
		Call Fn_Setup_CaptureFunctionExecutionTime("CaptureEndTime","RAC_MyWorklist_DelegateSignoffOperations",sAction,"","")	
End Select

'Releasing objects
Set objPerformSignoffDialog=Nothing
Set objDelegateSignoffDialog=Nothing

Function Fn_ExitTest()
	Set objPerformSignoffDialog=Nothing
	Set objDelegateSignoffDialog=Nothing
	ExitTest
End Function	

