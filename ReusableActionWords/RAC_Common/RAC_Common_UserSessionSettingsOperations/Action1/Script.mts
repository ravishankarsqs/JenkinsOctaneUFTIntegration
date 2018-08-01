'! @Name 			RAC_Common_UserSessionSettingsOperations
'! @Details 		This Action word to perform operations on user setting dialog
'! @InputParam1 	sAction 		: Action to be performed
'! @InputParam2 	sAutomationID 	: automation id
'! @InputParam3 	sInvokeOption 	: Invoke Option (menu or nooption)
'! @InputParam4 	sPerspective 	: Perspective name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com 
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			26 Mar 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_UserSessionSettingsOperations","RAC_Common_UserSessionSettingsOperations",OneIteration,"ModifySession","TestUser4","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction, sInvokeOption, sPerspective, sAutomationID
Dim objSettingDialog 
Dim sGroup,sRole,sSessionInfo,sProjectID,sOtherSiteName
Dim iStart,iEnd,iLength
'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters value in local variables
sAction = Parameter("sAction")
sAutomationID =  Parameter("sAutomationID")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")

'Fetch the group and role for user
sGroup = Fn_Setup_GetTestUserDetailsFromExcelOperations("getgroup","",sAutomationID)
sRole =  Fn_Setup_GetTestUserDetailsFromExcelOperations("getrole","",sAutomationID)

'Get [ User Settings ] object from xml file
Select Case lcase(sPerspective)
	Case "myteamcenter","","my teamcenter"
		Set objSettingDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jwnd_UserSettings","")
End Select

If Lcase(sAction)="setproject" Or Lcase(sAction)="setifprojectavailable" Then
	If dictUserSessionInfo("Project")="" Then
		dictUserSessionInfo("Project")=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ProgramValues_PRE","StandardSTDMembersProgramName",""))	
	End If
	If JavaWindow("jwnd_DefaultWindow").JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","label:=.*" & dictUserSessionInfo("Project") & ".*").Exist(0) Then
		ExitAction
	End If
End If

If Lcase(sAction)="getprojectid" or Lcase(sAction)="getowningproject" Then
	sInvokeOption="nooption"
End IF

Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditUserSetting"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_UserSessionSettingsOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

If Lcase(sAction)<>"getprojectid" and Lcase(sAction)<>"getowningproject" and Lcase(sAction)<>"getowningsite" and Lcase(sAction)<>"getremotesite" and Lcase(sAction)<>"getprimarysite" and Lcase(sAction)<>"getdefaultods" Then
	'checking user setting dialogs existence
	If Not Fn_UI_Object_Operations("RAC_Common_UserSessionSettingsOperations","Exist", objSettingDialog, GBL_DEFAULT_TIMEOUT,"","") Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on User Settings dialog as [ User Settings ] dialog does not exist","","","","","")
		Call Fn_ExitTest()
	End If	
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","User Session Settings",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "modifysession"			
		'Select group value
		If sGroup<>"" Then	
			If Fn_UI_JavaList_Operations("RAC_Common_UserSessionSettingsOperations", "Select", objSettingDialog,"jlst_Group",sGroup, "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify user session as fail to select value [ " & Cstr(sGroup) & " ] from Group dropdown on user settings dialog","","","","","")	
				Call Fn_ExitTest()	
			End If
		End If		
		'Select role value
		If sRole<>"" Then	
			If Fn_UI_JavaList_Operations("RAC_Common_UserSessionSettingsOperations", "Select", objSettingDialog,"jlst_Role",sRole, "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify user session as fail to select value [ " & Cstr(sRole) & " ] from Role dropdown on user settings dialog","","","","","")
				Call Fn_ExitTest()	
			End If
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		'Click on OK button
		If Fn_UI_JavaButton_Operations("RAC_Common_UserSessionSettingsOperations", "Click", objSettingDialog,"jbtn_OK") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify user session as fail to click on [ OK ] button of user settings dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","User Session Settings",sAction,"","")		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully modified session Group to [ " & Cstr(sGroup) & " ] and role as [ " & Cstr(sRole) & " ] from user settings dialog","","","","","")	
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully modified session Group to [ " & Cstr(sGroup) & " ] and role as [ " & Cstr(sRole) & " ] from user settings dialog","","","","","")	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "getprojectid"
		'Fetching datatable current selected row
		GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow
		sProjectID=""
		sSessionInfo=JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0").JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","path:=.*RACPerspectiveHeader.*","index:=1").GetROProperty("label")
		sSessionInfo=SPlit(sSessionInfo,"[")
		sProjectID=Split(sSessionInfo(2),"-")(0)
		DataTable.SetCurrentRow 1	
		Call Fn_CommonUtil_DataTableOperations("AddColumn","SessionPojectID","","")
		DataTable.Value("SessionPojectID","Global")= Cstr(trim(sProjectID))
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","User Session Settings",sAction,"","")		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully get project id [ "&  sProjectID & " ] from user settings dialog","","","","","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "getowningproject"
		'Fetching datatable current selected row
		GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow
		sProjectID=""
		sSessionInfo=JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0").JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","path:=.*RACPerspectiveHeader.*","index:=1").GetROProperty("label")
		sSessionInfo=SPlit(sSessionInfo,"[")
		sProjectID=Split(sSessionInfo(2),"]")(0)
		DataTable.SetCurrentRow 1	
		Call Fn_CommonUtil_DataTableOperations("AddColumn","SessionOwningPoject","","")
		DataTable.Value("SessionOwningPoject","Global")= Cstr(trim(sProjectID))
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","User Session Settings",sAction,"","")		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully get owning project name [ "&  sProjectID & " ] from user settings dialog","","","","","")	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "setproject","setifprojectavailable"		
		
		If Lcase(sAction)="setifprojectavailable" Then
			If Fn_UI_JavaList_Operations("RAC_Common_UserSessionSettingsOperations", "Exist", objSettingDialog,"jlst_Project",dictUserSessionInfo("Project"), "", "") = True Then
				If Fn_UI_JavaList_Operations("RAC_Common_UserSessionSettingsOperations", "Select", objSettingDialog,"jlst_Project",dictUserSessionInfo("Project"), "", "") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project [ " & Cstr(dictUserSessionInfo("Project")) & " ] from Project dropdown on user settings dialog","","","","","")	
					Call Fn_ExitTest()	
				End If
			End If	
		Else
			'Select project value
			If Fn_UI_JavaList_Operations("RAC_Common_UserSessionSettingsOperations", "Select", objSettingDialog,"jlst_Project",dictUserSessionInfo("Project"), "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project [ " & Cstr(dictUserSessionInfo("Project")) & " ] from Project dropdown on user settings dialog","","","","","")	
				Call Fn_ExitTest()	
			End If
		End IF
		
		'Click on OK button
		If Fn_UI_JavaButton_Operations("RAC_Common_UserSessionSettingsOperations", "Click", objSettingDialog,"jbtn_OK") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify user session as fail to click on [ OK ] button of user settings dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","User Session Settings",sAction,"","")		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully select project [ " & Cstr(dictUserSessionInfo("Project")) & " ] from user settings dialog","","","","","")

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
       Case "setavailableproject"			
		'Select project value
'		Call Fn_UI_JavaList_Operations("RAC_Common_UserSessionSettingsOperations", "Select", objSettingDialog,"jlst_Project","#0", "", "")
		objSettingDialog.JavaList("jlst_Project").Select "#0"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select project from user setting dialog as there is no project available under Project list","","","","","")
			Call Fn_ExitTest()
		End If
		'Click on OK button
		If Fn_UI_JavaButton_Operations("RAC_Common_UserSessionSettingsOperations", "Click", objSettingDialog,"jbtn_OK") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify user session as fail to click on [ OK ] button of user settings dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "getowningsite"
		'Fetching datatable current selected row
		GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow
		sSessionInfo=JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0").JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","path:=.*RACPerspectiveHeader.*","index:=1").GetROProperty("label")
		iStart=Instr(1,sSessionInfo,"[")
		iStart=iStart+1
		iEnd=Instr(1,sSessionInfo,"]")
		iLength=Cint(iEnd)-Cint(iStart)
		sSessionInfo=Mid(sSessionInfo,iStart,iLength)
		DataTable.SetCurrentRow 1	
		Call Fn_CommonUtil_DataTableOperations("AddColumn","SessionOwningSite","","")
		DataTable.Value("SessionOwningSite","Global")= Cstr(trim(sSessionInfo))
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","User Session Settings",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "getremotesite"
		'Fetching datatable current selected row
		GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow
		sSessionInfo=JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0").JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","path:=.*RACPerspectiveHeader.*","index:=1").GetROProperty("label")
		iStart=Instr(1,sSessionInfo,"[")
		iStart=iStart+1
		iEnd=Instr(1,sSessionInfo,"]")
		iLength=Cint(iEnd)-Cint(iStart)
		sSessionInfo=Mid(sSessionInfo,iStart,iLength)
		DataTable.SetCurrentRow 1	
		Call Fn_CommonUtil_DataTableOperations("AddColumn","SessionRemoteSite","","")
		DataTable.Value("SessionRemoteSite","Global")= Cstr(trim(sSessionInfo))
		
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","User Session Settings",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "getprimarysite"
		'Fetching datatable current selected row
		GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow
		sSessionInfo=JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0").JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","path:=.*RACPerspectiveHeader.*","index:=1").GetROProperty("label")
		iStart=Instr(1,sSessionInfo,"[")
		iStart=iStart+1
		iEnd=Instr(1,sSessionInfo,"]")
		iLength=Cint(iEnd)-Cint(iStart)
		sSessionInfo=Mid(sSessionInfo,iStart,iLength)
		
		DataTable.SetCurrentRow 1	
		Call Fn_CommonUtil_DataTableOperations("AddColumn","SessionRemoteSite","","")
		DataTable.Value("SessionRemoteSite","Global")= Cstr(trim(sSessionInfo))
		
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","User Session Settings",sAction,"","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "getdefaultods"
		'Fetching datatable current selected row
		GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow
		sSessionInfo=JavaWindow("title:=.* - Teamcenter .*","tagname:=.* - Teamcenter .*","resizable:=1","index:=0").JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","path:=.*RACPerspectiveHeader.*","index:=1").GetROProperty("label")
		iStart=Instr(1,sSessionInfo,"[")
		iStart=iStart+1
		iEnd=Instr(1,sSessionInfo,"]")
		iLength=Cint(iEnd)-Cint(iStart)
		sSessionInfo=Mid(sSessionInfo,iStart,iLength)
				
		sOtherSiteName=Environment.Value("TCDefaultODSName")
		
		DataTable.SetCurrentRow 1	
		Call Fn_CommonUtil_DataTableOperations("AddColumn","SessionRemoteSite","","")
		DataTable.Value("SessionRemoteSite","Global")= Cstr(trim(sSessionInfo))
		
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","User Session Settings",sAction,"","")			
End Select

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on user setting dialog due to error number [" & Cstr(Err.Number) & "] and error description  [" & Cstr(Err.Description) & "]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing object of user settings dialog
Set objSettingDialog = Nothing
	
Function Fn_ExitTest()
	'Releasing object of user settings dialog
	Set objSettingDialog = Nothing
	ExitTest
End Function

