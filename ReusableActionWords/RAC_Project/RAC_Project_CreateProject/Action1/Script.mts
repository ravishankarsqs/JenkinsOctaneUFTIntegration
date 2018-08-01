'! @Name 			RAC_Project_CreateProject
'! @Details 		Action word to perform operations on New Project creation window
'! @InputParam1 	sAction 					: String to indicate what action is to be performed
'! @InputParam2 	sID	 						: Project ID
'! @InputParam3 	sName						: Project Name
'! @InputParam4 	sDescription			 	: Project Description
'! @InputParam5 	sCollaborationCategories	: Project Collaboration Categories
'! @InputParam6 	sStatus						: Project status
'! @InputParam7 	sUseProgramSecurity			: Use Program Security option
'! @InputParam8 	sMembersForSelection		: Members for selection to project
'! @InputParam9 	sButton 					: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Shrikant Narkhede shrikant.narkhede@sqs.com
'! @Date 			31 Oct 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Project\RAC_Project_CreateProject","RAC_Project_CreateProject",OneIteration,"123456","Test Project","","","","","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sID,sName,sDescription,sCollaborationCategories,sStatus,sUseProgramSecurity,sMembersForSelection,sButton
Dim iProjectCount,iRandomNumber
Dim objProjectWindow

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter value
sAction = Parameter("sAction")
sID = Parameter("sID")
sName = Parameter("sName")
sDescription = Parameter("sDescription")
sCollaborationCategories = Parameter("sCollaborationCategories")
sStatus = Parameter("sStatus")
sUseProgramSecurity = Parameter("sUseProgramSecurity")
sMembersForSelection = Parameter("sMembersForSelection")
sButton = Parameter("sButton")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Project_CreateProject"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'creating object of [ Project Window ]
Set objProjectWindow = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Project_OR","jwnd_ProjectDefaultWindow","")

'Setting Item count
If Lcase(sAction)= "create" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACProjectCount","","")
	iProjectCount=Fn_CommonUtil_DataTableOperations("GetValue","RACProjectCount","","")
	If iProjectCount="" Then
		iProjectCount=1
	Else
		iProjectCount=iProjectCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACProjectCount",iProjectCount,"")
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Project_CreateProject",sAction,"","")

Select Case LCase(sAction)
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to create new Project
	Case "create"
		If sID="Assign" Or sID="" Then
			iRandomNumber=Fn_CommonUtil_GenerateRandomNumber(6)
			sID=Cstr(iRandomNumber)
		End If
		
		'Set Project ID		
		Call Fn_UI_Object_Operations("RAC_Project_CreateProject","SetTOProperty", objProjectWindow.JavaStaticText("jstx_ProjectLabel"),"","label","ID")
		If Fn_UI_JavaEdit_Operations("RAC_Project_CreateProject", "Set",  objProjectWindow, "jedt_ProjectEdit", sID )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new project as fail to set project id vlaue","","","","","")
			Call Fn_ExitTest()
		End If							
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		If sName="Assign" Or sName="" Then
			iRandomNumber=Fn_CommonUtil_GenerateRandomNumber(6)
			sName="AUT_" & Cstr(iRandomNumber)
		End If

		'Set Project Name		
		Call Fn_UI_Object_Operations("RAC_Project_CreateProject","SetTOProperty", objProjectWindow.JavaStaticText("jstx_ProjectLabel"),"","label","Name")
		If Fn_UI_JavaEdit_Operations("RAC_Project_CreateProject", "Set",  objProjectWindow, "jedt_ProjectEdit", sName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new project as fail to set project name vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		If sDescription="Assign" Then
			sDescription=sName
		End If
		
		If sDescription<>"" Then
			'Set Project Description		
			Call Fn_UI_Object_Operations("RAC_Project_CreateProject","SetTOProperty", objProjectWindow.JavaStaticText("jstx_ProjectLabel"),"","label","Description")
			If Fn_UI_JavaEdit_Operations("RAC_Project_CreateProject", "Set",  objProjectWindow, "jedt_ProjectEdit", sDescription )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new project as fail to set project description vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		If sMembersForSelection<>"" Then
			'Add member to project
			LoadAndRunAction "RAC_Project\RAC_Project_DefinitionOperations", "RAC_Project_DefinitionOperations", oneIteration, "AddMemberToProject",sMembersForSelection & ":projectteamadministrator", ""
		End if
		
		If Fn_UI_JavaButton_Operations("RAC_Project_CreateProject", "Click", objProjectWindow,"jbtn_Create")=False Then
		   Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new project as fail to click on [ Create ] button of new project creation window","","","","","")
		   Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACProjectNode","","")
		DataTable.Value("RACProjectNode","Global") = sID & "-" & sName
		
		'Store project ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACProjectID","","")
		DataTable.Value("RACProjectID","Global") = sID
		
		'Store project RACProjectName
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACProjectName","","")
		DataTable.Value("RACProjectName","Global") = sName
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Project_CreateProject",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new project of due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created project with project Id [ " & Cstr(Datatable.Value("RACProjectID", "Global")) & " ] and project name [ " & Cstr(Datatable.Value("RACProjectName", "Global")) & " ]","","","","","")
		End If
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER		
End Select

'Releasing object of new project dialog
Set objProjectWindow =Nothing

Function Fn_ExitTest()
	'Releasing object of new project dialog
	Set objProjectWindow =Nothing
	ExitTest
End Function
