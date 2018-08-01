'! @Name 			RAC_Common_BaselineOperations
'! @Details 		Action word to perform operations on Baseline dialog
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam2		sInvokeOption		: Baseline dialog invoke option
'! @InputParam3 	sBaselineTemplate 	: Baseline template name
'! @InputParam4 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			28 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_BaselineOperations","RAC_Common_BaselineOperations",OneIteration,"autobasiccreate","menu","BaselineTemplate_TCDefaultBaselineProcess",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_BaselineOperations","RAC_Common_BaselineOperations",OneIteration,"clickbutton","nooption","","Apply"

Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sInvokeOption,sBaselineTemplate,sButton
Dim sName,sBaselineRevisionID,sItemID,sJobName,sPerspective
Dim iBaselineCount
Dim objBaseline

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sBaselineTemplate = Parameter("sBaselineTemplate")
sButton = Parameter("sButton")

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Invoking [ Baseline... ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsBaseline"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_BaselineOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating object of [ Baseline... ] dialog
Select Case lcase(sPerspective)
	Case "","myteamcenter"
		Set objBaseline = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_Baseline","")
	Case "structuremanager", "structure manager"
		Set objBaseline = Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_Baseline","")
End Select

'Checking existance of Baseline... dialog
If Fn_UI_Object_Operations("RAC_Common_BaselineOperations","Exist", objBaseline, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Baseline... ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting Baseline count
If Lcase(sAction)= "autobasiccreate" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBaselineCount","","")
	iBaselineCount=Fn_CommonUtil_DataTableOperations("GetValue","RACBaselineCount","","")
	If iBaselineCount="" Then
		iBaselineCount=1
	Else
		iBaselineCount=iBaselineCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACBaselineCount",iBaselineCount,"")	
	
	'Get actual baseline teamplate name
	If sBaselineTemplate = "" Then
		sBaselineTemplate=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_BaselineValues_APL","BaselineTemplate_TCDefaultBaselineProcess",""))
	End If

End If

If sBaselineTemplate<>"" Then
	sBaselineTemplate=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_BaselineValues_APL",sBaselineTemplate,""))
	'Selecting baseline template
	If Lcase(sAction)= "autobasiccreate" Then
		If Fn_UI_JavaList_Operations("RAC_Common_BaselineOperations", "Exist", objBaseline,"jlst_BaselineTemplate",sBaselineTemplate, "", "") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select baseline template [ " & Cstr(sBaselineTemplate) & " ] from baseline template list","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	End If
End If

'Capture execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_BaselineOperations",sAction,"","")
		
Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic baseline with standard values
	Case "autobasiccreate", "autobasiccreatewitherror"

		'Select value for baseline template drop down
		If Fn_UI_JavaList_Operations("RAC_PSE_BaselineOperations", "Select", objBaseline, "jlst_BaselineTemplate", sBaselineTemplate, "", "") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Baseline ] dialog as failed to set value of Baseline Template java list as [" & sBaselineTemplate & "]","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
		'Fetch baseline ID
		sItemID = Fn_UI_JavaEdit_Operations("RAC_Common_BaselineOperations","GetText",objBaseline,"jedt_ItemID", "" )
		'Fetch Revision ID
		sBaselineRevisionID = Fn_UI_JavaEdit_Operations("RAC_Common_BaselineOperations","GetText",objBaseline,"jedt_RevisionID","" )
		'Fetch Name
		sName = Fn_UI_JavaEdit_Operations("RAC_Common_BaselineOperations","GetText",objBaseline,"jedt_Name","" )
		'Fetch Job Name
		sJobName = Fn_UI_JavaEdit_Operations("RAC_Common_BaselineOperations","GetText",objBaseline,"jedt_JobName","" )

		'Store nav tree revision node details in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBaselineRevisionNode","","")
		'DataTable.Value("RACBaselineRevisionNode","Global") = sItemID & "/" & sBaselineRevisionID & ";1-" & sName
		DataTable.Value("RACBaselineRevisionNode","Global") = sItemID & "/" & sBaselineRevisionID & "-" & sName
		
		'Store Baseline ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBaselineID","","")
		DataTable.Value("RACBaselineID","Global") = sItemID
		
		'Store Baseline Revision ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBaselineRevisionID","","")
		DataTable.Value("RACBaselineRevisionID","Global") = sBaselineRevisionID
		
		'Store Baseline Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBaselineName","","")
		DataTable.Value("RACBaselineName","Global") = sName
		
		'Store Baseline Job Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBaselineJobName","","")
		DataTable.Value("RACBaselineJobName","Global") = sJobName
		
		'Click on OK button
		If sButton = "" Then
			sButton = "OK"
		End If
		
		objBaseline.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBaseline", "Click", objBaseline, "jbtn_" & sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create Baseline... of baseline template [ " & Cstr(sBaselineTemplate) & " ] as fail to click on [" & sButton & "] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateBaseline",sAction,"","")
		If sAction <> "autobasiccreatewitherror" Then
			If Err.Number<0 then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create Baseline... of baseline template [ " & Cstr(sBaselineTemplate) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Baseline of baseline template [ " & Cstr(sBaselineTemplate) & " ] with Baseline Id [ " & Cstr(Datatable.Value("RACBaselineID", "Global")) & " ] , Baseline Revision Id [ " & Cstr(Datatable.Value("RACBaselineRevisionID", "Global")) & " ] and Baseline name [ " & Cstr(Datatable.Value("RACBaselineName", "Global")) & " ]","","","","","")
			End If	
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic baseline with standard values
	Case "clickbutton"
		objBaseline.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBaseline", "Click", objBaseline, "jbtn_" & sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [" & sButton & "] button on Baseline dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing object of Baseline... dialog
Set objBaseline =Nothing

Function Fn_ExitTest()
	'Releasing object of Baseline... dialog
	Set objBaseline =Nothing
	ExitTest
End Function

