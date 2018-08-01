'! @Name 			RAC_Common_SelectDatasetToLoadOperations
'! @Details 		This actionword is used to perform operations Select dataset to load in catia dialog
'! @InputParam1 	sAction 		: Action Name
'! @InputParam2 	sInvokeOption 	: Dialog invoke option
'! @InputParam3		sDatasetName 	: Dataset name to load
'! @InputParam4		sColumnName 	: table column name
'! @InputParam5		sValue 			: Column value
'! @InputParam6		sButton			: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			16 Dec 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_SelectDatasetToLoadOperations","RAC_Common_SelectDatasetToLoadOperations", oneIteration, "logintointegrationandloaddataset","","","","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sDatasetName,sColumnName,sValue,sButton
Dim objSelectDatasetToLoad,objLoadProcess
Dim iCounter
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="CATIA"

'Get action input parameters in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sDatasetName = Parameter("sDatasetName")
sColumnName = Parameter("sColumnName")
sValue = Parameter("sValue")
sButton = Parameter("sButton")


'Creating Object of [ Select Dataset to load ] dialog
Set objSelectDatasetToLoad=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_SelectDatasetToLoad","")
'Creating Object of [ Load Process ] window
Set objLoadProcess =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_LoadProcess","")

'Capture execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Select Dataset To Load in CATIA",sAction,"","")

Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to login to teamcenter - catia integration and select dataset to load	
	Case "logintointegrationandloadengineereddrawingdataset","loadengineereddrawingdataset", "logintointegrationengineereddrawing"
		LoadAndRunAction "RAC_Common\RAC_Common_CATIASubMenuOperations","RAC_Common_CATIASubMenuOperations", oneIteration, "SelectWithoutSync","","","CATIALoadinCATIA"
		
		If LCase(sAction)<>"loadengineereddrawingdataset" Then
			If GBL_TC_LOGINTYPE<>"SSO" Then
				If GBL_CATIA_TEAMCENTER_INTEGRATION_LOGIN_FLAG=False Then
					LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ReuseTcSession","RAC_LoginUtil_ReuseTcSession",OneIteration,True,True,"",GBL_TEAMCENTER_LAST_LOGGEDIN_USERID,"","racloadincatiamenuracless",""
				End If
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		'Checking existance of Select Dataset To Load dialog
		If Lcase(sAction) <> "logintointegrationengineereddrawing" Then
			bFlag=False
			For iCounter = 0 To 3
				If Fn_UI_Object_Operations("RAC_Common_SelectDatasetToLoadOperations","Exist", objSelectDatasetToLoad, GBL_DEFAULT_TIMEOUT,"","") Then
					bFlag=True
					Exit For
				End If	
			Next
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset from [ Select Dataset To Load ] dialog as dialog does not exist","","","","","")
				Call Fn_ExitTest()
			End If
			bFlag=False
					
			If sDatasetName="" Then
				sDatasetName="INT_A0_TCPLM_SEATING"
			End If
			
			If Fn_UI_JavaTable_Operations("RAC_Common_SelectDatasetToLoadOperations","selectrowext",objSelectDatasetToLoad,"jtbl_DatasetInfo","","Name",sDatasetName,"","") =False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as mentioned dataset not available in table","","","","","")
				Call Fn_ExitTest()
			End If
			
			'Click on OK button
			If Fn_UI_JavaButton_Operations("RAC_Common_SelectDatasetToLoadOperations", "Click", objSelectDatasetToLoad,"jbtn_OK") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as failed to click on [ OK ] button","","","","","")
				Call Fn_ExitTest()
			End If
			
			bFlag=False
			For iCounter=0 to 199
				'Checking existance of [ Load Process ] dialog
				If Fn_UI_Object_Operations("RAC_Common_SelectDatasetToLoadOperations","Exist", objLoadProcess, GBL_MIN_TIMEOUT,"","")=False Then
					bFlag=True
					Exit For
				Else
					wait GBL_MICRO_TIMEOUT
				End If
			Next
			
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as load process took more time than expected","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Select Dataset To Load in CATIA",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully select and loaded dataset [ " & Cstr(sDatasetName) & " ] in CATIA","","","","","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to login to teamcenter - catia integration and select dataset to load	
	Case "logintointegrationandloaddataset","loaddataset"
		LoadAndRunAction "RAC_Common\RAC_Common_CATIASubMenuOperations","RAC_Common_CATIASubMenuOperations", oneIteration, "SelectWithoutSync","","","CATIALoadinCATIA"		
		
		If LCase(sAction)<>"loaddataset" Then
			If GBL_TC_LOGINTYPE<>"SSO" Then
				If GBL_CATIA_TEAMCENTER_INTEGRATION_LOGIN_FLAG=False Then
					LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ReuseTcSession","RAC_LoginUtil_ReuseTcSession",OneIteration,True,True,"",GBL_TEAMCENTER_LAST_LOGGEDIN_USERID,"","racloadincatiamenuracless",""
				End If
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)				
		End If		
		
'		If sDatasetName<>"" Then
'			'Checking existance of Select Dataset To Load dialog
			bFlag=False
			For iCounter = 0 To 2
				If Fn_UI_Object_Operations("RAC_Common_SelectDatasetToLoadOperations","Exist", objSelectDatasetToLoad, GBL_DEFAULT_TIMEOUT,"","") Then
					bFlag=True
					Exit For
				End If	
			Next
'			If bFlag=False Then
'				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset from [ Select Dataset To Load ] dialog as dialog does not exist","","","","","")
'				Call Fn_ExitTest()
'			End If
'			bFlag=False
'					
			If bFlag=True Then
				If sDatasetName="" Then
					sDatasetName="JCI_CV5_GLOBAL_START_PART"
				End If
					
				If Fn_UI_JavaTable_Operations("RAC_Common_SelectDatasetToLoadOperations","selectrowext",objSelectDatasetToLoad,"jtbl_DatasetInfo","","Name",sDatasetName,"","") =False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as mentioned dataset not available in table","","","","","")
					Call Fn_ExitTest()
				End If
				
				'Click on OK button
				If Fn_UI_JavaButton_Operations("RAC_Common_SelectDatasetToLoadOperations", "Click", objSelectDatasetToLoad,"jbtn_OK") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as failed to click on [ OK ] button","","","","","")
					Call Fn_ExitTest()
				End If
			End If
'		End If
		
		bFlag=False
		For iCounter=0 to 199
			'Checking existance of [ Load Process ] dialog
			If Fn_UI_Object_Operations("RAC_Common_SelectDatasetToLoadOperations","Exist", objLoadProcess, GBL_MIN_TIMEOUT,"","")=False Then
				bFlag=True
				Exit For
			Else
				wait GBL_MICRO_TIMEOUT
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as load process took more time than expected","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Select Dataset To Load in CATIA",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully select and loaded dataset [ " & Cstr(sDatasetName) & " ] in CATIA","","","","","")
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to load in Catia
	Case "loadincatia","loadincatiaext"
		LoadAndRunAction "RAC_Common\RAC_Common_CATIASubMenuOperations","RAC_Common_CATIASubMenuOperations", oneIteration, "SelectWithoutSync","","","CATIALoadinCATIA"
		If LCase(sAction)="loadincatiaext" Then
			bFlag=False
			For iCounter=0 to 199
				'Checking existance of [ Load Process ] dialog
				If Fn_UI_Object_Operations("RAC_Common_SelectDatasetToLoadOperations","Exist", objLoadProcess, GBL_MIN_TIMEOUT,"","")=False Then
					bFlag=True
					Exit For
				Else
					wait GBL_MICRO_TIMEOUT
				End If
			Next
			
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load object in CATIA as load process took more time than expected","","","","","")
				Call Fn_ExitTest()
			End If
			
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Select Dataset To Load in CATIA",sAction,"","")
			If Err.Number<0 then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load object in CATIA due to error [ " & Cstr(Err.Description) & " ]","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully loaded selected object in CATIA","","","","","")
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to login to teamcenter - catia integration and select dataset to load	
	Case "logintointegrationandloadassembly"
		LoadAndRunAction "RAC_Common\RAC_Common_CATIASubMenuOperations","RAC_Common_CATIASubMenuOperations", oneIteration, "SelectWithoutSync","","","CATIALoadinCATIA"
		
		If GBL_TC_LOGINTYPE<>"SSO" Then
			If GBL_CATIA_TEAMCENTER_INTEGRATION_LOGIN_FLAG=False Then
				LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ReuseTcSession","RAC_LoginUtil_ReuseTcSession",OneIteration,True,True,"",GBL_TEAMCENTER_LAST_LOGGEDIN_USERID,"","racloadincatiamenuracless",""
			End If
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Checking existance of Select Dataset To Load dialog
		bFlag=False
		For iCounter = 0 To 3
			If Fn_UI_Object_Operations("RAC_Common_SelectDatasetToLoadOperations","Exist", objSelectDatasetToLoad, GBL_DEFAULT_TIMEOUT,"","") Then
				bFlag=True
				Exit For
			End If	
		Next
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset from [ Select Dataset To Load ] dialog as dialog does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		bFlag=False
				
		If sDatasetName="" Then
			sDatasetName="JCI_CV5_GLOBAL_STARTUP_ASSEMBLY_SEATING"
		End If
		
		If Fn_UI_JavaTable_Operations("RAC_Common_SelectDatasetToLoadOperations","selectrowext",objSelectDatasetToLoad,"jtbl_DatasetInfo","","Name",sDatasetName,"","") =False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as mentioned dataset not available in table","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on OK button
		If Fn_UI_JavaButton_Operations("RAC_Common_SelectDatasetToLoadOperations", "Click", objSelectDatasetToLoad,"jbtn_OK") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as failed to click on [ OK ] button","","","","","")
			Call Fn_ExitTest()
		End If
		
		bFlag=False
		For iCounter=0 to 180
			'Checking existance of [ Load Process ] dialog
			If Fn_UI_Object_Operations("RAC_Common_SelectDatasetToLoadOperations","Exist", objLoadProcess, GBL_MIN_TIMEOUT,"","")=False Then
				bFlag=True
				Exit For
			ElseIf Fn_UI_Object_Operations("RAC_Common_SelectDatasetToLoadOperations","Exist", objSelectDatasetToLoad,1,"","") Then
				If Fn_UI_JavaTable_Operations("RAC_Common_SelectDatasetToLoadOperations","selectrowext",objSelectDatasetToLoad,"jtbl_DatasetInfo","","Name",sDatasetName,"","") =False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as mentioned dataset not available in table","","","","","")
					Call Fn_ExitTest()
				End If
				
				'Click on OK button
				If Fn_UI_JavaButton_Operations("RAC_Common_SelectDatasetToLoadOperations", "Click", objSelectDatasetToLoad,"jbtn_OK") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as failed to click on [ OK ] button","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				wait GBL_MIN_MICRO_TIMEOUT
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as load process took more time than expected","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Select Dataset To Load in CATIA",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select assembly and load dataset [ " & Cstr(sDatasetName) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected assembly and loaded dataset [ " & Cstr(sDatasetName) & " ] in CATIA","","","","","")
		End If
End Select

'Releasing all objects
Set objSelectDatasetToLoad=Nothing
Set objLoadProcess =Nothing

Function Fn_ExitTest()
	'Releasing all objects	
	Set objSelectDatasetToLoad=Nothing
	Set objLoadProcess =Nothing
	ExitTest
End Function

