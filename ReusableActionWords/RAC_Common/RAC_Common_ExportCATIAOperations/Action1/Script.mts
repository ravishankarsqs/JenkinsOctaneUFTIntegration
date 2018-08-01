'! @Name 			RAC_Common_ExportCATIAOperations
'! @Details 		This actionword is used to perform operations Select dataset to Export in catia dialog
'! @InputParam1 	sAction 		: Action Name
'! @InputParam2 	sInvokeOption 	: Dialog invoke option
'! @InputParam3		sDatasetName 	: Dataset name to Export
'! @InputParam4		sColumnName 	: table column name
'! @InputParam5		sValue 			: Column value
'! @InputParam6		sButton			: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			16 Dec 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ExportCATIAOperations","RAC_Common_ExportCATIAOperations", oneIteration, "catiaexportassembly","","","","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sDatasetName,sColumnName,sValue,sButton
Dim objSelectDatasetToExport,objExportProcess,objSelectDatasetToLoad
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


'Creating Object of [ Select Dataset to Load ] dialog
Set objSelectDatasetToLoad=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_SelectDatasetToLoad","")
'Creating Object of [ Export Process ] window
Set objExportProcess =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_ExportProcess","")

'Capture execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Select Dataset To Export in CATIA",sAction,"","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CATIASubMenuOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to login to teamcenter - catia integration and export assembly	
	Case "catiaexportassembly","catiaexportassemblyext"
		LoadAndRunAction "RAC_Common\RAC_Common_CATIASubMenuOperations","RAC_Common_CATIASubMenuOperations", oneIteration, "SelectWithoutSync","","","CATIAExport"
		
		If LCase(sAction)="catiaexportassembly" Then
			'Checking existance of Select Dataset To Load dialog
			bFlag=False
			For iCounter = 0 To 3
				If Fn_UI_Object_Operations("RAC_Common_ExportCATIAOperations","Exist", objSelectDatasetToLoad, GBL_DEFAULT_TIMEOUT,"","") Then
					bFlag=True
					Exit For
				End If	
			Next
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and Load dataset from [ Select Dataset To Load ] dialog as dialog does not exist","","","","","")
				Call Fn_ExitTest()
			End If
			bFlag=False
					
			If sDatasetName="" Then
				sDatasetName="JCI_CV5_GLOBAL_STARTUP_ASSEMBLY_SEATING"
			End If
			
			If Fn_UI_JavaTable_Operations("RAC_Common_ExportCATIAOperations","selectrowext",objSelectDatasetToLoad,"jtbl_DatasetInfo","","Name",sDatasetName,"","") =False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as mentioned dataset not available in table","","","","","")
				Call Fn_ExitTest()
			End If
			
			'Click on OK button
			If Fn_UI_JavaButton_Operations("RAC_Common_ExportCATIAOperations", "Click", objSelectDatasetToLoad,"jbtn_OK") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as failed to click on [ OK ] button","","","","","")
				Call Fn_ExitTest()
			End If
		
		
			bFlag=False
			For iCounter=0 to 180
				'Checking existance of [ Export Process ] dialog
				If Fn_UI_Object_Operations("RAC_Common_ExportCATIAOperations","Exist", objExportProcess, GBL_MIN_TIMEOUT,"","")=False Then
					bFlag=True
					Exit For
				ElseIf Fn_UI_Object_Operations("RAC_Common_ExportCATIAOperations","Exist", objSelectDatasetToLoad,1,"","") Then
					If Fn_UI_JavaTable_Operations("RAC_Common_ExportCATIAOperations","selectrowext",objSelectDatasetToLoad,"jtbl_DatasetInfo","","Name",sDatasetName,"","") =False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as mentioned dataset not available in table","","","","","")
						Call Fn_ExitTest()
					End If
					
					'Click on OK button
					If Fn_UI_JavaButton_Operations("RAC_Common_ExportCATIAOperations", "Click", objSelectDatasetToLoad,"jbtn_OK") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as failed to click on [ OK ] button","","","","","")
						Call Fn_ExitTest()
					End If
				Else
					wait GBL_MIN_MICRO_TIMEOUT
				End If
			Next
		Else
			bFlag=False
			For iCounter=0 to 180
				'Checking existance of [ Export Process ] dialog
				If Fn_UI_Object_Operations("RAC_Common_ExportCATIAOperations","Exist", objExportProcess, GBL_MIN_TIMEOUT,"","")=False Then
					bFlag=True
					wait GBL_MIN_MICRO_TIMEOUT
					Exit For		
				Else
					wait GBL_MIN_MICRO_TIMEOUT
				End If
			Next
		End If
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and Load dataset [ " & Cstr(sDatasetName) & " ] from [ Select Dataset To Load ] dialog as Export process took more time than expected","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Select Dataset To Load in CATIA",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select assembly and Load dataset [ " & Cstr(sDatasetName) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected assembly and Export [ " & Cstr(sDatasetName) & " ] to local drive tcictemp export folder","","","","","")
		End If
End Select

'Releasing all objects
Set objSelectDatasetToLoad=Nothing
Set objExportProcess =Nothing

Function Fn_ExitTest()
	'Releasing all objects	
	Set objSelectDatasetToLoad=Nothing
	Set objExportProcess =Nothing
	ExitTest
End Function
