'! @Name 			RAC_Common_CreateDataset
'! @Details 		Action word to perform operations on new dataset creation dialog. eg. Part basic create , Part detail create
'! @InputParam1 	sAction			: Action Name
'! @InputParam2		sDatasetType	: Dataset type
'! @InputParam3		sInvokeOption	: New dataset creation dialog invoke option
'! @InputParam4		sPerspective	: Perspective name in which user wants to perform operations on New Design dialog
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			18 Jan 2016
'! @Version 		1.0
'! @Example 		dictDatasetInfo("DatasetType")="MSWord"
'! @Example 		dictDatasetInfo("DatasetName")="TestDataset"
'! @Example 		dictDatasetInfo("ImportFile")="C:\mainline\TestData\MultNamedACLs002_2\Dataset.doc"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateDataset","RAC_Common_CreateDataset",OneIteration, "create","MSWord", "menu", "myteamcenter"

Option Explicit
Err.Clear

'Declaring varaibles

'Variable Declaration
Dim sAction,sInvokeOption,sDatasetType
Dim objNewDataset,objDescription,objUploadFile,objImportFile,objChildObjects
Dim sDatasetName,sTestDataPath,sPerspective
Dim iDatasetCount

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sDatasetType = Parameter("sDatasetType")
sInvokeOption = Parameter("sInvokeOption")

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Invoking [ New Dataset ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",oneIteration,"Select","FileNewDataset"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreateDataset"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Creating Object of [ New Dataset ] Dialog
Select Case Lcase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "myteamcenter",""
		Set objNewDataset=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_NewDataset","")
End Select

'Creating Object of [ Upload File ] & [ Import File ] Dialog 
Set objUploadFile=objNewDataset.JavaDialog("jdlg_UploadFile")
Set objImportFile=objNewDataset.JavaDialog("jdlg_ImportFile")
			
'Checking existance of new dataset dialog
If Fn_UI_Object_Operations("RAC_Common_CreateDataset","Exist", objNewDataset,GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ New Dataset ] creation dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting dataset count
If Lcase(sAction)= "create" or Lcase(sAction)="autobasiccreate" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDatasetCount","","")
	iDatasetCount=Fn_CommonUtil_DataTableOperations("GetValue","RACDatasetCount","","")
	If iDatasetCount="" Then
		iDatasetCount=1
	Else
		iDatasetCount=iDatasetCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACDatasetCount",iDatasetCount,"")
	DataTable.SetCurrentRow iDatasetCount
End If

If Lcase(sAction)= "create" or Lcase(sAction)="autobasiccreate" Then
	'Clickng on [ More... ] Option to Choose Dataset Type
	If Fn_UI_JavaCheckBox_Operations("RAC_Common_CreateDataset", "Set", objNewDataset, "jckb_More", "ON")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to click on [ More ] option from new dataset creation dialog","","","","","")
		Call Fn_ExitTest()
	End If
	Call Fn_RAC_ReadyStatusSync(2)
	
	'Set Dataset Type
	Set objChildObjects = Fn_UI_Object_GetChildObjects("RAC_Common_CreateDataset",objNewDataset,"Class Name~label","JavaStaticText~" & sDatasetType)
	objChildObjects(0).Click 1,1
	Set objChildObjects =Nothing
	
	If Err.Number <> 0 Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as specified dataset type does not exist on new dataset creation dialog","","","","","")
		Call Fn_ExitTest()
	End If
	Call Fn_RAC_ReadyStatusSync(1)
End If

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to Create New Dataset	
	Case "create"				
		'Setting Dataset name
		If dictDatasetInfo("DatasetName")<>"" Then
			If dictDatasetInfo("DatasetName")="Assign" Then
				dictDatasetInfo("DatasetName") = Fn_Setup_GenerateObjectInformation("getname",sDatasetType)
			End If
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDataset","Set",objNewDataset,"jedt_Name",dictDatasetInfo("DatasetName")) = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to set dataset name value","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
		End IF
		
		'Setting Dataset Description
		If dictDatasetInfo("Description")<>"" Then
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDataset", "Set",  objNewDataset, "jedt_Description", dictDatasetInfo("Description") ) = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to set dataset description value","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
		End If
		
		'Setting Tool Used to Dataset
		If dictDatasetInfo("ToolUsed")<>"" Then
			If Fn_UI_JavaList_Operations("RAC_Common_CreateDataset", "Select", objNewDataset, "jlst_ToolUsed", dictDatasetInfo("ToolUsed"), "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to select dataset tool used value","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
		End If
		
		'Importing External Dataset File
		If dictDatasetInfo("ImportFile")<>"" Then
			'Clicking on browse button
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateDataset","Click", objNewDataset,"jbtn_Browse") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to click on [ Browse ] button of import file option","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
			
			'get Path of TestData
			sTestDataPath = Fn_Setup_GetAutomationFolderPath("TestData")
			dictDatasetInfo("ImportFile") = sTestDataPath  & "\" & dictDatasetInfo("ImportFile")
			'Set file path to upload external dataset
			If Fn_UI_Object_Operations("RAC_Common_CreateDataset","Exist",objUploadFile,GBL_MIN_TIMEOUT,"","")	Then
				If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDataset","Type",objUploadFile,"jedt_FileName",dictDatasetInfo("ImportFile")) = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to set dataset import file path value","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(1)
				
				'Clicking on upload button
				If Fn_UI_JavaButton_Operations("RAC_Common_CreateDataset", "Click",objUploadFile,"jbtn_Upload")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to click on [ Upload ] button of upload file dialog","","","","","")
					Call Fn_ExitTest()
				End If					
			ElseIf Fn_UI_Object_Operations("RAC_Common_CreateDataset","Exist",objImportFile,GBL_MICRO_TIMEOUT,"","") Then
				If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDataset","Type",objImportFile,"jedt_FileName",dictDatasetInfo("ImportFile")) = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to set dataset import file path value","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(1)

				'Clicking on import button
				If Fn_UI_JavaButton_Operations("RAC_Common_CreateDataset", "Click",objImportFile,"jbtn_Import")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to click on [ Upload ] button of upload file dialog","","","","","")
					Call Fn_ExitTest()
				End If					
			End If
			Call Fn_RAC_ReadyStatusSync(2)
		End If
		
		'Selecting Open On Create Option
		If dictDatasetInfo("OpenOnCreate")<>"" Then
			If Fn_UI_JavaCheckBox_Operations("RAC_Common_CreateDataset", "Set", objNewDataset, "jckb_OpenOnCreate", dictDatasetInfo("OpenOnCreate"))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to set [ Open On Create ] option","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
		End If
		
		'Get dataset name		
		sDatasetName = Cstr(Fn_UI_JavaEdit_Operations("RAC_Common_CreateDataset","GetText",objNewDataset,"jedt_Name","" ))
		
		'Clicking On OK Button to create Dataset
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateDataset", "Click",objNewDataset,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to click on [ OK ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created dataset of type [ " & CStr(sDatasetType) & " ] and name as [ " & Cstr(sDatasetName) & " ]","","","","","")
		
		'Set value of new dataset name in DatasetName column
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDatasetName","","")
		DataTable.SetCurrentRow iDatasetCount		
		DataTable.Value("RACDatasetName","Global") = sDatasetName
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to Create New Dataset	
	Case "autobasiccreate"	
		'Setting Dataset name
		sDatasetName=Fn_Setup_GenerateObjectInformation("getname",sDatasetType)
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDataset","Set",objNewDataset,"jedt_Name",sDatasetName) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to set dataset name value","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(1)
		
		'Setting Dataset Description
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDataset", "Set",  objNewDataset, "jedt_Description",sDatasetName & " Description" ) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to set dataset description value","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(1)
				
		'Get dataset name		
		sDatasetName = Cstr(Fn_UI_JavaEdit_Operations("RAC_Common_CreateDataset","GetText",objNewDataset,"jedt_Name","" ))
		
		'Clicking On OK Button to create Dataset
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateDataset", "Click",objNewDataset,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new dataset of type [ " & CStr(sDatasetType) & " ] as fail to click on [ OK ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created dataset of type [ " & CStr(sDatasetType) & " ] and name as [ " & Cstr(sDatasetName) & " ]","","","","","")
		
		'Set value of new dataset name in DatasetName column
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDatasetName","","")
		DataTable.Value("RACDatasetName","Global") = sDatasetName
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Relasing Object of [ New Dataset ] Dialog
Set objNewDataset=Nothing
Set objImportFile=Nothing
Set objUploadFile=Nothing

Public Function Fn_ExitTest()
	'Relasing Object of [ New Dataset ] Dialog
	Set objNewDataset=Nothing
	Set objImportFile=Nothing
	Set objUploadFile=Nothing
	ExitTest
End Function

