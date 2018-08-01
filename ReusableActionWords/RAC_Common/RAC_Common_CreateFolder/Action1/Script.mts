'! @Name 			RAC_Common_CreateFolder
'! @Details 		This actionword is used to create a folder
'! @InputParam1 	sFolderType		: Folder type
'! @InputParam2		sFolderName		: Folder name
'! @InputParam3		sDescription 	: Folder description
'! @InputParam4		sOpenOnCreate 	: Folder Open on create option
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			22 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateFolder","RAC_Common_CreateFolder",OneIteration,"Folder","AutomatedTest","Test folder created",""

Option Explicit
Err.Clear

'Declaring variables
Dim sFolderType,sFolderName,sDescription,sOpenOnCreate
Dim sCurrentCharacter,sCurrentNodeName
Dim iItemsCount,iCounter
Dim objNewFolder
Dim bFlag
Dim sTempFolderType,sTempFolderName

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sFolderType=Parameter("sFolderType")
sFolderName=Parameter("sFolderName")
sDescription=Parameter("sDescription")
sOpenOnCreate=Parameter("sOpenOnCreate")

'Creating object of [ New Folder ] Window
Set objNewFolder=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jwnd_NewFolder","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreateFolder"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Checking existance of [ New Folder ] dialog
If Fn_UI_Object_Operations("RAC_Common_CreateFolder","Exist",objNewFolder ,GBL_MIN_MICRO_TIMEOUT,"","")=False Then
	'Invoke new folder dialog by selecting "File->New->Folder..." menu
	LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileNewFolder"
	
	GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreateFolder"
	GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"
	
	'Checking existance of [ New Folder ] dialog	
	If Fn_UI_Object_Operations("RAC_Common_CreateFolder","Exist",objNewFolder ,"","","")=False Then	
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder of type [ " & Cstr(sFolderType) & " ] as [ New Folder ] dialog is not exist","","","","","")
		Call Fn_ExitTest()
	End If
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Folder Create","","Folder Name",sFolderName)

sTempFolderType=sFolderType

If sFolderType <> "" and sFolderName <> "" Then
	'Checking existance of [ New Folder ] dialog	
	If Fn_UI_Object_Operations("RAC_Common_CreateFolder","Exist",objNewFolder ,"","","") Then	
		'Select Folder type
		iItemsCount=Fn_UI_Object_Operations("RAC_Common_CreateFolder","getroproperty",objNewFolder.JavaTree("jtree_FolderType"),"","items count","")
		
		For iCounter=0 To iItemsCount-1
			sCurrentNodeName = objNewFolder.JavaTree("jtree_FolderType").GetItem(iCounter)
			If Trim(sCurrentNodeName)="Most Recently Used~" & Trim(sFolderType) Then
				sFolderType = "Most Recently Used~" & Trim(sFolderType)
				bFlag=True
				Exit For
			ElseIf Trim(sCurrentNodeName)="Complete List~" & Trim(sFolderType) Then
				sFolderType = "Complete List~" & Trim(sFolderType)
				bFlag=True
				Exit For
			End If
		Next
			
		If bFlag = True Then
			If Fn_UI_JavaTree_Operations("RAC_Common_CreateFolder","Select",objNewFolder,"jtree_FolderType",sFolderType,"","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder as fail to select folder type [ " & Cstr(sFolderType) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder as folder type [ " & Cstr(sFolderType) & " ] does not exist in folder type tree","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		sFolderType=sTempFolderType
		
		'Clicking on [ Next ] button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateFolder", "Click", objNewFolder,"jbtn_Next") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder as fail to click on [ Next ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Setting folder name
		If Fn_UI_JavaEdit_Operations("FolderCreate","Set",objNewFolder,"jedt_FolderName",sFolderName ) = False Then
			objNewFolder.JavaEdit("jedt_FolderName").Object.setText sFolderName
		End If
		
		'If value is not set after loop execution then set value in edit box by using Object.setText property
		If Trim(objNewFolder.JavaEdit("jedt_FolderName").GetROProperty("value")) <> Trim(sFolderName) Then		
			For iCounter = 1 to Len(sFolderName)
				sCurrentCharacter = Mid(sFolderName,iCounter, 1)
				If Asc(sCurrentCharacter) = 95 Then
					objNewFolder.JavaEdit("jedt_FolderName").PressKey "_", micShift
				Else
					objNewFolder.JavaEdit("jedt_FolderName").Type Chr(Asc(sCurrentCharacter))
				End If
			Next	
		End If						
		
		'Validating correct folder name is enter in folder name edit box	
		If Trim(objNewFolder.JavaEdit("jedt_FolderName").GetROProperty("value")) <> Trim(sFolderName) Then	
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder as fail to set folder name [ " & Cstr(sFolderName) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Setting folder Description
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateFolder", "Type",  objNewFolder, "jedt_Description", sDescription )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder as fail to set folder description [ " & Cstr(sDescription) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Setting open on create option
		If sOpenOnCreate<>"" Then
			If Fn_UI_JavaCheckBox_Operations("RAC_Common_CreateFolder", "Set",objNewFolder, "jchb_OpenOnCreate", sOpenOnCreate)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder as fail to set open on create option to [ " & Cstr(sOpenOnCreate) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			If Fn_UI_JavaCheckBox_Operations("RAC_Common_CreateFolder","Set", objNewFolder, "jchb_OpenOnCreate", "OFF")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder as fail to set open on create option to [ OFF ]","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on Finish button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateFolder", "Click",objNewFolder,"jbtn_Finish") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

		'Click on Cancel button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateFolder", "Click",objNewFolder,"jbtn_Cancel") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder as fail to click on [ Cancel ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Validating error while creating folder
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder of type [ " & Cstr(sFolderType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
	Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder of type [ " & Cstr(sFolderType) & " ] as [ New Folder ] dialog does not exist","","","","","")
		Call Fn_ExitTest()
	End If
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create folder : folder type & folder name parameter cannot be blank","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Folder Create","","Folder Name",sFolderName)

sTempFolderName=Mid(Environment.Value("TestName"),1,31)
sTempFolderName=Replace(Environment.Value("TestName")," ","")
sTempFolderName=Replace(Environment.Value("TestName"),"_","")

If sFolderName<>"AutomatedTest" and Instr(1,LCase(sTempFolderName),Lcase(Environment.Value("TestName")))=0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created folder [ " & Cstr(sFolderName) & " ] of type [ " & Cstr(sFolderType) & " ]","","","","","")
End If

'Releasing folder dialog object
Set objNewFolder = Nothing 

Function Fn_ExitTest()
	'Releasing folder dialog object
	Set objNewFolder = Nothing 
	ExitTest
End Function


