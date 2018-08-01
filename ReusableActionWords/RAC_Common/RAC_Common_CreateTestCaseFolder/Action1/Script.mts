'! @Name 			RAC_Common_CreateTestCaseFolder
'! @Details 		This actionword is used to create a test case folder
'! @InputParam1 	sTestCaseFolderName			: Test case folder Name
'! @InputParam2 	sTestCaseFolderDescription 	: Test case folder decription
'! @InputParam3 	sOpenOnCreate 				: Open folder on create option
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			22 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateTestCaseFolder","RAC_Common_CreateTestCaseFolder",OneIteration,"","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sTestCaseFolderName,sTestCaseFolderDescription,sOpenOnCreate
Dim aFolderName

'Get action parameter values in local variables
sTestCaseFolderName 		= Parameter("sTestCaseFolderName")
sTestCaseFolderDescription 	= Parameter("sTestCaseFolderDescription")
sOpenOnCreate 				= Parameter("sOpenOnCreate")

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Checking existance of [ Home~AutomatedTest ] folder
LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"Exist",GBL_AUTOMATEDTEST_FOLDER_PATH, ""

DataTable.SetCurrentRow 1


If Lcase(DataTable.Value("ReusableActionWordReturnValue","Global"))= "false" Then
	DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
	
	'Selecting [ Home ] folder
	LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration, "Select", "Home", ""
	
	'Creating [ AutomatedTest ] folder
	LoadAndRunAction "RAC_Common\RAC_Common_CreateFolder","RAC_Common_CreateFolder",OneIteration,"Folder","AutomatedTest","Folder Created","OFF"
End If

'Expanding [ Home~AutomatedTest ] folder
LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration, "Expand",GBL_AUTOMATEDTEST_FOLDER_PATH, ""

'Selecting [ Home~AutomatedTest ] folder
LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration, "Select",GBL_AUTOMATEDTEST_FOLDER_PATH, ""

If sTestCaseFolderDescription="" Then
	sTestCaseFolderDescription="Test Case Folder Created"
End If
		
'Generating folder name
If sTestCaseFolderName="" Then
	aFolderName = Replace(Environment.Value("TestName")," ","")
	aFolderName = Mid(aFolderName,1,31)
	aFolderName = Replace(aFolderName,"_","")
	
	If isNumeric(aFolderName) Then
		sTestCaseFolderName=CStr(aFolderName) & "_" & CStr(Fn_CommonUtil_GenerateRandomNumber(5))
	Else
		sTestCaseFolderName=CStr(aFolderName) & "_" & CStr(Fn_CommonUtil_GenerateRandomNumber(5))
	End If
End If

'Creating test case folder
LoadAndRunAction "RAC_Common\RAC_Common_CreateFolder","RAC_Common_CreateFolder",OneIteration,"Folder",sTestCaseFolderName,sTestCaseFolderDescription,sOpenOnCreate
	
GBL_TESTCASE_FOLDER_PATH = GBL_AUTOMATEDTEST_FOLDER_PATH & "~" & sTestCaseFolderName

'Printing log	
Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created test case folder [ " & Cstr(sTestCaseFolderName) & " ] under [ " & Cstr(GBL_AUTOMATEDTEST_FOLDER_PATH) & " ] folder","","","","DONOTSYNC","")

'Selecting [ Test Case ] folder
LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration, "Select",GBL_TESTCASE_FOLDER_PATH, ""
