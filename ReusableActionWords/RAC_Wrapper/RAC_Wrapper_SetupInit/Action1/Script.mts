'! @Name 			RAC_Wrapper_SetupInit
'! @Details 		This Action word use to invoke teamcenter application, login to it, set reset perspective and create test case folder in nav tree.
'! @InputParam1 	bLoginToRACFlag 		: Flag to indicate login to teamcenter application
'! @InputParam2	 	bCacheClear 			: Flag to indicate clearing teamcenter cache
'! @InputParam3 	bRelaunch 				: Flag to indicate relaunch teamcenter application
'! @InputParam4 	sLoginDetails 			: Teamcenter login details
'! @InputParam5 	sAutomationID 			: Teamcenter login user automation ID
'! @InputParam6 	sSite 					: Teamcenter site\database name
'! @InputParam7 	sTCType 				: Teamcenter type
'! @InputParam8 	sTeamcenterInvokeOption : Teamcenter application invoke option
'! @InputParam9 	bSetPespectiveFlag 		: Falg to indicate setting perspective
'! @InputParam10 	bResetPespectiveFlag 	: Falg to indicate re setting perspective
'! @InputParam11 	sPerspective 			: Perspective name
'! @InputParam12 	bSetGroupRole 			: Falg to indicate setting group role at login
'! @InputParam13 	bCreateFolderFlag 		: Falg to indicate for creating test case folder 
'! @InputParam14 	sTestCaseFolderName 	: Test case folder name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			05 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Wrapper\RAC_Wrapper_SetupInit", "Action1", OneIteration, True, True, True, "","AUTTestUser1", "DEV6_TC11","","", True, True, "My Teamcenter", True, True,""

Option Explicit
Err.Clear

'Declaring variables
Dim bLoginToRACFlag, bCacheClear,bRelaunch,sLoginDetails,sAutomationID,sSite,sTCType,sTeamcenterInvokeOption
Dim bSetPespectiveFlag,bResetPespectiveFlag,sPerspective,bSetGroupRole,bCreateFolderFlag,sTestCaseFolderName
Dim sUserDetails

'Get action parameter's value in local variables
bLoginToRACFlag = Parameter("bLoginToRACFlag")
bCacheClear = Parameter("bCacheClear")
bRelaunch = Parameter("bRelaunch")
sLoginDetails = Parameter("sLoginDetails")
sAutomationID = Parameter("sAutomationID")
sSite = Parameter("sSite")
sTCType = Parameter("sTCType")
sTeamcenterInvokeOption = Parameter("sTeamcenterInvokeOption")
bSetPespectiveFlag = Parameter("bSetPespectiveFlag")
bResetPespectiveFlag = Parameter("bResetPespectiveFlag")
sPerspective = Parameter("sPerspective")
bSetGroupRole = Parameter("bSetGroupRole")
bCreateFolderFlag = Parameter("bCreateFolderFlag")
sTestCaseFolderName = Parameter("sTestCaseFolderName")

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -Invoke and login to Teamcenter - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
If  Cbool(bLoginToRACFlag) = True Then
	LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ReuseTcSession","RAC_LoginUtil_ReuseTcSession",OneIteration,bCacheClear,bRelaunch,sLoginDetails,sAutomationID,sSite,sTeamcenterInvokeOption,sTCType
End If

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'Set Perspective- - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
If  Cbool(bSetPespectiveFlag) =True Then
	LoadAndRunAction "RAC_Common\RAC_Common_SetPerspective", "RAC_Common_SetPerspective", OneIteration, sPerspective
End If

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'ReSet Perspective- - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
If  Cbool(bResetPespectiveFlag) =True Then
	LoadAndRunAction "RAC_Common\RAC_Common_ResetPerspective", "RAC_Common_ResetPerspective", OneIteration
End If

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - Set Group Role- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
If sLoginDetails="" Then
	sLoginDetails = sAutomationID
End If

If  Cbool(bSetGroupRole) = True Then
	'Call actionword to set group and role
	LoadAndRunAction "RAC_Common\RAC_Common_UserSessionSettingsOperations","RAC_Common_UserSessionSettingsOperations", oneIteration,"modifysession",sLoginDetails,"menu",sPerspective	
Else
	'Call actionword to check if current group and role is as required. If not then set it
	LoadAndRunAction "RAC_Common\RAC_Common_UserSessionSettingsOperations","RAC_Common_UserSessionSettingsOperations", oneIteration,"VerifyCurrentSessionAndModify",sLoginDetails,"menu",sPerspective
End If

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'Create Testcase folder- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
If  Cbool(bCreateFolderFlag) = True Then	
	'Call actionword to create test case folder in nav tree of Teamcenter
	LoadAndRunAction "RAC_Common\RAC_Common_CreateTestCaseFolder","RAC_Common_CreateTestCaseFolder",OneIteration, sTestCaseFolderName, "", ""
	
	'Call actionword to expand Test case folder
	LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations", oneIteration, "Expand", GBL_TESTCASE_FOLDER_PATH, ""
	
	'Call actionword to select Test case folder
	LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations", oneIteration, "Select", GBL_TESTCASE_FOLDER_PATH, ""
End If
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

