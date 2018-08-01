'! @Name 			RAC_LoginUtil_ReuseTcSession
'! @Details 		To reuse\login to teamcenter session
'! @InputParam1		bCacheClear : Flag to indicate if content of cache is to be cleared
'! @InputParam2 	bRelaunch : Flag to indicate if Teamcenter application is relauched
'! @InputParam3		sLoginDetails : User id and password details
'! @InputParam4 	sAutomationID : Automation ID of user to login with
'! @InputParam5 	sSite : Server\dtabase\site name
'! @InputParam6 	sTeamcenterInvokeOption : Invoke option e.g. portal
'! @InputParam7 	sTCType : Teamcenter type
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			25 Mar 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ReuseTcSession","RAC_LoginUtil_ReuseTcSession",OneIteration,True,True,"","TestUser1","","portalbat",""

Option Explicit
Err.Clear

'Declaring variables
Dim bCacheClear, bRelaunch, sLoginDetails,sAutomationID,sSite,sTeamcenterInvokeOption,sTCType
Dim bReuse
DIm aLogin
Dim sSessionDetails
Dim objAboutTeamcenter,objDefaultWindow
Dim sTemp

'Get parameter values in local variables
bCacheClear = Parameter("bCacheClear")
bRelaunch = Parameter("bRelaunch")
sLoginDetails = Parameter("sLoginDetails")
sAutomationID = Parameter("sAutomationID")
sSite = Parameter("sSite")
sTeamcenterInvokeOption = Parameter("sTeamcenterInvokeOption")
sTCType = Parameter("sTCType")

GBL_TEAMCENTER_LAST_LOGGEDIN_USERID=sAutomationID
GBL_CATIA_TEAMCENTER_INTEGRATION_LOGIN_FLAG=False
GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_LoginUtil_ReuseTcSession"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

If sTeamcenterInvokeOption="catiatoolbarsaveracless" or sTeamcenterInvokeOption="catiatoolbarsave" Then
	'Do nothing
Else
	If Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")="" Then
		Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")="RAC"
	Else
		Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")=Environment.Value("CURRENTLY_RUNNING_APPLICATIONS") & "~RAC"
	End If
	GBL_APPLICATIONS_OPENED_IN_TEST=GBL_APPLICATIONS_OPENED_IN_TEST & "~RAC"
End If

'Creating object of teamcenter [ Default ] Window
Set objDefaultWindow =Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_DefaultWindow","")
'Creating object of About teamcenter Window
Set objAboutTeamcenter =Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_AboutTeamcenter","")

'Retrive login user details
If sLoginDetails="" and sAutomationID<>"" Then
	sTemp=sAutomationID
	sLoginDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sAutomationID)
	sAutomationID=sTemp
End If

'Spliting the details
aLogin  =Split(sLoginDetails,"~",-1,1)

If sTeamcenterInvokeOption="" Then
	sTeamcenterInvokeOption="portalbat"
End If

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'This is temporary implementation to handle catia-teamcenter integration login uplon clicking on Save button
IF sTeamcenterInvokeOption="catiatoolbarsaveracless" Or sTeamcenterInvokeOption="racloadincatiamenuracless" Then
	Dim objTeamcenterLogin
	Dim iCounter
	Dim bFlag
	'Creating object of [ Teamcenter Login ] window
	Set objTeamcenterLogin = Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_TeamcenterLogin","")
	'Checking existance of Teamcenter Login window
	If Fn_UI_Object_Operations("RAC_LoginUtil_ReuseTcSession","Exist",objTeamcenterLogin,GBL_MAX_TIMEOUT,"","")  Then
		'Set password
		If Fn_UI_JavaEdit_Operations("RAC_LoginUtil_ReuseTcSession", "SetSecure",  objTeamcenterLogin, "jedt_Password", Trim(aLogin(1)) ) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter password on CATIA Integration teamcenter login window","","","","","")
			Call Fn_ExitTest()
		End If	
	Else
		ExitAction
	End If
	
	'Click on "Login" button
	If Cint(Fn_UI_Object_Operations("RAC_LoginUtil_ReuseTcSession","getroproperty",objTeamcenterLogin.JavaButton("jbtn_Login"),"","enabled","")) = 1 Then
		If Fn_UI_JavaButton_Operations("RAC_LoginUtil_ReuseTcSession", "Click", objTeamcenterLogin, "jbtn_Login")	= False Then	
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as fail to click on [ Login ] button of teamcenter CATIA Integration login window","","","","","")
			Call Fn_ExitTest()
		End If
	Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as fail to click on [ Login ] button of teamcenter CATIA Integration login window","","","","","")
		Call Fn_ExitTest()
	End If

	bFlag=False
	For iCounter=0 to 59
		'Checking existance of Teamcenter Login window
		If Fn_UI_Object_Operations("RAC_LoginUtil_ReuseTcSession","Exist",objTeamcenterLogin,"","","") Then
			wait GBL_MIN_MICRO_TIMEOUT
		Else
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Invoke TeamCenter Login dialog","","Invoke option",sTeamcenterInvokeOption)
			bFlag=True
			Exit For
		End If
	Next
	'Releasing object of [ Teamcenter Login ] window
	Set objTeamcenterLogin=Nothing

	If bFlag = False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter - CATIA Integration teamcenter","","","","","")
		Call Fn_ExitTest()
	End If
	GBL_CATIA_TEAMCENTER_INTEGRATION_LOGIN_FLAG=True	
	ExitAction
End If
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

If CBool(bRelaunch) =True Then
	LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_KillProcess","RAC_LoginUtil_KillProcess",OneIteration,""
	
	'Set the bReuse flag to False, as bRelaunch is True
	bReuse = False	
	If CBool(bCacheClear) =True Then
		'Clear teamcenter cache
		Call Fn_Setup_ClearRACCache()
	End If
	
	'launching Tc Application from the Path mentioned in EnvironmentVariables.xml
	LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_InvokeTeamCenter","RAC_LoginUtil_InvokeTeamCenter",OneIteration,sTeamcenterInvokeOption,"",sTCType
	
	'Retrive site 
	If sSite="" or sSite="Site1" Then
		sSite=Environment.Value("Site1")
	End If	
	
''		If sSite="" or sSite="Site1" Then
''			sSite=Environment.Value("Site1")
''		ElseIf sSite="Custom" Then	
''		   sSite=GBL_Tc_ENVIRONMENT_NAME_FOR_CUSTOM_LAUNCH
''		End If
	
	'Login To Tc Application with Supplied Data
	If lcase(Environment.Value("IsGroupRoleRequired"))="true" then
		LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_TeamcenterLogin","RAC_LoginUtil_TeamcenterLogin",OneIteration,aLogin(0),aLogin(1),aLogin(2),aLogin(3),sSite
	Else
		LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_TeamcenterLogin","RAC_LoginUtil_TeamcenterLogin",OneIteration,aLogin(0),aLogin(1),"","",""
	End if
	If GBL_TEAMCENTER_INVOKE_OPTION="catiatoolbarsave" Then
		ExitAction
	End If
Else
	bReuse=True
End If

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_LoginUtil_ReuseTcSession"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'If bReuse = True, Check for the Existing Session Details
If bReuse = True Then
	sSessionDetails = ".*" & Lcase(aLogin(0)) & ".*" & aLogin(2) & " / " & aLogin(3) & ".*"
	'Checking Exisstance of teamcenter Default Window
	If Fn_UI_Object_Operations("RAC_LoginUtil_ReuseTcSession","Exist",objDefaultWindow,GBL_ZERO_TIMEOUT,"","") Then
		'Maximizing teamcenter default window
		objDefaultWindow.Maximize
		Call Fn_UI_Object_Operations("RAC_RACLoginUtil_KillProcess","settoproperty",objDefaultWindow.JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","label:=.*IMC.*","index:=1"),"","label", sSessionDetails)
		If Fn_UI_Object_Operations("RAC_LoginUtil_ReuseTcSession","Exist",objDefaultWindow.JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","label:=.*IMC.*","index:=1"),GBL_DEFAULT_TIMEOUT,"","")=False Then
			bReuse = False
		End If
	Else
		bReuse = False
	End if

	If bReuse = False Then
		'launching Tc Application from the Path mentioned in EnvironmentVariables.xml
		LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_InvokeTeamCenter","RAC_LoginUtil_InvokeTeamCenter",OneIteration,sTeamcenterInvokeOption,"",sTCType
		'Login To Tc Application with Supplied Data
		LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_TeamcenterLogin","RAC_LoginUtil_TeamcenterLogin",OneIteration,aLogin(0),aLogin(1),aLogin(2),aLogin(3),sSite
	End If
Else
	sSessionDetails = ".*" & Lcase(aLogin(0)) & ".*" & aLogin(2) & " / " & aLogin(3) & ".*"
	'If Fn_UI_Object_Operations("RAC_LoginUtil_ReuseTcSession","Exist",objDefaultWindow.JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","path:=.*RACPerspectiveHeader.*","logical_location:=X_UNK__Y_SMALL","index:=1","label:=" & sSessionDetails),GBL_ZERO_TIMEOUT,"","")=False Then
	If Fn_UI_Object_Operations("RAC_LoginUtil_ReuseTcSession","Exist",objDefaultWindow.JavaStaticText("toolkit class:=org.eclipse.swt.widgets.Label","label:=" & sSessionDetails),1,"","")=False Then
		LoadAndRunAction "RAC_Common\RAC_Common_UserSessionSettingsOperations","RAC_Common_UserSessionSettingsOperations",OneIteration,"ModifySession",sAutomationID,"",""
	End If
End If
'Invoking Help->About menu
LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"SelectAndDoNotPrintLog","HelpAbout"
Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
If Len(Environment.Value("TestName"))>156 Then
	GBL_SYSLOG_IMAGE_PATH = Environment.Value("BatchFolderName") & "\" & Mid(Environment.Value("TestName"),1,156) & "_SysLog.png"
Else
	GBL_SYSLOG_IMAGE_PATH = Environment.Value("BatchFolderName") & "\" & Environment.Value("TestName") & "_SysLog.png"
End If
objDefaultWindow.CaptureBitmap GBL_SYSLOG_IMAGE_PATH,True

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_LoginUtil_ReuseTcSession"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

If Fn_UI_JavaButton_Operations("RAC_LoginUtil_ReuseTcSession","Click",objAboutTeamcenter,"jbtn_OK")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ OK ] button of About Teamcenter dialog while performing login to teamcenter operation","","","","","")
	Call Fn_ExitTest()
End IF

GBL_CURRENT_EXECUTABLE_APP="RAC"	

'Releasing objects
Set objDefaultWindow = Nothing
Set objAboutTeamcenter = Nothing
	
Function Fn_ExitTest()
	'Releasing objects
	Set objDefaultWindow = Nothing
	Set objAboutTeamcenter = Nothing
	ExitTest
End Function

