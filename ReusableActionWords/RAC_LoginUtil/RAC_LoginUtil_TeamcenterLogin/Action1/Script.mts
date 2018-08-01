'! @Name 			RAC_LoginUtil_TeamcenterLogin
'! @Details 		To login to teamcenter application
'! @InputParam1 	sUserName 	: User name to be set
'! @InputParam2 	sPassWord 	: Password
'! @InputParam3 	sGroup 		: Group value
'! @InputParam4 	sRole 		: Role to be set
'! @InputParam5 	sServer 	: Server name
'! @Author 			Kundan Kudale kundan.kudale@sqs.com
'! @Reviewer 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Date 			03 Dec 2015
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_TeamcenterLogin","RAC_LoginUtil_TeamcenterLogin",OneIteration,"infodba","infodba", "", "","TcWeb1"

Option Explicit
Err.Clear

'Declaring variables
Dim sUserName,sPassword, sGroup,sRole,sServer
Dim objTeamcenterLogin,objDefaultWindow,objSSOLogin
Dim iCounter
Dim bFlag
Dim sLoginType

'Get action parameter values in local variables
sUserName = Parameter("sUserName")
sPassword = Parameter("sPassword")
sGroup = Parameter("sGroup")
sRole = Parameter("sRole")
sServer = Parameter("sServer")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_LoginUtil_TeamcenterLogin"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Creating object of [ Teamcenter Login ] window
Set objTeamcenterLogin = Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_TeamcenterLogin","")
'Creating object of [ Teamcenter Default ] window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_DefaultWindow","")
'Creating object of [ Teamcenter SSO Login ] page
Set objSSOLogin=Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","wbpge_SSOLogin","")

'Check existence of Teamcenter login window
sLoginType=""

For iCounter=1 to 60
	If Fn_WEB_UI_WebObject_Operations("RAC_LoginUtil_PLMLauncher","Exist", objSSOLogin,1,"","") Then
		sLoginType="SSO"
		GBL_TC_LOGINTYPE="SSO"
		Exit For
	ElseIf Fn_UI_Object_Operations("RAC_LoginUtil_TeamcenterLogin","Exist", objTeamcenterLogin.JavaButton("jbtn_Login"),0,"", "") Then
		sLoginType="Non-SSO"
		GBL_TC_LOGINTYPE="Non-SSO"
		Exit For
	End If
Next

If sLoginType="" Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as teamcenter login window\page does not exist","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Login to Teamcenter application","","","")

If sLoginType="Non-SSO" Then

'Click on "Clear" button
If Fn_UI_JavaButton_Operations("RAC_LoginUtil_TeamcenterLogin", "Click", objTeamcenterLogin, "jbtn_Clear")	 = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as fail to click on [ clear ] button of teamcenter login window","","","","","")
	Call Fn_ExitTest()
End If

'Enter login details
If  sUserName <> "" And sPassWord <> "" Then
	'Set user name
	If Fn_UI_JavaEdit_Operations("RAC_LoginUtil_TeamcenterLogin", "Set",  objTeamcenterLogin, "jedt_UserID", Trim(sUserName) ) = False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter user name on teamcenter login window","","","","","")
		Call Fn_ExitTest()
	End If	
	'Set password
	If Fn_UI_JavaEdit_Operations("RAC_LoginUtil_TeamcenterLogin", "SetSecure",  objTeamcenterLogin, "jedt_Password", Trim(sPassWord) ) = False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter password on teamcenter login window","","","","","")
		Call Fn_ExitTest()
	End If	
	'Set Group
	If sGroup <> "" Then
		If Fn_UI_JavaEdit_Operations("RAC_LoginUtil_TeamcenterLogin", "Set",  objTeamcenterLogin, "jedt_Group", Trim(sGroup) )=False Then	
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter group on teamcenter login window","","","","","")
			Call Fn_ExitTest()
		End If		
	End If	
	'Set role
	If sRole <> "" Then
		If Fn_UI_JavaEdit_Operations("RAC_LoginUtil_TeamcenterLogin", "Set",  objTeamcenterLogin, "jedt_Role", Trim(sRole)) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter role on teamcenter login window","","","","","")
			Call Fn_ExitTest()
		End If	
	End If	
	'Set server name
	If sServer <> "" Then
		If Fn_UI_JavaList_Operations("RAC_LoginUtil_TeamcenterLogin", "Select", objTeamcenterLogin, "jlst_Server", Trim(sServer), "", "")	 = False Then	
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set Server name on teamcenter login window","","","","","")
			Call Fn_ExitTest()
		End If	
	End If	
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as username and password are invalid","","","","","")
	Call Fn_ExitTest()
End If

'Click on "Login" button
If Cint(Fn_UI_Object_Operations("RAC_LoginUtil_TeamcenterLogin","getroproperty",objTeamcenterLogin.JavaButton("jbtn_Login"),"","enabled","")) = 1 Then
	If Fn_UI_JavaButton_Operations("RAC_LoginUtil_TeamcenterLogin", "Click", objTeamcenterLogin, "jbtn_Login")	= False Then	
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as fail to click on [ Login ] button of teamcenter login window","","","","","")
		Call Fn_ExitTest()
	End If
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as fail to click on [ Login ] button of teamcenter login window","","","","","")
	Call Fn_ExitTest()
End If
ElseIf sLoginType="SSO" Then
	'Set user name
	If Fn_Web_UI_WebEdit_Operations("RAC_LoginUtil_TeamcenterLogin", "Set", objSSOLogin, "wbedt_UserID", Trim(sUserName)) = False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter user name on teamcenter SSO login page","","","","","")
		Call Fn_ExitTest()
	End If	
	
	'Set password
	If Fn_Web_UI_WebEdit_Operations("RAC_LoginUtil_TeamcenterLogin", "SetSecure", objSSOLogin, "wbedt_Password", Trim(sPassword)) = False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter password on teamcenter SSO login page","","","","","")
		Call Fn_ExitTest()
	End If
	
	'Click on "Login" button	
	If Fn_WEB_UI_WebButton_Operations("RAC_LoginUtil_TeamcenterLogin", "Click", objSSOLogin, "wbbtn_Login","","","")	= False Then		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as fail to click on [ Login ] button of teamcenter SSO login page","","","","","")
		Call Fn_ExitTest()
	End If
	
	For iCounter = 1 to 10
		If Browser("wbbrw_TeamcenterSecurityAgent").Page("wbpge_TeamcenterSecurityAgent").JavaWindow("jwnd_PluginEmbeddedFrame").JavaDialog("jdlg_SecurityWarning").Exist(3) Then
			Browser("wbbrw_TeamcenterSecurityAgent").Page("wbpge_TeamcenterSecurityAgent").JavaWindow("jwnd_PluginEmbeddedFrame").JavaDialog("jdlg_SecurityWarning").JavaCheckBox("jckb_AcceptRisk").Set "ON"
			wait 1
			Browser("wbbrw_TeamcenterSecurityAgent").Page("wbpge_TeamcenterSecurityAgent").JavaWindow("jwnd_PluginEmbeddedFrame").JavaDialog("jdlg_SecurityWarning").JavaButton("jbtn_Run").Click			
		End If
	Next
	
End If

If GBL_TEAMCENTER_INVOKE_OPTION="catiatoolbarsave" Then
	'Creating object of [ Teamcenter Save Manager ] window
	Set objDefaultWindow = Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_TeamcenterSaveManager","")
	
	'Verify existence of Teamcenter Save Manager
	bFlag = False
	For iCounter = 1 to 30
		If Fn_UI_Object_Operations("RAC_LoginUtil_TeamcenterLogin","Exist", objDefaultWindow, "","","") Then
			bFlag = True
			Exit For
		Else
			wait(GBL_MIN_MICRO_TIMEOUT)
		End If
	Next
	
	'Validating error
	If Err.Number < 0 Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to Teamcenter Save Manager due to error [ " & Cstr(Err.Description) & " ]","","","","","")
		Call Fn_ExitTest()
	End If
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Login to Teamcenter Save Manager application","","","")
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully logged into Teamcenter Save Manager with user [ " & Cstr(sUserName) & " ]","","","","DONOTSYNC","")
	
	Set objDefaultWindow = Nothing
	ExitAction
End If

'Verify existence of Teamcenter Default window
bFlag = False
For iCounter = 1 to 30
	If Fn_UI_Object_Operations("RAC_LoginUtil_TeamcenterLogin","Exist", objDefaultWindow, "","","") Then
		bFlag = True
		Exit For
	Else
		wait(GBL_MIN_MICRO_TIMEOUT)
	End If
Next


If bFlag Then

	If Dialog("dlg_NonProjectAdministrator").Exist(3) Then
		Dialog("dlg_NonProjectAdministrator").WinButton("wbtn_OK").Click	
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
	End If
	
	'maximizing teamcenter default window
	objDefaultWindow.Maximize
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
	'Validating error
	If Err.Number < 0 Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter due to error [ " & Cstr(Err.Description) & " ]","","","","","")
		Call Fn_ExitTest()
	End If
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Login to Teamcenter application","","","")
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully logged into Teamcenter with user [ " & Cstr(sUserName) & " ]","","","","DONOTSYNC","")	
Else			
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to Teamcenter with user [ " & Cstr(sUserName) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Relasing all objects
Set objTeamcenterLogin=Nothing
Set objDefaultWindow=Nothing
	
Function Fn_ExitTest()
	'Relasing all objects
	Set objTeamcenterLogin=Nothing
	Set objDefaultWindow=Nothing
	ExitTest
End Function

