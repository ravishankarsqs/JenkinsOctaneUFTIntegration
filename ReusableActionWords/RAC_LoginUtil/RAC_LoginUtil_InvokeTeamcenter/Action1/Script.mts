'! @Name RAC_LoginUtil_InvokeTeamCenter
'! @Details To invoke Teamcenter application
'! @InputParam1 ModuleName : Name of the module
'! @InputParam2 InvokeOption :Type of invoke option
'! @InputParam3 TCType : Teamcenter type
'! @Author Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer Kundan Kudale kundan.kudale@sqs.com
'! @Date 25 Mar 2016
'! @Version 1.0
'! @Example LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_InvokeTeamCenter","RAC_LoginUtil_InvokeTeamCenter",oneIteration,"","portalbat",""

Option Explicit
Err.Clear

'Declaring variables
Dim sModuleName,sInvokeOption,sTCType																																	 
Dim sRACAppExecutablePath
Dim objTeamcenterLogin
Dim iCounter
Dim bFlag

'Getting action parameter values in local variables
sInvokeOption = Parameter("sInvokeOption")
sModuleName = Parameter("sModuleName")
sTCType = Parameter("sTCType")

GBL_TEAMCENTER_INVOKE_OPTION=sInvokeOption

'Capturing start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Invoke TeamCenter Login dialog","","Invoke option",sInvokeOption)

Select Case lcase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to invoke teamcenter from portal.bat file
	Case "portalbat"
		'Retrive application executable path
		If lCase(sTCType)="" Then
			sRACAppExecutablePath = Environment.Value("RACAppExecutable")
		End If
		'Setting module name
		If sModuleName="" Then
			sModuleName = "com.teamcenter.rac.gettingstarted.GettingStartedPerspective"
		End If
		'Executing teamcenter path
		SystemUtil.Run sRACAppExecutablePath
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to invoke teamcenter from CATIA Command Start Teamcenter
	Case "catiacommandstartteamcenter"
		LoadAndRunAction "CATIA_Common\CATIA_Common_CATIAStartCommand","CATIA_Common_CATIAStartCommand",OneIteration, "StartFromUtility", "StartTeamcenter", "CATIA_Command"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to invoke teamcenter from CATIA Toolbar's Start Teamcenter option
	Case "catiatoolbarstartteamcenter"
		LoadAndRunAction "CATIA_Common\CATIA_Common_ToolbarOperations","CATIA_Common_ToolbarOperations",OneIteration, "clickext", 1, "StartTeamcenter", "CATIA_Common_TLB"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to invoke teamcenter from CATIA Toolbar's Save option
	Case "catiatoolbarsave"
		LoadAndRunAction "CATIA_Common\CATIA_Common_ToolbarOperations","CATIA_Common_ToolbarOperations",OneIteration, "clickext", 1, "Save", "CATIA_Common_TLB"
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_LoginUtil_InvokeTeamCenter"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sInvokeOption

'Creating object of [ Teamcenter Login ] window
Set objTeamcenterLogin = Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_TeamcenterLogin","")

bFlag=False
For iCounter=0 to 5
	'Checking existance of Teamcenter Login window
	If Fn_UI_Object_Operations("Fn_RACLoginUtil_InvokeTeamCenter","Exist",objTeamcenterLogin,"","","")  Then						  							
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Invoke TeamCenter Login dialog","","Invoke option",sInvokeOption)
		bFlag=True
		Exit For
	End If
Next
'Releasing object of [ Teamcenter Login ] window
Set objTeamcenterLogin=Nothing

If bFlag = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as fail to invoke teamcenter application from location [ " & Cstr(sRACAppExecutablePath) & " ]","","","","","")
	Call Fn_ExitTest()
End If

Function Fn_ExitTest()
	'Releasing object of [ Teamcenter Login ] window
	Set objTeamcenterLogin=Nothing
	ExitTest
End Function

