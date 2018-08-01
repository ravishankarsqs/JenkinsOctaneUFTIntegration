'! @Name RAC_LoginUtil_TeamcenterExit
'! @Details To exit from teamcenter application
'! @Author Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer Kundan Kudale kundan.kudale@sqs.com
'! @Date 25 Mar 2016
'! @Version 1.0
'! @Example  LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_TeamcenterExit","RAC_LoginUtil_TeamcenterExit",oneIteration

Option Explicit
Err.Clear

'Declaring variables
Dim objDefaultWindow,objExitDialog,objCommonWinDialog

GBL_CURRENT_EXECUTABLE_APP="NA"
GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_LoginUtil_TeamcenterExit"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Creating object on Teamcenter Default Window
Set objDefaultWindow =Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_DefaultWindow","")
'Creating object on Exit Teamcenter dialog
Set objExitDialog =Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_Exit","")
'Creating object on Common windows dialog
Set objCommonWinDialog = Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","wdlg_CommonWinDialog","")

'Checking existance of teamcenter Default Window
If Fn_UI_Object_Operations("Fn_RACLoginUtil_TeamcenterExit","Exist", objDefaultWindow, "","","") Then 
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Teamcenter Exit","","","")
	
	'Select Menu [File -> Exit]	
	LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileExit"
	
	GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_LoginUtil_TeamcenterExit"
	GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

  	 'Click on [Yes] button
	If Fn_UI_Object_Operations("Fn_RACLoginUtil_TeamcenterExit","Exist", objExitDialog,"","", "") Then									
		If Fn_UI_JavaButton_Operations("Fn_RACLoginUtil_TeamcenterExit", "Click", objExitDialog, "jbtn_Yes")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail exit from teamcenter appliaction as failed to click on [ Yes ] button","","","","","")
			Call Fn_ExitTest()
		End If          														   
	Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to exit from teamcenter as exit teamcenter dialog does not exist","","","","","")
		Call Fn_ExitTest()
	End If

	'Checking existance of Warning dialog
	objCommonWinDialog.SetTOProperty "text","Warning"
	If objCommonWinDialog.Exist(GBL_MICRO_TIMEOUT) Then
        If Fn_WIN_UI_WinButtonOperations("RAC_LoginUtil_TeamcenterExit","Click",objCommonWinDialog,"wbtn_No","","","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail exit from teamcenter appliaction as failed to click on [ No ] button of warning dialog","","","","","")
			Call Fn_ExitTest()
		End IF
	End If
	
	If objDefaultWindow.JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_CheckedOutObjects").Exist(GBL_MICRO_TIMEOUT) Then
        If Fn_UI_JavaButton_Operations("Fn_RACLoginUtil_TeamcenterExit", "Click", objDefaultWindow.JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_CheckedOutObjects"), "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail exit from teamcenter appliaction as failed to click on [ Close ] button of [ Checked Out Objects ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
	End If
	
	If Browser("wbbrw_TeamcenterSecurityServices").Exist(GBL_ZERO_TIMEOUT) Then
		Browser("wbbrw_TeamcenterSecurityServices").Close
	End If
	
	'Click on "Logout" button	
	If Browser("wbbrw_TeamcenterSecurityAgent").Page("wbpge_TeamcenterSecurityAgent").Exist(GBL_ZERO_TIMEOUT) Then
		If Fn_WEB_UI_WebButton_Operations("Fn_RACLoginUtil_TeamcenterExit", "Click", Browser("wbbrw_TeamcenterSecurityAgent").Page("wbpge_TeamcenterSecurityAgent"), "wbbtn_Logout","","","")	= False Then		
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to teamcenter as fail to click on [ Logout ] button of teamcenter SSO logout page","","","","","")
			Call Fn_ExitTest()
		End If
	End If
	
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Teamcenter Exit","","","")
	
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully exit from Teamcenter application","","","","DONOTSYNC","")
	
	'Capturing currently running application information	
	If Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")<>"" Then
		IF Instr(1,Environment.Value("CURRENTLY_RUNNING_APPLICATIONS"),"~RAC") Then
			Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")=Replace(Environment.Value("CURRENTLY_RUNNING_APPLICATIONS"),"~RAC","")
		ElseIF Instr(1,Environment.Value("CURRENTLY_RUNNING_APPLICATIONS"),"RAC~") Then
			Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")=Replace(Environment.Value("CURRENTLY_RUNNING_APPLICATIONS"),"RAC~","")	
		ElseIF Instr(1,Environment.Value("CURRENTLY_RUNNING_APPLICATIONS"),"RAC") Then
			Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")=Replace(Environment.Value("CURRENTLY_RUNNING_APPLICATIONS"),"RAC","")	
		End If
	End IF
	
End If
GBL_CURRENT_EXECUTABLE_APP="NA"

'Releasing all required objects
Set objDefaultWindow = Nothing
Set objExitDialog =Nothing
Set objCommonWinDialog =Nothing

Function Fn_ExitTest()
	'Releasing all required objects
	Set objDefaultWindow = Nothing
	Set objExitDialog =Nothing
	Set objCommonWinDialog =Nothing
	ExitTest
End Function

