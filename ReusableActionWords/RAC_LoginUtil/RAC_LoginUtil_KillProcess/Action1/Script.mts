'! @Name 		RAC_LoginUtil_KillProcess
'! @Details 	To kill teamcenter related running process on test case failure
'! @InputParam1 sProcessToKill : Name of the process to be killed
'! @Author 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 	Kundan Kudale kundan.kudale@sqs.com
'! @Date 		26 Mar 2016
'! @Version 	1.0
'! @Example 	LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_KillProcess","RAC_LoginUtil_KillProcess",OneIteration,"Teamcenter.exe"

Option Explicit
Err.Clear

'Declaring variables
Dim sProcessToKill
Dim bReturn
Dim aProcessToKill
Dim iWindowCount,iCounter
Dim sComputerName,sWindowTitle
Dim objTeamcenterWindow,objDefaultWindow,objExitDialog,objCommonWinDialog
Dim objWMIService,objProcess,objProcessCollection

Call Fn_Setup_ReporterFilter("DisableAll")

'Get action parameter values in local variables
sProcessToKill = Parameter("sProcessToKill")

'Creating object of [ Teamcenter Default ] window
Set objDefaultWindow =Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_DefaultWindow","")
'Creating object on Exit Teamcenter dialog
Set objExitDialog =Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_Exit","")
'Creating object on Common windows dialog
Set objCommonWinDialog = Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","wdlg_CommonWinDialog","")

'Checking existnce of Teamcenter Default Window
If Fn_UI_Object_Operations("RAC_RACLoginUtil_KillProcess","Exist", objDefaultWindow,GBL_ZERO_TIMEOUT,"","") Then

	'Call actionword to close all open dialogs in teamcenter
	LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_CloseAllDialogs","RAC_LoginUtil_CloseAllDialogs", oneIteration
	
	If Cint( Fn_UI_Object_Operations("RAC_RACLoginUtil_KillProcess","getroproperty",objDefaultWindow,"","enabled",""))=1 Then		
		objDefaultWindow.Close
		
		If objExitDialog.Exist(6) Then
			objExitDialog.JavaButton("jbtn_Yes").Click
		End If
		
		objCommonWinDialog.SetTOProperty "text","Warning"
		If objCommonWinDialog.Exist(GBL_MIN_TIMEOUT) Then
	        Call Fn_WIN_UI_WinButtonOperations("RAC_RACLoginUtil_KillProcess","Click",objCommonWinDialog,"wbtn_No","","","")
		End If
		
		If objDefaultWindow.JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_CheckedOutObjects").Exist(GBL_MIN_TIMEOUT) Then
	        Call Fn_UI_JavaButton_Operations("RAC_RACLoginUtil_KillProcess", "Click", objDefaultWindow.JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_CheckedOutObjects"), "jbtn_Close")
	        For iCounter = 1 To 20
	        	If objDefaultWindow.Exist(0) Then
	        		wait 0,500
	        	Else
	        		Exit For
	        	End If
	        Next
		End If
	Else
		Call Fn_CommonUtil_WindowsApplicationOperations("Terminate","teamcenter.exe")
		wait 10
	End If	
End If

If objDefaultWindow.Exist(0) Then
	Call Fn_CommonUtil_WindowsApplicationOperations("Terminate","teamcenter.exe")
	Call Fn_CommonUtil_WindowsApplicationOperations("Terminate","javaw.exe")
	SystemUtil.CloseProcessByName("Teamcenter.exe")
End If

If Browser("wbbrw_SSOLogin").Page("wbpge_SSOLogin").Exist(GBL_ZERO_TIMEOUT) Then
	Browser("wbbrw_SSOLogin").Page("wbpge_SSOLogin").Close
	wait 2
End If

If Browser("wbbrw_TeamcenterSecurityServices").Exist(GBL_ZERO_TIMEOUT) Then
	Browser("wbbrw_TeamcenterSecurityServices").Close
	wait 2
End If

If Browser("wbbrw_TeamcenterSecurityAgent").Page("wbpge_TeamcenterSecurityAgent").Exist(GBL_ZERO_TIMEOUT) Then
	Call Fn_WEB_UI_WebButton_Operations("RAC_RACLoginUtil_KillProcess","Click",Browser("wbbrw_TeamcenterSecurityAgent").Page("wbpge_TeamcenterSecurityAgent"),"wbbtn_Logout","","","")
	wait 3
End If

SystemUtil.CloseProcessByName "iexplore.exe"

If Window("Class Name:=Window","text:=TAO ImR").Exist(0) Then
	Window("Class Name:=Window","text:=TAO ImR").Close
	wait 1
End If

If JavaWindow("jwnd_TeamcenterLogin").Exist(0) Then
	If JavaWindow("jwnd_TeamcenterLogin").JavaWindow("jwnd_Login").Exist(0) Then
		JavaWindow("jwnd_TeamcenterLogin").JavaWindow("jwnd_Login").JavaButton("jbtn_OK").Click		
	Else
		JavaWindow("jwnd_TeamcenterLogin").Close
	End If
	wait 3
End If

If JavaWindow("jwnd_TeamcenterLogin").Exist(0) Then
	JavaWindow("jwnd_TeamcenterLogin").Close
	wait 2
End If

'Added on 8-Aug-2017 for 17.04 Dev IDE Execution
Call Fn_CommonUtil_WindowsApplicationOperations("Terminate","teamcenter.exe")
Call Fn_CommonUtil_WindowsApplicationOperations("Terminate","javaw.exe")
Call Fn_CommonUtil_WindowsApplicationOperations("Terminate","java.exe")
SystemUtil.CloseProcessByName("Teamcenter.exe")

Call Fn_Setup_ReporterFilter("EnableAll")

Systemutil.CloseProcessByWndTitle "Teamcenter",False
'Closes RCAF Console
Systemutil.CloseProcessByWndTitle "RCAF Console",True

'Clears teamcenter cache files
Call Fn_Setup_ClearRACCache()

'Releasing object of [ Teamcenter Default ] window
Set objDefaultWindow = Nothing
Set objExitDialog = Nothing
Set objCommonWinDialog = Nothing
