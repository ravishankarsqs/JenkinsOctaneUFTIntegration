'! @Name 			FMW_Setup_TestcaseExit
'! @Details 		To exit the test case with closing all open applications
'! @InputParam1 	sAppName 	: Application names
'! @InputParam2 	bTCKill 	: Application exit flag
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			28 Mar 2016
'! @Version 		1.0
'! @Example  		LoadAndRunAction "FMW_Setup\FMW_Setup_TestcaseExit","FMW_Setup_TestcaseExit",oneIteration,"RAC~NX","True~True"
Option Explicit
Err.Clear
On Error Resume Next

'Declaring variables
Dim sAppName,bTcKill
Dim TotalTestExecutionTime,TestExecutionMins,TestExecutionSec
Dim objCurentRun,objFSO
Dim aAppName
Dim iCounter

'Get action parameter values in local variables
sAppName = Parameter("sAppName")
bTcKill = Parameter("bTCKill")

Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")=""
GBL_TESTCASE_EXIT_FLAG=True

If Instr(1,Lcase(sAppName),"catia") and Instr(1,Lcase(sAppName),"rac") Then
	If Cbool(bTcKill) Then
		'Calling menu [ File : Close ]
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations", oneIteration,"Select","FileClose",""
		'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_KillProcess","RAC_LoginUtil_KillProcess",OneIteration,""
		'LoadAndRunAction "CATIA_LoginUtil\CATIA_LoginUtil_KillProcess","CATIA_LoginUtil_KillProcess",OneIteration,""
		'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ExitPLMLauncher","RAC_LoginUtil_ExitPLMLauncher",oneIteration
	End If
ElseIf Instr(1,Lcase(sAppName),"rac") Then
	If Cbool(bTcKill) Then
		'Calling menu [ File : Close ]
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations", oneIteration,"Select","FileClose",""
		'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_KillProcess","RAC_LoginUtil_KillProcess",OneIteration,""
		'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ExitPLMLauncher","RAC_LoginUtil_ExitPLMLauncher",oneIteration
	End If
ElseIf Instr(1,Lcase(sAppName),"catia") Then
	If Cbool(bTcKill) Then
		'LoadAndRunAction "CATIA_LoginUtil\CATIA_LoginUtil_KillProcess","CATIA_LoginUtil_KillProcess",OneIteration,""
		'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ExitPLMLauncher","RAC_LoginUtil_ExitPLMLauncher",oneIteration
	End If
End If
		
'aAppName=Split(sAppName,"~")
'For iCounter=0 to Ubound(aAppName)
'	Select Case LCase(aAppName(iCounter))
'		Case "rac"
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			''Kill Teamcenter if Flag is True
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			If Cbool(bTcKill) Then
'				'Calling menu [ File : Close ]
'				LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations", oneIteration,"Select","FileClose",""
'
'				'Kill the application
'				LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_KillProcess","RAC_LoginUtil_KillProcess",OneIteration,""
'			End If
'		Case "nx"
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			''Kill NX if Flag is True
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			If Cbool(bTcKill) Then
'				'Kill the application
'				LoadAndRunAction "NX_LoginUtil\NX_LoginUtil_KillProcess","NX_LoginUtil_KillProcess",OneIteration,""
'			End If
'		Case "catia"
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			''Kill catia if Flag is True
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			If Cbool(bTcKill) Then
'				'Kill the application
'				LoadAndRunAction "CATIA_LoginUtil\CATIA_LoginUtil_KillProcess","CATIA_LoginUtil_KillProcess",OneIteration,""
'			End If					
'		Case "awc"
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			''Kill awc if Flag is True
'			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			If Cbool(bTcKill) Then
'				'Kill the application
'				LoadAndRunAction "AWC_LoginUtil\AWC_LoginUtil_KillProcess","AWC_LoginUtil_KillProcess",OneIteration,""
'			End If		
'	End Select
'Next
'' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
'If Instr(1,GBL_APPLICATIONS_OPENED_IN_TEST,"PLMLAUNCHER") Then
'	LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ExitPLMLauncher", "RAC_LoginUtil_ExitPLMLauncher",oneIteration
'End If

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Calculating total execution time
GBL_TEST_EXECUTION_END_TIME=Now()		 
TotalTestExecutionTime=datediff("s",GBL_TEST_EXECUTION_START_TIME,GBL_TEST_EXECUTION_END_TIME)
TestExecutionMins=TotalTestExecutionTime/60
TestExecutionMins=Split(Cstr(TestExecutionMins),".")
TestExecutionSec=TotalTestExecutionTime-Cint(TestExecutionMins(0))*60
TestExecutionMins(0)=Cint(TestExecutionMins(0))
TestExecutionSec=Cint(TestExecutionSec)
GBL_TEST_EXECUTION_TOTAL_TIME=Cstr(TestExecutionMins(0)) + " Min " + Cstr(TestExecutionSec) + " Sec"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
If LCase(Environment.Value("StepLog"))="true" Then
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Log Test Result in log file
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Call Fn_LogUtil_PrintAndUpdateScriptLog("updatelog","[" & Cstr(time) & "] - QTP [" & Environment.Value("ActionName") & "] - End","","","","","")
		Call Fn_LogUtil_PrintAndUpdateScriptLog("updatelog","-----------------------------------------------------------------------------------------------","","","","","")
		Call Fn_LogUtil_PrintAndUpdateScriptLog("updatelog","[" & Cstr(time) & "] - Final - Pass | Test Execution Result: PASS","","","","","")
		Call Fn_LogUtil_PrintAndUpdateScriptLog("updatelog","[" & Cstr(time) & "] - Final - Pass | Test Execution Total Time : [ " & Cstr(TestExecutionMins(0)) & " ] Minute [ " & Cstr(TestExecutionSec) & " ] Seconds","","","","","")
		Call Fn_LogUtil_PrintAndUpdateScriptLog("updatelog","-----------------------------------------------------------------------------------------------","","","","","")				
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Upload Test log in QC
		If LCase(Environment.Value("UploadLog"))="true" Then
			Call Fn_LogUtil_UploadAttachmentInQCOperations("Attach",Environment.Value("TestLogFile"))
		End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
If LCase(Environment.Value("QCStepLog"))="true" Then
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Log Test Result in QC
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Call Fn_LogUtil_UpdateQCStepLog("Test Execution End","Passed","Test Case [ "+Environment.Value("TestName")+" ] Execution End","","Final - PASS | Test Execution Total Time : [ " + Cstr(TestExecutionMins(0)) + " ] Minute [ " + Cstr(TestExecutionSec) + " ] Seconds")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Set final pass status of Test case
    Set objCurentRun = QCUtil.CurrentRun
	objCurentRun.Status = "Passed"
	objCurentRun.Post
	objCurentRun.Refresh
	Set objCurentRun =Nothing
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
End If
Reporter.ReportEvent micPass,"Final - PASS","Test Case [ " & Environment.Value("TestName") & " ] Execution End"
Call Fn_LogUtil_UpdateTestCaseBatchExecutionResult("","","","PASS","All VP Pass","")

If GBL_SYSLOG_IMAGE_PATH<>"" Then
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.DeleteFile GBL_SYSLOG_IMAGE_PATH,True
	Set objFSO = Nothing
End If

Call Fn_ExitTest()
Function Fn_ExitTest()
	ExitTest
End Function

