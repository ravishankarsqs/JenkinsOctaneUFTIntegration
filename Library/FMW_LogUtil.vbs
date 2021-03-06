Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Function Name																|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. Fn_LogUtil_UpdateDetailLog														|	sandeep.navghane@sqs.com	|	16-Jan-2015	|	Function Used to update detail\technical log in test case log file
'002. Fn_LogUtil_CreateTestCaseLogFile													|	sandeep.navghane@sqs.com	|	16-Jan-2015 |	Function Used to create test case log file
'003. Fn_LogUtil_PrintAndUpdateScriptLog												|   sandeep.navghane@sqs.com	|	11-Feb-2016 |	Function Used to update Test Case Log file, batch excel and mic report
'004. Fn_LogUtil_UpdateQCStepLog														|	sandeep.navghane@sqs.com	|	16-Jan-2015 |	Function used to update step by step test case log in QC
'005. Fn_LogUtil_UploadAttachmentInQCOperations											|	sandeep.navghane@sqs.com	|	16-Jan-2015 |	Function used to upload attachmenta in QC\ALM test set in test lab
'006. Fn_LogUtil_UpdateTestCaseBatchExecutionResult										|	sandeep.navghane@sqs.com	|	16-Jan-2015 |	Function used to update test case result in batch execution detail excel file
'007. Fn_LogUtil_CaptureFunctionExecutionTime											|	sandeep.navghane@sqs.com	|	16-Jan-2015 |	Function Used to calculate function performance time
'008. Fn_LogUtil_PrintStepHeaderLog														|	sandeep.navghane@sqs.com	|	16-Jan-2015 |	Function use to print step header information in test case
'009. Fn_LogUtil_PrintTestCaseHeaderLog													|	sandeep.navghane@sqs.com	|	16-Jan-2015 |	Function use to print step header information in test case
'010. Fn_LogUtil_CreateBatchExecutionResultSummarySheet									|	sandeep.navghane@sqs.com	|	04-Apr-2016 |	Function Used to create batch execution result excel
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start 
'Function Name			 :	Fn_LogUtil_UpdateDetailLog
'
'Function Description	 :	Function Used to update detail\technical log in test case log file
'
'Function Parameters	 :  1.sFilePath	: Test case log file path
'							2.sText		: Text to print in log file
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Test log file should exist
'
'Function Usage		     :  bReturn = Fn_LogUtil_UpdateDetailLog("C:\GOG_Schertz\Reports\PM_1- Able to Create parts.log","[ RAC_FolderCreate ] Fn_UI_JavaButton_Operations >> FAIL : [ Fail to click on OK button ] : OK button is disabled")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Jan-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Replacement for function : Fn_WriteLogFile ' Delete this comment once implementation is completed
Public Function Fn_LogUtil_UpdateDetailLog(sFilePath,sText)
 	'Declaring variables
	Dim objFSO,objFile
	
	'Initially function returns False
	Fn_LogUtil_UpdateDetailLog=False
	
	If LCase(Environment.Value("StepLog"))="true" Then
		'Check detail Log flag
		If CBool(Environment.Value("DetailLog")) = True Then
			'Creating object of File System
			Set objFSO = CreateObject("Scripting.FileSystemObject")	
			If sFilePath = "" Then
				'If file path is blank then function reads default test case log file path
				sFilePath = Environment.Value("TestLogFile")
			End If
			'Open text file
			Set objFile = objFSO.OpenTextFile(sFilePath,8)
			'Update detail\technical log
			objFile.Write sText
			objFile.Write vblf
			'Releasing objects of File System
			Set objFile = Nothing
			Set objFSO = Nothing
			Fn_LogUtil_UpdateDetailLog=True
		End If
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_LogUtil_CreateTestCaseLogFile
'
'Function Description	 :	Function Used to create test case log file
'
'Function Parameters	 :  1.sLogFileName	: Test case log file name
'
'Function Return Value	 : 	Test case log file path or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Valid test case log file name
'
'Function Usage		     :  bReturn = Fn_LogUtil_CreateTestCaseLogFile(Environment.Value("TestName") & ".log")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Jan-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Replacement for function : Fn_CreateLogFile ' Delete this comment once implementation is completed
Public Function Fn_LogUtil_CreateTestCaseLogFile(sLogFileName)
	'Declaring variables
	Dim sFilePath
	Fn_LogUtil_CreateTestCaseLogFile=False
	'Test case log file path
	sFilePath = Environment.Value("BatchFolderName") & "\" & sLogFileName
	If Len(sFilePath )>259 Then
		sLogFileName=Mid(sLogFileName,1,240-Len(Environment.Value("BatchFolderName"))) & "_SHORTEN" & ".log"
		sFilePath = Environment.Value("BatchFolderName") & "\" & sLogFileName
	End If
	'Creting test case log file	
	Fn_LogUtil_CreateTestCaseLogFile=Fn_FSOUtil_FileOperations("createfile",sFilePath,"","")
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_LogUtil_PrintAndUpdateScriptLog
'
'Function Description	:	Function used to update Test Case Log file, batch excel and mic report.
'
'Function Parameters	:   1.sLogType: Type of the log
'						    2.sTestLogComment: Log statements to be entered in test log file
'							3.sStepName: Step name for which log need to be print
'							4.bKillTCSession: Boolean variable option for Teamcenter session kill
'							5.bExitTest: Boolean variable option for Exit test iteration
'							6.iReadystatusIterations: Number of iteration for ready status of app
'							7.sAppName: App Name to kill or Iterate syncronization
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_LogUtil_PrintAndUpdateScriptLog("pass_action","Successfull created item","Item creation","","","","1","RAC")
'Function Usage		     :	Call Fn_LogUtil_PrintAndUpdateScriptLog("pass_verification","Successfull verified item is created","Item creation","","","","1","RAC")
'Function Usage		     :	Call Fn_LogUtil_PrintAndUpdateScriptLog("fail_action","Fail to create item","Item creation","true","true","","","WEB")
'Function Usage		     :	Call Fn_LogUtil_PrintAndUpdateScriptLog("fail_verification","Fail to verify item is created","Item creation","true","true","","","WEB")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  11-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_LogUtil_PrintAndUpdateScriptLog(sLogType, sTestLogComment, sStepName, bKillTCSession, bExitTest, iReadystatusIterations, sAppName)
	'Declaring Variables
	Dim objFSO, objFile
	Dim sFileName, sImagePath, sBatchLogComment
	Dim aAppName, aIterations
	Dim iTotalTestExecutionTime, iTestExecutionMins, iTestExecutionSec
	Dim iCounter
	Dim bFlag
	
	If GBL_DISABLEPRINTANDUPDATELOG_REPORTING=True Then
		Exit Function
	End If
	
	'Setting varibales to initial value
	If lcase(sLogType) <> "step_header" or lcase(sLogType) <> "updatelog" Then
		If sStepName = "" Then
			sStepName = Cstr(GBL_STEP_NUMBER)
		End If
	End If
	sBatchLogComment = ""
	
	'Setting variables to handle FAIL action
	If lcase(sLogType) = "fail_verification" or lcase(sLogType) = "fail_action" Then
		GBL_DISABLEPRINTANDUPDATELOG_REPORTING=True
		If sAppName = "" Then
			If GBL_APP_NAME_TO_EXIT_ON_FAILURE = "" Then
				GBL_APP_NAME_TO_EXIT_ON_FAILURE = Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")
			End If
			sAppName = GBL_APP_NAME_TO_EXIT_ON_FAILURE
		End If
		If bKillTCSession = "" Then
			bKillTCSession = True
		End If
		If bExitTest = "" Then
			bExitTest = True
		End If
	'Setting variables to handle PASS action
	ElseIf lcase(sLogType) = "pass_verification" or lcase(sLogType)="pass_action" Then
		If sAppName = "" Then
			If GBL_APP_NAME_TO_SYNC = "" Then
				GBL_APP_NAME_TO_SYNC = GBL_CURRENT_EXECUTABLE_APP
			End IF
			sAppName = GBL_APP_NAME_TO_SYNC
		End If
		If iReadystatusIterations = "" Then
			If GBL_APP_SYNC_ITERATIONS = "" Then
				GBL_APP_SYNC_ITERATIONS = GBL_MIN_SYNC_ITERATIONS
			End If
			iReadystatusIterations = GBL_APP_SYNC_ITERATIONS
		End If
		If iReadystatusIterations = "DONOTSYNC" Then
			sAppName = ""
		End IF
	End If
	
	On Error Resume Next
	'Creating object of File system
	If lCase(Environment.Value("StepLog")) = "true" Then
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		sFileName = Environment.Value("TestLogFile")
		Set objFile = objFSO.OpenTextFile(sFileName,8)
	End If
	
	'Setting Exceution time variable
	GBL_STEP_EXECUTION_TIME = datediff("s",GBL_LAST_LOG_UPDATION_TIME,Now())
	
	Select Case lcase(sLogType)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to print Step header in Log file
		Case "step_header"
			If lCase(Environment.Value("StepLog")) = "true" Then
				objFile.WriteLine "-----------------------------------------------------------------------------------------------"
				objFile.WriteLine sTestLogComment	
				objFile.WriteLine "-----------------------------------------------------------------------------------------------"
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to update Log file
		Case "updatelog"
			If lCase(Environment.Value("StepLog")) = "true" Then
				objFile.WriteLine sTestLogComment	
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to print pass action in Log file
		Case "pass_action"
			If lCase(Environment.Value("StepLog")) = "true" Then
				objFile.WriteLine Time() & " - " & "Action - PASS [ "& sStepName &" ] | " & sTestLogComment	
			End If
			''Call Fn_LogUtil_UpdateTestScriptStepWiseLogInExcel("",sStepName,sTestLogComment,"PASS",GBL_STEP_EXECUTION_TIME)   ###need to implement
			If lCase(Environment.Value("QCStepLog")) = "true" Then
				Call Fn_LogUtil_UpdateQCStepLog(sStepName,"Passed","","",sTestLogComment)
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to print pass verification in Log file
		Case "pass_verification"
			If lCase(Environment.Value("StepLog")) = "true" Then
				objFile.WriteLine Time() & " - " & "Verify - PASS [ "& sStepName &" ] | " & sTestLogComment	
			End If
			''Call Fn_LogUtil_UpdateTestScriptStepWiseLogInExcel("",sStepName,sTestLogComment,"PASS",GBL_STEP_EXECUTION_TIME)   ###need to implement
			If lCase(Environment.Value("QCStepLog")) = "true" Then
				Call Fn_LogUtil_UpdateQCStepLog(sStepName,"Passed","","",sTestLogComment)
			End If
			GBL_VERIFICATION_COUNTER = GBL_VERIFICATION_COUNTER+1
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to print fail action in Log file
		Case "fail_action"
			If lCase(Environment.Value("StepLog")) = "true" Then
				objFile.WriteLine Time() & " - " & "Action - FAIL [ "& sStepName &" ] | " & sTestLogComment	
				If sBatchLogComment = "" Then
					sBatchLogComment = "FAIL: " & sTestLogComment
				Else
					sBatchLogComment = "FAIL: " & sBatchLogComment
				End If
			End If
			''Call Fn_LogUtil_UpdateTestScriptStepWiseLogInExcel("",sStepName,sTestLogComment,"FAIL",GBL_STEP_EXECUTION_TIME)   ###need to implement
			If lCase(Environment.Value("QCStepLog")) = "true" Then
                Call Fn_LogUtil_UpdateQCStepLog(sStepName,"Failed","","",sTestLogComment)
			Else
            	Reporter.ReportEvent micFail,sStepName,"Test Execution Fail"				
			End If
			Call Fn_CommonUtil_DataTableOperations("ExportDataTable","","","")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to print fail verification in Log file
		Case "fail_verification"
			If lCase(Environment.Value("StepLog")) = "true" Then
				objFile.WriteLine Time() & " - " & "Verify - FAIL [ "& sStepName &" ] | " & sTestLogComment	
				If sBatchLogComment = "" Then
					sBatchLogComment = "FAIL: " & sTestLogComment
				Else
					sBatchLogComment = "FAIL: " & sBatchLogComment
				End If
			End If
			''Call Fn_LogUtil_UpdateTestScriptStepWiseLogInExcel("",sStepName,sTestLogComment,"FAIL",GBL_STEP_EXECUTION_TIME)   ###need to implement
			If lCase(Environment.Value("QCStepLog")) = "true" Then
                Call Fn_LogUtil_UpdateQCStepLog(sStepName,"Failed","","",sTestLogComment)
			Else
            	Reporter.ReportEvent micFail,sStepName,"Test Execution Fail"				
			End If
			Call Fn_CommonUtil_DataTableOperations("ExportDataTable","","","")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_LogUtil_PrintAndUpdateScriptLog ] : Fail to perform operation [ " & Cstr(sLogType) & " ] : No valid case was passed for function [Fn_LogUtil_PrintAndUpdateScriptLog] for Log Type")
	End Select
	
	'Setting Log update variable time to Now
	GBL_LAST_LOG_UPDATION_TIME = Now()
	
	'Uploading failure snapshot and log
	If Lcase(sLogType) = "fail_action" or Lcase(sLogType) = "fail_verification" Then
		'Uploading failure image
		sImagePath = Environment.Value("BatchFolderName") +"\"   & Environment.Value("TestName") & ".png"
		Desktop.CaptureBitmap sImagePath,True
	
		If lCase(Environment.Value("UploadFailImage")) = "true" Then            
			Call Fn_LogUtil_UploadAttachmentInQCOperations("Attach",sImagePath)
			objFile.WriteLine vbLF & Time() & " - " & "Test Script Failure Image Name := [ "& Environment.Value("TestName") &".png ]"
		End If
		
		If lCase(Environment.Value("UploadSysLogImage")) = "true" Then
			If GBL_SYSLOG_IMAGE_PATH<>"" Then
				If objFSO.FileExists(GBL_SYSLOG_IMAGE_PATH) Then
					Call Fn_LogUtil_UploadAttachmentInQCOperations("Attach",GBL_SYSLOG_IMAGE_PATH)
					objFile.WriteLine vbLF & Time() & " - " & "SysLog Image Name := [ "& Environment.Value("TestName") &"_SysLog.png ]"
				End If	
			End If
		End If	
		'uploading Log file
		If lCase(Environment.Value("UploadLog")) = "true" Then
			If lCase(Environment.Value("StepLog")) = "true" Then
                Call Fn_LogUtil_UploadAttachmentInQCOperations("Attach",Environment.Value("TestLogFile"))
			End If
		End If
	End If
	'Release Objects
	Set objFSO = Nothing
	Set objFile = Nothing
	
	'Killing application
	bFlag=False
	If lcase(bKillTCSession) = "true" Then		
		If Instr(1,Lcase(sAppName),"catia") Then
			'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_KillProcess","RAC_LoginUtil_KillProcess",OneIteration,Environment.Value("KillProcesses")
			'LoadAndRunAction "CATIA_LoginUtil\CATIA_LoginUtil_KillProcess","CATIA_LoginUtil_KillProcess",OneIteration,""
			'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ExitPLMLauncher","RAC_LoginUtil_ExitPLMLauncher",oneIteration
		ElseIf Instr(1,Lcase(sAppName),"rac") Then
			'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_KillProcess","RAC_LoginUtil_KillProcess",OneIteration,Environment.Value("KillProcesses")
			'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ExitPLMLauncher","RAC_LoginUtil_ExitPLMLauncher",oneIteration
		ElseIf Instr(1,Lcase(sAppName),"tcra") Then
			LoadAndRunAction "TcRA_LoginUtil\RAC_TcRA_KillProcess","RAC_TcRA_KillProcess",OneIteration,""
		ElseIf Instr(1,Lcase(sAppName),"tpdm") Then
			LoadAndRunAction "TPDM_LoginUtil\TPDM_LoginUtil_KillProcess","TPDM_LoginUtil_KillProcess",OneIteration,""
		ElseIf Instr(1,Lcase(sAppName),"fpdm") Then			
		Else
			'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_KillProcess","RAC_LoginUtil_KillProcess",OneIteration,Environment.Value("KillProcesses")
			'LoadAndRunAction "CATIA_LoginUtil\CATIA_LoginUtil_KillProcess","CATIA_LoginUtil_KillProcess",OneIteration,""
			'LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ExitPLMLauncher","RAC_LoginUtil_ExitPLMLauncher",oneIteration
		End If
		
'		Err.Clear
'		If Instr(1,Lcase(sAppName),"plmlauncher") Then
'			sAppName=Replace(Lcase(sAppName),"plmlauncher","<<skip>>")
'		End If
'		
'		aAppName = Split(sAppName,"~")
'		
'		For iCounter = 0 To ubound(aAppName)
'			Select Case lcase(aAppName(iCounter))
'				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				'Case to kill RAC application
'				Case "rac"
'					LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_KillProcess","RAC_LoginUtil_KillProcess",OneIteration,Environment.Value("KillProcesses")
'				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				'Case to kill awc application
'				Case "awc"
'					LoadAndRunAction "AWC_LoginUtil\AWC_LoginUtil_KillProcess","AWC_LoginUtil_KillProcess",OneIteration,""
'				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				'Case to kill catia application
'				Case "catia"
'					LoadAndRunAction "CATIA_LoginUtil\CATIA_LoginUtil_KillProcess","CATIA_LoginUtil_KillProcess",OneIteration,""
'				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				'Case to ext PLM Launcher application
'				Case "<<skip>>"
'					bFlag=True
'				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				'Case to handle invalid request
'				Case Else
'					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_LogUtil_PrintAndUpdateScriptLog ] : Fail to perform operation [ " & Cstr(sLogType) & " ]: No valid case was passed for function [Fn_LogUtil_PrintAndUpdateScriptLog] for App Name")
'			End Select
'		Next
'		
'		If bFlag=True Then
'			LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ExitPLMLauncher","RAC_LoginUtil_ExitPLMLauncher",oneIteration
'		End If
		
		'Calculating total execution time
		GBL_TEST_EXECUTION_END_TIME = Now()		 
		iTotalTestExecutionTime = datediff("s",GBL_TEST_EXECUTION_START_TIME,GBL_TEST_EXECUTION_END_TIME)
		iTestExecutionMins = iTotalTestExecutionTime/60
		iTestExecutionMins = Split(Cstr(iTestExecutionMins),".")
		iTestExecutionSec = iTotalTestExecutionTime-Cint(iTestExecutionMins(0))*60
		iTestExecutionMins(0) = Cint(iTestExecutionMins(0))
		GBL_TEST_EXECUTION_TOTAL_TIME = Cstr(iTestExecutionMins(0)) & " Min " & Cstr(iTestExecutionSec) & " Sec"
		TestFailImageInsertFlag = True		
		Call Fn_LogUtil_UpdateTestCaseBatchExecutionResult("","","","FAIL",Cstr(sBatchLogComment),"")
	End If

	'Exit from Test Case
	If lcase(bExitTest) = "true" Then
		ExitTest
	End If
	
	'Performing synchronization of application
	aAppName = Split(sAppName,"~")
	aIterations = Split(iReadystatusIterations,"~")
	For iCounter = 0 To ubound(aAppName)
		Select Case lcase(aAppName(iCounter))
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to sync RAC application
			Case "rac"
				Call Fn_RAC_ReadyStatusSync(aIterations(iCounter))
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to sync AWC application
			Case "awc"
				Call Fn_AWC_ReadyStatusSync(aIterations(iCounter))
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to handle invalid request
			Case Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_LogUtil_PrintAndUpdateScriptLog ] : Fail to perform operation [ " & Cstr(sLogType) & " ] : No valid case was passed for function [Fn_LogUtil_PrintAndUpdateScriptLog] to synchronize application")
		End Select
	Next
	If iReadystatusIterations = "DONOTSYNC" Then
		iReadystatusIterations = ""
	End If
	GBL_LOG_ADDITIONAL_INFORMATION = ""
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_LogUtil_UpdateQCStepLog
'
'Function Description	 :	Function used to update step by step test case log in QC
'
'Function Parameters	 :  1.sStep: Step name
'							2.sStatus: Step status
'							3.sDescription: Step Description
'							4.sExpected: Step Expected result
'							5.sActualResult: Step Actual result
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	UFT\QTP should be connected to ALM\QC
'
'Function Usage		     :  bReturn = Fn_LogUtil_UpdateQCStepLog("STEP 01","Passed","Create item","Item should get created","Succssfully created item")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Jan-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_LogUtil_UpdateQCStepLog(sStep,sStatus,sDescription,sExpected,sActualResult)	
	'Declaring variables
	Dim bFlag
	Dim objCurentRun,objCurentTest
	Dim iCounter,iCount,iStep,iStepNumber
	Dim objStepFactory,objDesignStepFactory
	Dim objStepList,objDesignStepList,objAddStep
	Dim sCurrentActualResult,sStepName,sStepName1
	Dim sStepNameType1,sStepNameType2,sStepNameType3,sStepNameType4
	
	sStep=Cstr(sStep)
	sStepNameType1="step" & sStep
	sStepNameType2="step " & sStep
	sStepNameType3=""
	sStepNameType4=""
	If Len(sStep)=1 Then
		sStepNameType3="step0" & sStep
		sStepNameType4="step 0" & sStep
	End If
	'Initially set function return value as False
	Fn_LogUtil_UpdateQCStepLog=False
	
	'Creating object of currently executing Test run
	Set objCurentRun = QCUtil.CurrentRun
	'Creating object of currently executing Test Case
	Set objCurentTest = QCUtil.CurrentTest

	'Creating object of currently executing Test run Step Factory
	Set objStepFactory = objCurentRun.StepFactory
	'Creating object of currently executing Test case Design Step Factory
	Set objDesignStepFactory = objCurentTest.DesignStepFactory

	'Creating object of currently executing Test run Step List
	Set objStepList = objStepFactory.NewList("")
	'Creating object of currently executing Test case Design Step List
	Set objDesignStepList = objDesignStepFactory.NewList("")

	'Checking Description value
	If sDescription="" Then
		For iCount=1 to objDesignStepList.Count
				sStepName=Trim(Cstr(objDesignStepList.Item(iCount).StepName))
			If lCase(sStepName)=LCase(sStep) or lCase(sStepName)=LCase(sStepNameType1) or lCase(sStepName)=LCase(sStepNameType2) or lCase(sStepName)=LCase(sStepNameType3) or lCase(sStepName)=LCase(sStepNameType4) Then		
				sStep=objDesignStepList.Item(iCount).StepName
				'Getting Description
				sDescription=objDesignStepList.Item(iCount).StepDescription
				'Getting Expected result
				sExpected=objDesignStepList.Item(iCount).StepExpectedResult
				Exit for
			End If
		Next
	End If
	'This sets step count
	iStep = objStepList.Count

	'Set flag
	bFlag=False
	'Looping with Test run Step List
	For iCounter=1 to objStepList.Count
		'Match Step name with current step name
		sStepName=Trim(Cstr(objStepList.Item(iCounter).Field("ST_STEP_NAME")))
		If lCase(sStepName)=LCase(sStep) or lCase(sStepName)=LCase(sStepNameType1) or lCase(sStepName)=LCase(sStepNameType2) or lCase(sStepName)=LCase(sStepNameType3) or lCase(sStepName)=LCase(sStepNameType4) Then
			sStep=objStepList.Item(iCounter).Field("ST_STEP_NAME")
			'Getting Actual result of step
			sCurrentActualResult=objStepList.Item(iCounter).Field("ST_ACTUAL")
			'Append result
			sActualResult=sCurrentActualResult+vblf+sActualResult
			iStep =iCounter
			bFlag=True
			Exit for
		End If
	Next

	If bFlag=False Then
		'Adding New step in current run
		'Creating object of Add Item in Current run test case
		Set objAddStep=objStepFactory.AddItem(Null)
		'Adding Step name
		objAddStep.Field("ST_STEP_NAME")=sStep
		'Adding Step status
		objAddStep.Field("ST_STATUS") = sStatus
		'Adding Step description
		objAddStep.Field("ST_DESCRIPTION") = sDescription
		'Adding Step expected result
		objAddStep.Field("ST_EXPECTED") =sExpected
		'Adding Step actual result
		objAddStep.Field("ST_ACTUAL") = sActualResult
		'Post changes in QC
		objAddStep.Post
		'Refresh changes
		objAddStep.Refresh
		Set objAddStep=Nothing
	Else
		'Updating existing step in current run
		'Updating Step status
		objStepList.Item(iStep).Field("ST_STATUS") = sStatus
		'Updating Step description
		objStepList.Item(iStep).Field("ST_DESCRIPTION") = sDescription
		'Updating Step expected result
		objStepList.Item(iStep).Field("ST_EXPECTED") = sExpected
		'Updating Step actual result
		objStepList.Item(iStep).Field("ST_ACTUAL") = sActualResult
		'Post changes in QC
		objStepList.Post
		'Refresh changes
		objStepList.Refresh
	End If

	'If Currently running test case status is fail then mark the test case as Fail
	If lcase(sStatus)="failed" Then
		Reporter.ReportEvent micFail,sStep,"Test Execution Fail"
		objCurentRun.Status = "Failed"
		objCurentRun.Post
		objCurentRun.Refresh
	End If

	'Releasing all objects
	Set objStepList = Nothing
	Set objStepFactory = Nothing
	Set objCurentRun = Nothing
	Set objDesignStepFactory =Nothing
	Set objDesignStepList =Nothing
	
	If Err.Number <> 0 Then
		Fn_LogUtil_UpdateQCStepLog=False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>: [ Fn_LogUtil_UpdateQCStepLog ] : Fail to update QC\ALM step log due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	Else
		Fn_LogUtil_UpdateQCStepLog=True
	End If	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_LogUtil_UploadAttachmentInQCOperations
'
'Function Description	 :	Function used to upload attachmenta in QC\ALM test set in test lab
'
'Function Parameters	 :  1.Action			: Action name to perform
'							2.sAttachmentPath	: Attachment path
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	UFT\QTP should be connected to ALM\QC
'
'Function Usage		     :  bReturn = Fn_LogUtil_UploadAttachmentInQCOperations("Attach","C:\GOG_Schertz\Reports\PM_19- Able to create real BOM.log")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Jan-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_LogUtil_UploadAttachmentInQCOperations(sAction,sAttachmentPath)
 	'Declaring variables
	Dim objCurentTestSet,objAttachments,objAttachmentItem
    
	'Initially set function return value as False
	Fn_LogUtil_UploadAttachmentInQCOperations=False
	If Lcase(Cstr(Environment.Value("QCStepLog")))="false" Then
		Fn_LogUtil_UploadAttachmentInQCOperations=True
		Exit Function
	End If

	'Creating object of Currently running test set
    Set objCurentTestSet = QCUtil.CurrentTestSet	
	'Creating object of Currently running test set Attachments
	Set objAttachments = objCurentTestSet.Attachments	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to upload attachments to QC\ALM test set in test lab
		Case "Attach"
			'Creating object of new Item
			Set objAttachmentItem = objAttachments.AddItem(Null)
			'Uploading attachment to QC
			objAttachmentItem.FileName = sAttachmentPath
			'Setting Attachment type
			objAttachmentItem.Type = 1
			'Post attachment
			objAttachmentItem.Post
			'Refreshing QC client
			objAttachmentItem.Refresh
			'Releasing object of new Item
			Set objAttachmentItem= Nothing
	End Select	
	'Releasing object of Currently running test set Attachments
	Set objAttachments= Nothing
	'Releasing object of Currently running test set
	Set objCurentTestSet = Nothing
	
    If Err.Number <> 0 Then
		Fn_LogUtil_UploadAttachmentInQCOperations=False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_LogUtil_UploadAttachmentInQCOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] log due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	Else
		Fn_LogUtil_UploadAttachmentInQCOperations=True
	End If	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_LogUtil_UpdateTestCaseBatchExecutionResult
'
'Function Description	 :	Function used to update test case result in batch execution detail excel file
'
'Function Parameters	 :  1.sResultSheetLocation : Btach Execution Details file path
'							2.sTestCaseName		   : Test case name to add in sheet
'							3.sDate				   : Execution start date and time
'							4.sResultData		   : Execution Result
'							5.sComments	           : Btach file comments
'							6.iSheetNumber	       : Sheet number
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Batch execution details file should exist
'
'Function Usage		     :  bReturn = Fn_LogUtil_UpdateTestCaseBatchExecutionResult("C:\GOG_Schertz\Reports\BtachExecutionDetails.xlsx","Test1",Cstr(Date) & "-" & cstr(Time),"PASS","All VP Pass","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Jan-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Replacement for function : Fn_Update_TestResult ' Delete this comment once implementation is completed
Public Function Fn_LogUtil_UpdateTestCaseBatchExecutionResult(sResultSheetLocation,sTestCaseName,sDate,sResultData,sComments,iSheetNumber)
	'Declaring variables
	Const xlCellTypeLastCell = 11
	Const xlContinuous = 1
	Const xlCenter = -4108
	Dim iCounter
	Dim iColumnCount,iRowCount
	Dim objFile,objFSO,objExcel,objWorkbook,objWorkSheet,objRange
	
	'Initially set function return value as False
	Fn_LogUtil_UpdateTestCaseBatchExecutionResult = False
	
	'Get batch execution file path
	If sResultSheetLocation = "" Then
		sResultSheetLocation=Environment.Value("BatchFolderName") & "\" & Environment.Value("BatchExecutionFileName")
	End If
	'Checking existance of result file
	If Fn_FSOUtil_FileOperations("fileexist",sResultSheetLocation,"","") Then	
				
		'Creating excel object
		Set objExcel = CreateObject("Excel.Application")		
		objExcel.AlertBeforeOverwriting = False
		objExcel.Visible = False
		objExcel.DisplayAlerts = False
		'Open batch execution result file
		Set objWorkbook = objExcel.Workbooks.Open(sResultSheetLocation)
		If iSheetNumber = "" Then
			iSheetNumber = 1
		End If  
		'Creating work sheet object
		Set objWorkSheet = objWorkbook.Worksheets(iSheetNumber)
		objWorkSheet.Activate
		'Creating used range object		
		Set objRange = objWorkSheet.UsedRange		
		iRowCount=objRange.Rows.Count
		
		'Updating test case name	
		If sTestCaseName<>"" Then
'			iColumnCount = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sResultSheetLocation,iSheetNumber,"","Test Case Name",True)
'			iColumnCount = CInt(Split(iColumnCount,":")(1))			
			iRowCount=iRowCount+1
			objWorkSheet.Cells(iRowCount,2).Value=sTestCaseName			
			objWorkSheet.Cells(iRowCount,9).Value=GBL_TESTCASE_ID
			objWorkSheet.Cells(iRowCount,10).Value=GBL_TESTSET_NAME
'			iColumnCount = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sResultSheetLocation,iSheetNumber,"","Sr No",True)							
'			iColumnCount = CInt(Split(iColumnCount, ":")(1))
			objWorkSheet.Cells(iRowCount,1).Value=Cstr(iRowCount-1)			
		End If
		
		'Updating test case execution start time	
		If GBL_TEST_EXECUTION_START_TIME<>"" Then
'			iColumnCount = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sResultSheetLocation,iSheetNumber,"","Start Time",True)	
'			iColumnCount = CInt(Split(iColumnCount, ":")(1))
			objWorkSheet.Cells(iRowCount,7).Value="st- " & Cstr(GBL_TEST_EXECUTION_START_TIME)
		End If
		'Updating test case execution end time
		If GBL_TEST_EXECUTION_END_TIME<>"" Then
'			iColumnCount = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sResultSheetLocation,iSheetNumber,"","End Time",True)	
'			iColumnCount = CInt(Split(iColumnCount, ":")(1))
			objWorkSheet.Cells(iRowCount,8).Value="et- " & Cstr(GBL_TEST_EXECUTION_END_TIME)
		End If
		'Updating test case execution duration
		If GBL_TEST_EXECUTION_TOTAL_TIME<>"" Then
'			iColumnCount = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sResultSheetLocation,iSheetNumber,"","Test Duration",True)	
'			iColumnCount = CInt(Split(iColumnCount, ":")(1))
			objWorkSheet.Cells(iRowCount,6).Value=Cstr(GBL_TEST_EXECUTION_TOTAL_TIME)
		End If
		'Updating test case execution date
		If sDate<>"" Then
'			iColumnCount = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sResultSheetLocation,iSheetNumber,"","Date",True)	
'			iColumnCount = CInt(Split(iColumnCount, ":")(1))			
			objWorkSheet.Cells(iRowCount,3).Value=sDate
		End If
		'Updating test case final result
		If sResultData<>"" Then
'			iColumnCount = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sResultSheetLocation,iSheetNumber,"","Result",True)	
'			iColumnCount = CInt(Split(iColumnCount, ":")(1))			
			objRange.SpecialCells(xlCellTypeLastCell).Activate
			objWorkSheet.Cells(iRowCount, 4).Value = UCase(Trim(sResultData))
			If iCounter = 0 Then
				objExcel.Cells(iRowCount, 4).Font.Bold = True
				objExcel.Cells(iRowCount, 4).HorizontalAlignment = xlCenter
				objExcel.Cells(iRowCount, 4).VerticalAlignment = xlCenter	
				If InStr(LCase(sResultData),"pass") <> 0 Then
					objExcel.Cells(iRowCount, 4).Interior.ColorIndex = 35
					GBL_LASTEXECUTED_ACTIONWORD_NAME = "NA"
					GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"
				ElseIf InStr(LCase(sResultData),"fail") <> 0 Then
					objExcel.Cells(iRowCount, 4).Interior.Color = RGB(255,128,128)
				End If				
			End If	
			If GBL_TEAMCENTER_LAST_LOGGEDIN_USERID<>"" Then
				GBL_TEAMCENTER_LAST_LOGGEDIN_USERID=Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",GBL_TEAMCENTER_LAST_LOGGEDIN_USERID)
			Else
				GBL_TEAMCENTER_LAST_LOGGEDIN_USERID="{EMPTY}"
			End If
			objWorkSheet.Cells(iRowCount, 11).Value = Cstr(GBL_TEAMCENTER_LAST_LOGGEDIN_USERID)
			objWorkSheet.Cells(iRowCount, 12).Value = Cstr(GBL_LASTEXECUTED_ACTIONWORD_NAME)
			objWorkSheet.Cells(iRowCount, 13).Value = Cstr(GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME)
		End If
		'Updating test case comments if any
		If sComments<>"" Then
'			iColumnCount = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sResultSheetLocation,iSheetNumber,"","Comments",True)	
'			iColumnCount = CInt(Split(iColumnCount, ":")(1))			
			objWorkSheet.Cells(iRowCount,5).Value=sComments
		End If		
		objRange.Borders.LineStyle = xlContinuous
		objRange.Range("A1").Activate
		'Save excel changes
		objWorkbook.Save
		objExcel.Quit
		wait 1
		Fn_LogUtil_UpdateTestCaseBatchExecutionResult = True
	End IF
	'Releasing excel objects
    Set objRange = Nothing
    Set objWorkSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_LogUtil_CaptureFunctionExecutionTime
'
'Function Description	 :	Function Used to calculate function performance time
'
'Function Parameters	 :  1.sAction:Action to be performed
'							2.sFunctionName: Function name
'							3.sFunctionActionName: Function Action name ( optional )
'							4.sParameterName : Parameter name ( optional )
'							5.sParameterValue : Parameter value ( optional )
'
'Function Return Value	 : 	Performance\Execution time
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Fn_ObjectDeleteOperations","Delete","","")
'Function Usage		     :  Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Fn_WorkflowProcess_Operation","","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  09-Feb-2016	    |	 1.0		|	Vrushali Sahare	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_LogUtil_CaptureFunctionExecutionTime(sAction,sFunctionName,sFunctionActionName,sParameterName,sParameterValue)
'	'Declaring varaibles
'	Dim iPerformanceDuration
'	
'	Select Case lCase(sAction)
'	    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "capturestarttime",""
'			GBL_FUNCTION_EXECUTION_START_TIME=Now()
'       '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		Case "captureendtime"
'			GBL_FUNCTION_EXECUTION_END_TIME=Now()
'			If GBL_FUNCTION_EXECUTION_START_TIME<>"" Then
'				iPerformanceDuration=Cstr(DateDiff("s",GBL_FUNCTION_EXECUTION_START_TIME,GBL_FUNCTION_EXECUTION_END_TIME))
'			End If			
'			GBL_FUNCTION_EXECUTION_START_TIME=""
'			GBL_FUNCTION_EXECUTION_END_TIME=""
'	End Select 
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_LogUtil_PrintStepHeaderLog
'
'Function Description	 :	Function use to print step header information in test case
'
'Function Parameters	 :  1.iStepNumber:Step number
'							2.sStepDescription: Step Description
'							3.sExpectedResult: Step Expected result
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :  Call Fn_LogUtil_PrintStepHeaderLog(2,GBL_STEP_DESCRIPTION,GBL_STEP_EXPECTED_RESULT)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  09-Mar-2016	    |	 1.0		|	Vrushali Sahare	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_LogUtil_PrintStepHeaderLog(iStepNumber,sStepDescription,sExpectedResult)
	If iStepNumber="" Then
		Exit Function
	End If
	'If QCStepLog Flag is true then print step information according to QC steps
	'Need to implement code
	If Cint(iStepNumber)=0 Then
		GBL_STEP_NUMBER=1
		Call Fn_LogUtil_PrintAndUpdateScriptLog("step_header","Step No. : " & Cstr(GBL_STEP_NUMBER) & vbNewLine & "Description : " & Cstr(sStepDescription) & vbNewLine & "Expected Result : " & Cstr(sExpectedResult),"","","","","")
	Else
		GBL_STEP_NUMBER = GBL_STEP_NUMBER + 1
		Call Fn_LogUtil_PrintAndUpdateScriptLog("step_header","Step No. : " & Cstr(GBL_STEP_NUMBER) & vbNewLine & "Description : " & Cstr(sStepDescription) & vbNewLine & "Expected Result : " & Cstr(sExpectedResult),"","","","","")
	End If
	GBL_STEP_DESCRIPTION=""
	GBL_STEP_EXPECTED_RESULT=""
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_LogUtil_PrintTestCaseHeaderLog
'
'Function Description	 :	Function used to print test case header log
'
'Function Parameters	 :  NA
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Test case log file should be exist
'
'Function Usage		     :	Call Fn_LogUtil_PrintTestCaseHeaderLog("")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  10-Mar-2016	    |	 1.0		|		Kundan Kudale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_LogUtil_PrintTestCaseHeaderLog()
	'Declaring variables
	Dim sParamName
	Dim iParamCount,iCounter
	Dim objNetWork,objFSO,objFile
	
	If Environment.Value("UserName")="" Then
	  Set objNetWork = CreateObject("WScript.NetWork") 
	  Environment.Value("UserName")=objNetWork.UserName
	  Set objNetWork = Nothing
	End If
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(Environment.Value("TestLogFile"),8)
	objFile.WriteLine "-----------------------------------------------------------------------------------------------"
	objFile.WriteLine "Setup Details"
	objFile.WriteLine "-----------------------------------------------------------------------------------------------"
	objFile.WriteLine "Teamcenter Release" & Fn_CommonUtil_AddDivider("Tab",1) & " - " & Environment.Value("TcRelease")
	objFile.WriteLine "Teamcenter Build" & Fn_CommonUtil_AddDivider("Tab",1) & " - " & Environment.Value("TcBuild")
	objFile.WriteLine "Tester Name" & Fn_CommonUtil_AddDivider("Tab",3) & " - " & Environment.Value("UserName")
	objFile.WriteLine Fn_CommonUtil_AddDivider("NewLine","")
	objFile.WriteLine "-----------------------------------------------------------------------------------------------"
	objFile.WriteLine "Test Case Information"
	objFile.WriteLine "-----------------------------------------------------------------------------------------------"
	objFile.WriteLine "Feature" & Fn_CommonUtil_AddDivider("Tab",3) & " - " & DataTable("Feature", dtGlobalSheet)
	objFile.WriteLine "Category" & Fn_CommonUtil_AddDivider("Tab",2) & " - " & DataTable("Category", dtGlobalSheet)	
	iParamCount = DataTable.GetSheet("Global").GetParameterCount
	For iCounter = 1 To iParamCount
	  sParamName = DataTable.GetSheet("Global").GetParameter(iCounter).Name
	  If StrComp("Version",sParamName)=0 Then
		objFile.WriteLine "Version" & Fn_CommonUtil_AddDivider("Tab",3) & " - " & DataTable("Version", dtGlobalSheet)
	  End if 
	Next
	objFile.WriteLine "TestCase" & Fn_CommonUtil_AddDivider("Tab",2) & " - " & Environment.Value("TestName")
	objFile.WriteLine "Test Run Date" & Fn_CommonUtil_AddDivider("Tab",1) & " - " & MonthName(Month(Date)) & "/" & Day(Date) & "/" & Year(Date)
	objFile.WriteLine "Test Run Time" & Fn_CommonUtil_AddDivider("Tab",1) & " - " & TimeSerial(Hour(Time) , Minute(Time), Second(Time))
	objFile.WriteLine Fn_CommonUtil_AddDivider("NewLine","")
	objFile.WriteLine "-----------------------------------------------------------------------------------------------"
	objFile.WriteLine "Test Script Logs"
	objFile.WriteLine "-----------------------------------------------------------------------------------------------"
	objFile.WriteLine Fn_CommonUtil_AddDivider("NewLine","")
	
	Set objFile =Nothing
	Set objFSO =Nothing	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_LogUtil_CreateBatchExecutionResultSummarySheet
'
'Function Description	 :	Function Used to create batch execution result excel
'
'Function Parameters	 :  1.sReportFolderLocation	:	Execution report location
'
'Function Return Value	 : 	Batch execution result excel path or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	
'
'Function Usage		     :  Call Fn_LogUtil_CreateBatchExecutionResultSummarySheet("C:\ApplePOC_AUT_10.1\Reports")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  04-Apr-2016	    |	 1.0		|	Vrushali Sahare	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_LogUtil_CreateBatchExecutionResultSummarySheet(sReportFolderLocation)
	'Declaring variables
	Const xlCenter = -4108
	Const xlLeft = -4131
	Const xlUnderlineStyleNone = -4142
	Const xlAutomatic = -4105
	Const xlNone = -4142
	Const xlContinuous = 1
	Const xlThin = 2
	Const xlDiagonalDown = 5
	Const xlDiagonalUp = 6
	Const xlEdgeLeft = 7
	Const xlEdgeTop = 8
	Const xlEdgeBottom = 9
	Const xlEdgeRight = 10
	Const xlInsideVertical = 11 
	Const xl2003=56
	Dim objExcel,objWorkBook,objWorkSheet,objRange
	Dim sExcelPath,sExcelVersion

	On Error Resume Next
		
	'Creating excel object
	Set objExcel = CreateObject("Excel.Application")
		
	If Err.Number = 0 Then
		'Retrive excel version
		sExcelVersion = objExcel.Version
		'Creating workbook object
		Set objWorkBook = objExcel.Workbooks.Add		
		If Err.Number <> 0 Then
			Fn_LogUtil_CreateBatchExecutionResultSummarySheet = False
			Exit Function
		End If			
		objExcel.Visible = False
		objExcel.DisplayAlerts = False 		
		Set objWorkBook = objExcel.Application.ActiveWorkbook
		'Remove unwanted sheets
		If objWorkbook.Worksheets.Count > 1 Then
			Do While objWorkbook.Worksheets.Count > 1
				objWorkBook.Worksheets(objWorkbook.Worksheets.Count).delete
			Loop
		Elseif objWorkbook.Worksheets.Count < 1 Then
			objExcel.ActiveWorkbook.Worksheets.Add
		End If
		'Creating sheet object
		Set objWorkSheet = objWorkBook.WorkSheets(1)
		'Updating summary information
		objWorkSheet.Name = "Test Execution Details"
	
		objWorkSheet.Cells(1,1).Value = "Sr No"
		objWorkSheet.Cells(1,2).Value = "Test Case Name"
		objWorkSheet.Cells(1,3).Value = "Date"
		objWorkSheet.Cells(1,4).Value = "Result"
		objWorkSheet.Cells(1,5).Value = "Comments"
		objWorkSheet.Cells(1,6).Value = "Test Duration"
		objWorkSheet.Cells(1,7).Value = "Start Time"   
		objWorkSheet.Cells(1,8).Value = "End Time"    
		objWorkSheet.Cells(1,9).Value = "Test Case ID"
		objWorkSheet.Cells(1,10).Value = "Test Set Name"
		objWorkSheet.Cells(1,11).Value = "TC Last LoggedIn UserID"
		objWorkSheet.Cells(1,12).Value = "Action Word Name"
		objWorkSheet.Cells(1,13).Value = "Action Word Case Name"
		
		'Formatting fields
		objWorkSheet.Columns(1).ColumnWidth = 10
		objWorkSheet.Columns(2).ColumnWidth = 35
		objWorkSheet.Columns(3).ColumnWidth = 20
		objWorkSheet.Columns(4).ColumnWidth = 10
		objWorkSheet.Columns(5).ColumnWidth = 25
		objWorkSheet.Columns(6).ColumnWidth = 30
		objWorkSheet.Columns(7).ColumnWidth = 20
		objWorkSheet.Columns(8).ColumnWidth = 10
		objWorkSheet.Columns(9).ColumnWidth = 20
		objWorkSheet.Columns(10).ColumnWidth = 40
		objWorkSheet.Columns(11).ColumnWidth = 20
		objWorkSheet.Columns(12).ColumnWidth = 40
		objWorkSheet.Columns(13).ColumnWidth = 40
		
		Set objRange = objWorkSheet.Range("A1:M1")
		objRange.HorizontalAlignment = xlCenter
		objRange.VerticalAlignment = xlCenter 
		objRange.Font.Name = "Arial"
		objRange.Font.Size = "12"
		objRange.Font.Bold = True
		objRange.Interior.ColorIndex = "37"
		objRange.Borders.LineStyle = xlContinuous   				
	
		Set objWorkSheet = objWorkBook.WorkSheets(1)
		objWorkSheet.Activate
		Set objRange = objWorkSheet.UsedRange
		objRange.Range("A1").Activate
	
		sExcelPath = sReportFolderLocation & "\BatchExecutionDetails.xlsx"	
		If sExcelVersion = "12.0" Then
			objWorkBook.SaveAs(sExcelPath), xl2003
		Else
			objWorkBook.SaveAs(sExcelPath)
		End If
		objExcel.Quit
		'Releasing all created objects
		Set objRange = Nothing
		Set objWorkSheet = Nothing
		Set objWorkBook = Nothing
		Set objExcel = Nothing
		Fn_LogUtil_CreateBatchExecutionResultSummarySheet = sExcelPath
	Else
		Fn_LogUtil_CreateBatchExecutionResultSummarySheet = False
	End If
End Function