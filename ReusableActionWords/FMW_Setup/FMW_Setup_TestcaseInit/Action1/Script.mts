'! @Name FMW_Setup_TestcaseInit
'! @Details To initialise test case settings
'! @Author Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer Kundan Kudale kundan.kudale@sqs.com
'! @Date 10-Mar-2016
'! @Version 1.0
'! @Example LoadAndRunAction "FMW_Setup\FMW_Setup_TestcaseInit", "FMW_Setup_TestcaseInit", oneIteration

Option Explicit

GBL_TEST_EXECUTION_START_TIME = Now()
GBL_LAST_LOG_UPDATION_TIME= Now()
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Variables Declaration
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Dim sTestLogFile

Environment.Value("TestLogFile") = ""
Environment.Value("IsGroupRoleRequired")="False"
On Error Resume Next

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Set the Automation Folder path value to Script Variable
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Environment.Value("AutomationDirPath") = Fn_CommonUtil_EnvironmentVariablesOperations("Get","User","AutomationDir","")
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Import Environment Variable File
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Environment.LoadFromFile(Environment.Value("AutomationDirPath") & "\AutomationXML\SetupXML\EnvironmentVariables.xml")

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Create Test Log File & Write QART Specifi Info
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Create Log File
sTestLogFile =Fn_LogUtil_CreateTestCaseLogFile(Environment.Value("TestName") & ".log")

If sTestLogFile <> "" Then
	Environment.Value("TestLogFile") = sTestLogFile
End If

'Terminating excel process
Call Fn_CommonUtil_WindowsApplicationOperations("terminateall", "EXCEL.EXE")
Call Fn_CommonUtil_WindowsApplicationOperations("terminateall", "EXCEL.EXE *32")

Environment.Value("PRINT_STANDARD_LOGS")=False
Environment.Value("CURRENTLY_RUNNING_APPLICATIONS")=""

GBL_CURRENT_EXECUTABLE_APP=""

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Run Action Only for First Row of DataTable
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Call Fn_Setup_SetActionIterationMode("")
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Adding Test case details in Btach Execution Details file
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

If LCase(Environment.Value("QCStepLog"))="true" Then
	Dim objQCUtil
	
	Set objQCUtil=QCUtil.CurrentTestSet
	GBL_TESTSET_NAME=objQCUtil.Name
	Set objQCUtil=Nothing
	
	Set objQCUtil=QCUtil.CurrentTest
	GBL_TESTCASE_ID=objQCUtil.ID
	Set objQCUtil=Nothing
End If

If Environment.Value("TestName")<>"SendExecutionResultMail" Then
	Call Fn_LogUtil_UpdateTestCaseBatchExecutionResult("",Environment.Value("TestName"),Cstr(Date()) & " - " & Cstr(Time()),"","","")
End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
If LCase(Environment.Value("StepLog"))="true" Then

	'Print Test case header log in Test Log File
	Call Fn_LogUtil_PrintTestCaseHeaderLog()
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Print Action Execution Start
	Call Fn_LogUtil_PrintAndUpdateScriptLog("updatelog", " [ " & Cstr(time) & " ]  -  QTP [ " & Environment.Value("ActionName") & " ] - Start", "", "", "", "", "")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
If LCase(Environment.Value("QCStepLog"))="true" Then
	Call Fn_LogUtil_UpdateQCStepLog("Test Execution Start","","Test Case [ " & Environment.Value("TestName") & " ] Execution Start","","")
End If

