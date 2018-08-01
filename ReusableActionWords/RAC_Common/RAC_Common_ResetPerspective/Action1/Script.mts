'! @Name 		RAC_Common_ResetPerspective
'! @Details 	To reset a perspective in Teamcenter
'! @Author 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 	Kundan Kudale kundan.kudale@sqs.com
'! @Date 		26 Mar 2016
'! @Version 	1.0
'! @Example 	LoadAndRunAction "RAC_Common\RAC_Common_ResetPerspective","RAC_Common_ResetPerspective",OneIteration

Option Explicit
Err.Clear

'Declaring variables
Dim objResetPerspective

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Calling menu [ Window -> Reset Perspective ]
LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","WindowResetPerspective"
'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ResetPerspective"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Reset Perspective","","","")

'Creating object Reset respective dialog
Set objResetPerspective=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jwnd_ResetPerspective","")

'Clicking on [ Yes ] button
If Fn_UI_JavaButton_Operations("RAC_Common_ResetPerspective", "Click", objResetPerspective,"jbtn_Yes")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Yes ] button while performing reset pesspective operation","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Reset Perspective","","","")

'Releasing object Reset respective dialog
Set objResetPerspective=Nothing

Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully reset perspective","","","","DONOTSYNC","")

Function Fn_ExitTest()
	'Releasing object Reset respective dialog
	Set objResetPerspective=Nothing
	ExitTest
End Function
