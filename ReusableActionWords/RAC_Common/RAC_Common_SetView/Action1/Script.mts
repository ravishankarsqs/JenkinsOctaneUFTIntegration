'! @Name 			RAC_Common_SetView
'! @Details 		This actionword is used to set the view in teamcenter
'! @InputParam1 	sInvokeOption 	: Show View dialog invoke option
'! @InputParam2 	sViewName 		: Full Path of the View to Set
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			23 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_SetView","RAC_Common_SetView",OneIteration,"Menu","Visualization~Compare"

Option Explicit
Err.Clear

'Declaring variables
Dim sViewName,sInvokeOption
Dim objShowView
Dim aViewName

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action input parameters in local variables
sInvokeOption = Parameter("sInvokeOption")
sViewName = Parameter("sViewName")

'Creating object of [ Show View ] Window
Set objShowView=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_ShowView","")

Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Invoking Show View dialog by calling menu [ Window->Show View->Other ]
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration, "Select", "WindowShowViewOther"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_SetView"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Checking existance of Show View dialog
If Fn_UI_Object_Operations("RAC_Common_SetView", "Exist",objShowView,"","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set view [ " & Cstr(sViewName) & " ] as [ Show View ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Set View","","View Name",sViewName)

aViewName=split(sViewName,"~")
If Fn_UI_JavaTree_Operations("RAC_Common_SetView","Expand",objShowView,"jtree_ViewTree",aViewName(0),"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand [ " & Cstr(aViewName(0)) & " ] node from show view tree","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'Select view from view tree
If Fn_UI_JavaTree_Operations("RAC_Common_SetView","Select",objShowView,"jtree_ViewTree",sViewName,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ " & Cstr(sViewName) & " ] node from show view tree","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'Click on [ OK ] button
If Fn_UI_JavaButton_Operations("RAC_Common_SetView", "Click", objShowView, "jbtn_OK") =False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select view [ " & Cstr(sViewName) & " ] as fail to click on [ OK ] button of show view dialog","","","","","")	
	Call Fn_ExitTest()
End If

Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
'Capturing execution end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Set View","","View Name",sViewName)
Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected view [ " & Cstr(sViewName) & " ] from [ Show View ] dialog","","","",GBL_MIN_SYNC_ITERATIONS,"")

'Releasing object
Set objShowView=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objShowView = Nothing 
	ExitTest
End Function


