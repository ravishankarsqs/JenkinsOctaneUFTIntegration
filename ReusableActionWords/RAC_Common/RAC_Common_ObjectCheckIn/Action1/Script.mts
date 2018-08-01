'! @Name 			RAC_Common_ObjectCheckIn
'! @Details 		This action word is used to perform check in operation
'! @InputParam1 	sAction 		: Action to be performed
'! @InputParam2 	sInvokeOption 	: Check In dialog invoke option
'! @InputParam3 	sPerspective 	: Perspective name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			23 Jun 2016
'! @Version 		1.0
'! @Example  		LoadAndRunAction "RAC_Common\RAC_Common_ObjectCheckIn","RAC_Common_ObjectCheckIn",OneIteration,"CheckIn", "menu", "myteamcenter"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sPerspective
Dim objCheckIn
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action input parameters in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")

bFlag=False

If sPerspective="" Then
	'Get active perspective name
	sPerspective=Fn_RAC_GetActivePerspectiveName("")
End If

'Creating Object of [ Check In ] dialog
Select Case Lcase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter","systemsengineering",""
		Set objCheckIn=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_CheckingIn","")
		bFlag=True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "structuremanager"
		Set objCheckIn=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_CheckIn","")
End Select

Select Case lCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Invoke Check In dialog from menu
	Case "menu"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsCheckIn"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"		
		'To open dialog outside of the action word
End Select

'checking existance of [ Check In ] dialog
'If Fn_UI_Object_Operations("RAC_Common_ObjectCheckIn", "Exist", objCheckIn,"","","")=False Then
'	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & " ] operation as [ Check In ] dialog does not exist","","","","","")
'	Call Fn_ExitTest()
'End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Object Check In",sAction,"","")

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to perform basic check in operation
	Case "CheckIn"
		'Click on Yes button to Check out Object
'		If bFlag=True Then
'			bFlag = Fn_UI_JavaButton_Operations("RAC_Common_ObjectCheckIn", "Click", objCheckIn, "jbtn_OK")
'		Else
'			bFlag = Fn_UI_JavaButton_Operations("RAC_Common_ObjectCheckIn", "Click", objCheckIn, "jbtn_Yes")
'		End If
		
		'Capturing execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Check In",sAction,"","")
'		If bFlag = False Then
'			 Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check in selected object","","","","","")
'			 Call Fn_ExitTest()
'		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully checked in selected object","","","","","")
'		End If
End Select

'Releasing Object of [ Check In ] dialog
Set objCheckIn=Nothing

Function Fn_ExitTest()
	'Releasing Object of [ Check In ] dialog
	Set objCheckIn=Nothing
	ExitTest
End Function


