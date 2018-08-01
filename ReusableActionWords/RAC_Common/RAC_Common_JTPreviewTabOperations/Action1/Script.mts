'! @Name 			RAC_Common_JTPreviewTabOperations
'! @Details 		This actionword is used to perform JT Preview tab operations in Teamcenter application
'! @InputParam1 	sAction 				: Action to be performed
'! @Author 			Mohjini Deshmukh mohini.deshmukh@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			03 Aug 2016
'! @Version 		1.0
'! @Example 		dictJTPreviewTabInfo("ImageName")="ViewerTabDatasetImage_Circle"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_JTPreviewTabOperations","RAC_Common_JTPreviewTabOperations",OneIteration,"VerifyDatasetImageExist"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction
Dim objDefaultWindow
Dim  objInsightObject

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")

'Creating Object of Teamcenter main window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_DefaultWindow","")

'Selecting Viewer tab
LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select","JT Preview",""

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_JTPreviewTabOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
	
'Capture business functionality start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","JT Preview Tab Operations",sAction,"","")

Select Case sAction
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
	Case "VerifyDatasetImageExist"
	
		Set objInsightObject = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_" & dictJTPreviewTabInfo("ImageName"),"")
				
		If Fn_UI_Object_Operations("RAC_Common_JTPreviewTabOperations","Exist", objInsightObject,"", "", "") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as image [ " & Cstr(dictJTPreviewTabInfo("ImageName")) & " ] was not found under JT Preview tab","","","","DONOTSYNC","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as image [ " & Cstr(dictJTPreviewTabInfo("ImageName")) & " ] was found under JT Preview tab","","","","","")
		End If
		
End Select			

'Capture business functionality end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","JT Preview Tab Operations",sAction,"","")

'Releasing teamcenter main window object
Set objDefaultWindow=Nothing

Function Fn_ExitTest()
	'Releasing teamcenter main window object
	Set objDefaultWindow=Nothing
	ExitTest
End Function

