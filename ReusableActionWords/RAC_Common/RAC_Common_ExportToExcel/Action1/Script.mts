'! @Name 			RAC_Common_ExportToExcel
'! @Details 		Action word to perform export to excel opertations
'! @InputParam1 	sAction 						: String to indicate what action is to be performed
'! @InputParam2 	sInvokeOption					: Export To Excel dialog invoke option
'! @InputParam3 	sOutputOption	 				: Export To Excel output option
'! @InputParam4 	bCheckOutObjectBeforeExport	 	: Checkout object before Export option flag
'! @InputParam5 	sOutputTemplateOption	 		: Export To Excel output template option
'! @InputParam6 	sOutputTemplateName	 			: Export To Excel output template name
'! @InputParam7 	sButton 						: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			28 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ExportToExcel","RAC_Common_ExportToExcel",OneIteration,"Export","Menu","Static Snapshot","","Export All Visible Columns","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sOutputOption,bCheckOutObjectBeforeExport,sOutputTemplateOption,sOutputTemplateName,sButton
Dim objExportToExcel

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sOutputOption= Parameter("sOutputOption")
bCheckOutObjectBeforeExport= Parameter("bCheckOutObjectBeforeExport")
sOutputTemplateOption= Parameter("sOutputTemplateOption")
sOutputTemplateName= Parameter("sOutputTemplateName")
sButton = Parameter("sButton")

'Invoking [ Export To Excel ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsExportObjectsToExcel"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Creating object of [ Export To Excel ] dialog
Set objExportToExcel =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_ExportToExcel","")	

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ExportToExcel"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of [  Export To Excel ] dialog
If Fn_UI_Object_Operations("RAC_Common_ExportToExcel","Exist", objExportToExcel,"","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Export To Excels ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_ExportToExcel",sAction,"","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Export To Excel
	Case "Export"
		'Selecting Output option
		If sOutputOption<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_ExportToExcel","SetTOProperty",objExportToExcel.JavaRadioButton("jrdb_Output"),"","attached text",sOutputOption)
			If Fn_UI_JavaRadioButton_Operations("RAC_Common_ExportToExcel","Set",objExportToExcel,"jrdb_Output","ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to export selected objects to excel as fail to select Output option [ " & Cstr(sOutputOption) & " ] from Export To Excel dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		
		'Selecting Output Template option
		If sOutputTemplateOption<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_ExportToExcel","SetTOProperty",objExportToExcel.JavaRadioButton("jrbd_OutputTemplate"),"","attached text",sOutputTemplateOption)
			If Fn_UI_JavaRadioButton_Operations("RAC_Common_ExportToExcel","Set",objExportToExcel,"jrbd_OutputTemplate","ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to export selected objects to excel as fail to select Output Template option [ " & Cstr(sOutputTemplateOption) & " ] from Export To Excel dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		
		'Selecting Output Template Name
		If sOutputTemplateName<>"" Then
			If Fn_UI_JavaList_Operations("RAC_Common_ExportToExcel", "Select", objExportToExcel,"jlst_OutputTemplate",sOutputTemplateName,"", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to export selected objects to excel as fail to select Output Template name [ " & Cstr(sOutputTemplateName) & " ] from Export To Excel dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If	
		
		'Click on Add button
		If Fn_UI_JavaButton_Operations("RAC_Common_ExportToExcel","Click",objExportToExcel,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to export selected objects to excel as fail to click on [ OK ] button from Export To Excel dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		'Capturing execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_ExportToExcel",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to export selected objects to excel due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully exported selected objects to excel","","","","","")
		End If		
End Select

'Creating object of [ Export To Excel ] dialog
Set objExportToExcel=Nothing

Function Fn_ExitTest()
	'Creating object of [ Export To Excel ] dialog
	Set objExportToExcel=Nothing
	ExitTest
End Function
