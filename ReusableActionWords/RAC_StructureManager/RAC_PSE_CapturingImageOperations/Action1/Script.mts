'! @Name 			RAC_PSE_CapturingImageOperations
'! @Details 		Action word to perform operations on Capturing Image dialog
'! @InputParam1 	sAction 			: Action to be performed e.g. AutoReplaceBasic
'! @InputParam2 	sInvokeOption 		: Method to invoke Replace dialog e.g. menu
'! @InputParam3 	sDatasetName		: Dataset Name
'! @InputParam4 	sImageFormat		: Image Format
'! @InputParam5 	sDescription	 	: Description
'! @InputParam6 	sButton			 	: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			28 June 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_PSE\RAC_PSE_CapturingImageOperations","RAC_PSE_CapturingImageOperations",OneIteration, "autobasiccapture","graphicstoolbar","","","",""

Option Explicit
Err.Clear

'Declaring varaibles

'Variable Declaration
Dim sAction,sInvokeOption,sDatasetName,sImageFormat,sDescription,sButton
Dim objCanvasBean,objCapturingImage
Dim iCaptureImageCount

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sDatasetName = Parameter("sDatasetName")
sImageFormat = Parameter("sImageFormat")
sDescription = Parameter("sDescription")
sButton = Parameter("sButton")

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Set objCanvasBean=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jobj_CanvasBean","")
Set objCapturingImage=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_CapturingImage","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_CapturingImageOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Invoking [ New Dataset ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "graphicstoolbar",""
		objCanvasBean.Object.getViewerBean.setToolbarVisibility "Create Markup", True
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		If Fn_UI_InsightObject_Operations(sFunctionName,"click",JavaWindow("jwnd_StructureManager"),"iobj_ImageCapture",1,1,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to capture image as fail to click on [ Image Capture ] button from [ Graphics ] tab","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Checking existance of Capturing Image dialog
If Fn_UI_Object_Operations("RAC_PSE_CapturingImageOperations","Exist", objCapturingImage,GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Capturing Image ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting dataset count
If Lcase(sAction)= "capture" or Lcase(sAction)="autobasiccapture" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACCaptureImageCount","","")
	iCaptureImageCount=Fn_CommonUtil_DataTableOperations("GetValue","RACCaptureImageCount","","")
	If iCaptureImageCount="" Then
		iCaptureImageCount=1
	Else
		iCaptureImageCount=iCaptureImageCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACCaptureImageCount",iCaptureImageCount,"")
	DataTable.SetCurrentRow iCaptureImageCount
End If

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to Create Capturing Image	
	Case "capture","autobasiccapture"				
		If Lcase(sAction)="autobasiccapture" Then
			sDatasetName="Assign"
			sImageFormat="JPEG 24bit"
		End If
		
		'Setting Dataset name
		If sDatasetName="Assign" Then
			sDatasetName = "AUT_Image" & Cstr(Fn_CommonUtil_GenerateRandomNumber(5))
		End If
		If Fn_UI_JavaEdit_Operations("RAC_PSE_CapturingImageOperations","Set",objCapturingImage,"jedt_DatasetName",sDatasetName) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create Capturing Image as fail to set dataset name value","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(1)
				
		'Setting Dataset Image format
		If sImageFormat<>"" Then
			If Fn_UI_JavaList_Operations("RAC_PSE_CapturingImageOperations", "Select", objCapturingImage, "jlst_ImageFormat", sImageFormat, "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create Capturing Image as fail to select Image Format value","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
		End If
		
		'Get dataset name		
		sDatasetName = Cstr(Fn_UI_JavaEdit_Operations("RAC_PSE_CapturingImageOperations","GetText",objCapturingImage,"jedt_DatasetName","" ))
		
		'Clicking On OK Button to create Dataset
		If Fn_UI_JavaButton_Operations("RAC_PSE_CapturingImageOperations", "Click",objCapturingImage,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create Capturing Image as fail to click on [ OK ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully capture image from Graphics tab and save with name as [ " & Cstr(sDatasetName) & " ]","","","","","")
		
		'Set value of Capturing Image name in datatable column
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACCaptureImageName","","")
		DataTable.SetCurrentRow iCaptureImageCount		
		DataTable.Value("RACCaptureImageName","Global") = sDatasetName	
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Relasing Object of [ Capturing Image ] Dialog
Set objCapturingImage=Nothing
Set objCanvasBean=Nothing

Public Function Fn_ExitTest()
	'Relasing Object of [ Capturing Image ] Dialog
	Set objCapturingImage=Nothing
	Set objCanvasBean=Nothing
	ExitTest
End Function

