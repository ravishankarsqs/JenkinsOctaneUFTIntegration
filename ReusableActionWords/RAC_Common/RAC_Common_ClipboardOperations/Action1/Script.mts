'! @Name 			RAC_Common_ClipboardOperations
'! @Details 		This actionword is used to perform operations on teamcenter clipboard contents
'! @InputParam1 	sAction 	: Action Name
'! @InputParam2 	sObjectName : Object Name
'! @InputParam3		bAppend 	: Append Option
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			23 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ClipboardOperations","RAC_Common_ClipboardOperations", oneIteration, "Open","AL-3000669-Enter Part Description",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ClipboardOperations","RAC_Common_ClipboardOperations", oneIteration, "GetClipBoardContents","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sObjectName,bAppend
Dim objClipboardContents,objDefaultWindow,objShellWindow
Dim sContent
Dim iCounter
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action input parameters in local variables
sAction = Parameter("sAction")
sObjectName = Parameter("sObjectName")
bAppend = Parameter("bAppend")

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Creating Object of [ Teamcenter Default ] Window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_DefaultWindow","")
'Creating Object of [ Clipboard Contents ] dialog
Set objClipboardContents =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_ClipboardContents","")

'Capture execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Teamcenter Clipboard operation",sAction,"","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ClipboardOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Check no of items copied into clipboard
If Fn_UI_Object_Operations("RAC_Common_ClipboardOperations","getroproperty",objDefaultWindow.JavaCheckBox("jckb_Clipboard"),"","attached text","")="0" Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & " ] operation on teamcenter clipboard as [ Clipboard ] checkbox does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Open [ Clipboard Contents ] dialog
objDefaultWindow.Maximize
objDefaultWindow.JavaCheckBox("jckb_Clipboard").Set "ON"

bFlag=False
'Checking Existance of [ Clipboard Contents ] dialog
If Fn_UI_Object_Operations("RAC_Common_ClipboardOperations","Exist",objClipboardContents,"","","") =False Then
	Set objClipboardContents =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_ClipboardContents@2")
	'Creating object of shell window
	Set objShellWindow =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_ShellWindow")	
	For iCounter=0 to 10
		objShellWindow.SetTOProperty "index",iCounter
		'Checking existance of [ Clipboard Contents ] dialog
		If objClipboardContents.Exist(GBL_MICRO_TIMEOUT) Then
			bFlag=True
			Exit For
		End If
	Next
	'Releasing object of shell window
	Set objShellWindow =Nothing
Else	
	bFlag=True
End If
If bFlag=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & " ] operation on teamcenter clipboard as [ Clipboard ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to open contents from clipboard	
	Case "Open"
		bFlag=False
		For iCounter=0 to Cint(Fn_UI_Object_Operations("RAC_Common_ClipboardOperations","getroproperty",objClipboardContents.JavaTable("jtbl_Contents"),"","rows",""))-1
			sContent=objClipboardContents.JavaTable("jtbl_Contents").Object.getItem(iCounter).getData().toString()
			If Trim(sObjectName)=Trim(sContent) Then
				objClipboardContents.JavaTable("jtbl_Contents").SelectCell iCounter,0
				bFlag=True
				Exit For
			End If
		Next
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to open object [ " & Cstr(sObjectName) & " ] from teamcenter clipboard as object [ " & Cstr(sObjectName) & " ] does not exist in clipboard","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on Open button
		If Fn_UI_JavaButton_Operations("RAC_Common_ClipboardOperations", "Click", objClipboardContents, "jbtn_Open")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to open object [ " & Cstr(sObjectName) & " ] from teamcenter clipboard as fail to click on [ Open ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_ClipboardOperations", "Click", objClipboardContents, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to open object [ " & Cstr(sObjectName) & " ] from teamcenter clipboard as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		'Capturing execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Teamcenter Clipboard operation",sAction,"","")		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully opened object [ " & Cstr(sObjectName) & " ] from teamcenter clipboard","","","","","")	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to retrive teamcenter clipboard contents
	Case "GetClipBoardContents"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_ClipboardOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		
		sContent=""
		sContent=objClipboardContents.JavaTable("jtbl_Contents").Object.getItem(0).getData().toString()
		For iCounter=1 to Cint(Fn_UI_Object_Operations("RAC_Common_ClipboardOperations","getroproperty",objClipboardContents.JavaTable("jtbl_Contents"),"","rows",""))-1
			sContent = sContent & "~" & objClipboardContents.JavaTable("jtbl_Contents").Object.getItem(iCounter).getData().toString()
		Next
		
		If sContent<>"" Then
			DataTable.Value("ReusableActionWordReturnValue","Global")=sContent	
		End If
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
		
		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_ClipboardOperations", "Click", objClipboardContents, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to get\store\retrive teamcenter clipboard contents as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
End Select
'Releasing all objects
Set objClipboardContents=Nothing
Set objDefaultWindow=Nothing

Function Fn_ExitTest()
	'Releasing all objects	
	Set objClipboardContents=Nothing
	Set objDefaultWindow=Nothing
	ExitTest
End Function


