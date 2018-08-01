'! @Name 			RAC_Common_ToolbarOperations
'! @Details 		This actionword is used to perform toolbar button operations
'! @InputParam1 	sAction 		: Action to be performed
'! @InputParam2 	sButtonName 	: toolbar Button tag name
'! @InputParam3 	sPopupMenuSelect: pop up menu
'! @InputParam4 	iInstance 		: toolbar Button instance counter
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			28 Mar 2016
'! @Version 		1.0
'! @Example  		LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click", "StartOpenInNX", "","RAC_Common_TLB"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction, sButtonName,sPopupMenuSelect
Dim iInstance
Dim sFilePath,sContents
Dim bFlag,bTimeCaptureFlag
Dim objToolbarButton,objDefaultWindow

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sButtonName = Parameter("sButtonName")
sPopupMenuSelect = Parameter("sPopupMenuSelect")
iInstance = Parameter("iInstance")

'Creating object of [ teamcenter default ] Window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jwnd_DefaultWindow","")

'Checking existance of Teamcenter default window
If Not objDefaultWindow.Exist(GBL_DEFAULT_TIMEOUT) Then
	GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
	GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence of teamcenter main window,toolbar button operation [ " & Cstr(sAction) & " ] fail","","","","","")
	Call Fn_ExitTest()
End If

'Retrive toolbar button name
If sButtonName<>"" Then
	sButtonName=Fn_RAC_GetXMLNodeValue("RAC_Common_ToolbarOperations","",sButtonName)	
	If sButtonName=False Then
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail perform operation [ " & Cstr(sAction) & " ] on teamcenter toolbar as button name [ " & Cstr(sButtonName) & " ] is invalid", "","","","","")
		Call Fn_ExitTest()
	End If
End If

If iInstance="" Then
	iInstance=1
End If

bTimeCaptureFlag=False
IF GBL_FUNCTION_EXECUTION_START_TIME="" Then
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
	bTimeCaptureFlag=True
End If

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'case to click on toolbar button
	Case "Click"
		bFlag=False
		Set objToolbarButton=Fn_GetToolbarButton(sButtonName,iInstance)
		If objToolbarButton is Nothing Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on toolbar button [ " & Cstr(sButtonName) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			If  "1" = objToolbarButton.GetItemProperty (sButtonName, "enabled")  Then
				objToolbarButton.Press sButtonName
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on toolbar button [ " & Cstr(sButtonName) & " ]","","","","DONOTSYNC","")
				bFlag=True
			Else
				If sButtonName <> "Open Worklist" Then
					GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
					GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on toolbar button [ " & Cstr(sButtonName) & " ]","","","","","")
					Call Fn_ExitTest()
				Else
					bFlag=False
				End If
			End If
		End If
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to perform toolbar state operations
	Case "IsEnabled","VerifyEnabled","VerifyDisabled"
		Set objToolbarButton=Fn_GetToolbarButton(sButtonName,iInstance)
		If objToolbarButton is Nothing Then
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on teamcenter toolbar as toolbar button [ " & Cstr(sButtonName) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		Else		
			If  "1" = objToolbarButton.GetItemProperty (sButtonName, "enabled")  Then
				If sAction="IsEnabled" Then
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
					DataTable.SetCurrentRow 1		
					DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_ToolbarOperations"
					DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
				ElseIf sAction="VerifyEnabled" Then
					If bTimeCaptureFlag=True Then
						Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
					End If
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified toolbar button [" & Cstr(sButtonName) & " ] is enabled","","","","DONOTSYNC","")
				ElseIf sAction="VerifyDisabled" Then
					GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
					GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as toolbar button [" & Cstr(sButtonName) & " ] is not disabled","","","","","")
					Call Fn_ExitTest()
				End If				
			Else
				If sAction="IsEnabled" Then
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
					Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
					DataTable.SetCurrentRow 1		
					DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_ToolbarOperations"
					DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
				ElseIf sAction="VerifyEnabled" Then
					GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
					GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as toolbar button [" & Cstr(sButtonName) & " ] is not enabled","","","","","")
					Call Fn_ExitTest()
				ElseIf sAction="VerifyDisabled" Then
					If bTimeCaptureFlag=True Then
						Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
					End If
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified toolbar button [" & Cstr(sButtonName) & " ] is disabled","","","","DONOTSYNC","")
				End If
			End if
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to select toolbar button popup menu
	Case "PopupMenuSelect"
		Set objToolbarButton=Fn_GetToolbarButton(sButtonName,iInstance)
		If objToolbarButton is Nothing Then
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to toolbar button [ " & Cstr(sButtonName) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		Else
			If  "1" = objToolbarButton.GetItemProperty (sButtonName, "enabled")  Then
				objToolbarButton.Click 10,10,"RIGHT"
				sContents = objDefaultWindow.WinMenu("wmnu_ContextMenu").BuildMenuPath(sPopupMenuSelect)
               	objDefaultWindow.WinMenu("wmnu_ContextMenu").Select sContents
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected popup menu [ " & Cstr(sPopupMenuSelect) & " ] of Toolbar Button [ " & Cstr(sButtonName) & " ]","","","","1","")
			Else
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenuSelect) & " ] of Toolbar Button [ " & Cstr(sButtonName) & " ]","","","","","")
				Call Fn_ExitTest()
			End if
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to select toolbar button popup menu
	Case "ShowDropdownAndSelect"
		Set objToolbarButton=Fn_GetToolbarButton(sButtonName,iInstance)
		If objToolbarButton is Nothing Then
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to toolbar button [ " & Cstr(sButtonName) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		Else
			If  "1" = objToolbarButton.GetItemProperty (sButtonName, "enabled")  Then
				objToolbarButton.ShowDropdown sButtonName
				sContents = objDefaultWindow.WinMenu("wmnu_ContextMenu").BuildMenuPath(sPopupMenuSelect)
                objDefaultWindow.WinMenu("wmnu_ContextMenu").Select sContents
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected popup menu [ " & Cstr(sPopupMenuSelect) & " ] of Toolbar Button [ " & Cstr(sButtonName) & " ]","","","","1","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select popup menu [ " & Cstr(sPopupMenuSelect) & " ] of Toolbar Button [ " & Cstr(sButtonName) & " ]","","","","","")
			End if
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to validate toolbar button popup menu
	Case "VerifyPopupMenuExist","VerifyPopupMenuNonExist"
		Set objToolbarButton=Fn_GetToolbarButton(sButtonName,iInstance)
		If objToolbarButton is Nothing Then
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to toolbar button [ " & Cstr(sButtonName) & " ] does not exist","","","","","")
			Call Fn_ExitTest()
		Else
			If  "1" = objToolbarButton.GetItemProperty (sButtonName, "enabled")  Then
				objToolbarButton.ShowDropdown sButtonName
				sContents = objDefaultWindow.WinMenu("wmnu_ContextMenu").BuildMenuPath(sPopupMenuSelect)
               	If objDefaultWindow.WinMenu("wmnu_ContextMenu").GetItemProperty (sContents,"Exists") = True Then
					If sAction="VerifyPopupMenuExist" Then
						If bTimeCaptureFlag=True Then
							Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
						End If
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified existence of popup menu [ " & Cstr(sPopupMenuSelect) & " ] of Toolbar Button [ " & Cstr(sButtonName) & " ]","","","","DONOTSYNC","")
					ElseIf sAction="VerifyPopupMenuNonExist" Then
						GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
						GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : Verification fail as popup menu [ " & Cstr(sPopupMenuSelect) & " ] of Toolbar Button [ " & Cstr(sButtonName) & " ] is exist","","","","","")
						Call Fn_ExitTest()
					End If
				Else				
					If sAction="VerifyPopupMenuNonExist" Then
						If bTimeCaptureFlag=True Then
							Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
						End If
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified non existence of popup menu [ " & Cstr(sPopupMenuSelect) & " ] of Toolbar Button [ " & Cstr(sButtonName) & " ]","","","","DONOTSYNC","")
					ElseIf sAction="VerifyPopupMenuExist" Then
						GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
						GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - FAIL : Verification fail as popup menu [ " & Cstr(sPopupMenuSelect) & " ] of Toolbar Button [ " & Cstr(sButtonName) & " ] is not exist","","","","","")
						Call Fn_ExitTest()
					End If
				End If
			Else
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as Toolbar Button [ " & Cstr(sButtonName) & " ] is disabled","","","","","")
				Call Fn_ExitTest()
			End if
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to validate toolbar button
	Case "ButtonExist","VerifyButtonExist","VerifyButtonNonExist"
		Set objToolbarButton=Fn_GetToolbarButton(sButtonName,iInstance)
		If objToolbarButton is Nothing Then
			If sAction="VerifyButtonNonExist" Then
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified Toolbar Button [ " & Cstr(sButtonName) & " ] does not exist","","","","DONOTSYNC","")
			ElseIf sAction="VerifyButtonExist" Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as toolbar Button [ " & Cstr(sButtonName) & " ] does not exist","","","","","")
				Call Fn_ExitTest()
			Else
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail toolbar button [ " & Cstr(sButtonName) & " ] does not exist","","","","","")
				Call Fn_ExitTest()
			End If	
		Else
			If sAction="ButtonExist" Then	
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				DataTable.SetCurrentRow 1		
				DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_ToolbarOperations"
				DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
			ElseIf sAction="VerifyButtonExist" Then
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Toolbar operation",sAction,"Toolbar button Name",sButtonName)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified Toolbar Button [ " & Cstr(sButtonName) & " ] is exist","","","","DONOTSYNC","")
			ElseIf sAction="VerifyButtonNonExist" Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as Toolbar Button [ " & Cstr(sButtonName) & " ] is exist","","","","","")
				Call Fn_ExitTest()
			End If	
		End If		
End Select

If Err.Number <> 0 Then
	GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ToolbarOperations"
	GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] operation on navigation tree due to error number as [ " & Cstr(Err.Number) & " ] and error description as [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing object
Set objToolbarButton=Nothing
Set objDefaultWindow=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objToolbarButton=Nothing
	Set objDefaultWindow=Nothing	
	ExitTest
End Function

Function Fn_GetToolbarButton(sButtonName,iInstance)
	'Declaring required variables
	Dim iInstanceCount,iCounter
	Dim sContents
	Dim objChild,objJavaToolbar
	Dim bFlag

	iInstanceCount=1
	'Checking existance of Teamcenter default window
	If Fn_UI_Object_Operations("RAC_Common_ToolbarOperations","Exist", objDefaultWindow, GBL_DEFAULT_TIMEOUT,"","") Then
		bFlag=False
		Set objJavaToolbar = Description.Create() 
		objJavaToolbar("to_class").Value = "JavaToolbar"
		objDefaultWindow.Maximize
		'following statement is used to change the focus to the main window
		objDefaultWindow.Type(micAlt) 
		Wait GBL_MICRO_TIMEOUT
		'Get the total of Toolbar objects
		Set objChild =Fn_UI_Object_GetChildObjects("RAC_Common_ToolbarOperations", objDefaultWindow, "to_class", "JavaToolbar")
	 
		For iCounter = 0 to objChild.count-1
			sContents = objChild(iCounter).GetContent()
			If Instr(sContents, sButtonName) > 0 Then
				If Cint(iInstanceCount)=Cint(iInstance) Then
					Set Fn_GetToolbarButton = objChild(iCounter)
					bFlag=True
					Exit For
				Else
					iInstanceCount=iInstanceCount+1
				End If
			End If
		Next
 
		If bFlag=False Then
			Set Fn_GetToolbarButton =Nothing
		End If
		'Releasing required objects
		Set objJavaToolbar=Nothing
		Set objChild =Nothing
	End If
End Function

