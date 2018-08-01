'! @Name 			RAC_Common_DispatcherAdministrationConsoleOperations
'! @Details 		Action word to perform Dispatcher Administration Console opertations
'! @InputParam1 	sAction 						: String to indicate what action is to be performed
'! @InputParam2 	sInvokeOption					: Dispatcher Administration Console dialog invoke option
'! @InputParam3 	sButton 						: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Mohini Deshmukh Mohini.Deshmukh@SQS.com
'! @Date 			28 Aug 2017
'! @Version 		1.0
'! @Example 		dictDispatcher("PrimaryColumn")="Primary Objects"
'! @Example 		dictDispatcher("PrimaryValue")="229983"
'! @Example 		dictDispatcher("VerifyColumn")="State"				
'! @Example 		dictDispatcher("VerifyValues")="COMPLETE"	
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_DispatcherAdministrationConsoleOperations","RAC_Common_DispatcherAdministrationConsoleOperations",OneIteration,"SearchAndVerify","Menu","Close"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sButton
Dim objDispatcherRequestAdministrationConsole,objRequests
Dim iCount,iCounter,iRowIndex
Dim aColumns,aValues
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sButton = Parameter("sButton")

'Invoking [ Dispatcher Administration Console ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","TranslationAdministratorConsoleALL"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_DispatcherAdministrationConsoleOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Creating object of [ Dispatcher Administration Console ] dialog
Set objDispatcherRequestAdministrationConsole =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_DispatcherRequestAdministrationConsole","")		
'Checking existance of [  Dispatcher Administration Console ] dialog
If Fn_UI_Object_Operations("RAC_Common_DispatcherAdministrationConsoleOperations","Exist", objDispatcherRequestAdministrationConsole,"","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Dispatcher Administration Consoles ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_DispatcherAdministrationConsoleOperations",sAction,"","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Dispatcher Administration Console
	Case "SearchAndVerify"
		'Selecting provider
		If dictDispatcher("Provider")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_DispatcherAdministrationConsoleOperations","SetTOProperty",objDispatcherRequestAdministrationConsole.JavaList("jlst_RequestsFilter"),"","attached text","Provider")
			If Fn_UI_JavaList_Operations("RAC_Common_DispatcherAdministrationConsoleOperations", "Select", objDispatcherRequestAdministrationConsole,"jlst_OutputTemplate",dictDispatcher("Provider"),"", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Provide [ " & Cstr(dictDispatcher("Provider")) & " ] from Dispatcher Administration Console dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
		End If
		'Selecting service
		If dictDispatcher("Service")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_DispatcherAdministrationConsoleOperations","SetTOProperty",objDispatcherRequestAdministrationConsole.JavaList("jlst_RequestsFilter"),"","attached text","Service")
			If Fn_UI_JavaList_Operations("RAC_Common_DispatcherAdministrationConsoleOperations", "Select", objDispatcherRequestAdministrationConsole,"jlst_RequestsFilter",dictDispatcher("Service"),"", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Service [ " & Cstr(dictDispatcher("Service")) & " ] from Dispatcher Administration Console dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
		End If
		'Selecting state
		If dictDispatcher("State")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_DispatcherAdministrationConsoleOperations","SetTOProperty",objDispatcherRequestAdministrationConsole.JavaList("jlst_RequestsFilter"),"","attached text","State")
			If Fn_UI_JavaList_Operations("RAC_Common_DispatcherAdministrationConsoleOperations", "Select", objDispatcherRequestAdministrationConsole,"jlst_RequestsFilter",dictDispatcher("State"),"", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select State [ " & Cstr(dictDispatcher("State")) & " ] from Dispatcher Administration Console dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(1)
		End If
		
		'Click on refresh all requests button
		If Fn_UI_JavaToolbar_Operations("RAC_Common_DispatcherAdministrationConsoleOperations", "Click", objDispatcherRequestAdministrationConsole, "jtlbr_DispatcherToolbar", "Refresh All Requests (SHIFT + F5)", "", "", "")= False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Refresh All Requests (F5) ] button from Dispatcher Administration Console dialog","","","","","")
			Call Fn_ExitTest()
		End If
		wait 6
		Call Fn_RAC_ReadyStatusSync(6)
		
		'Check existence of primary object 
		Set objRequests = objDispatcherRequestAdministrationConsole.JavaTable("jtbl_Requests")
		bFlag = False
		For iCounter = 0 to cint(objRequests.GetROProperty("rows")) -1
			If Trim(objRequests.GetCellData(iCounter,dictDispatcher("PrimaryColumn"))) =Trim(dictDispatcher("PrimaryValue")) Then
				bFlag = True
				iRowIndex = iCounter
				Exit For
			End If
		Next
		
		If bFlag = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as column [ " & Cstr(dictDispatcher("PrimaryColumn")) & " ] does not contain  [ " & Cstr(dictDispatcher("PrimaryValue")) & " ] value on Dispatcher Administration Console dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Sort the table result in descending order based on creation date
		objDispatcherRequestAdministrationConsole.JavaTable("jtbl_Requests").SelectColumnHeader "Creation Date"
		Wait 2
		objDispatcherRequestAdministrationConsole.JavaTable("jtbl_Requests").SelectColumnHeader "Creation Date"
		Wait 2
		
		'Select the row in table
		If Fn_UI_JavaTable_Operations("RAC_Common_DispatcherAdministrationConsoleOperations","SelectRow",objDispatcherRequestAdministrationConsole,"jtbl_Requests", iRowIndex,"","","","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select row [" & iRowIndex & "] on Dispatcher Administration Console dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Verify Column values
		aColumns=Split(dictDispatcher("VerifyColumn"),"~")
       	aValues=Split(dictDispatcher("VerifyValues"),"~")
		
		For iCounter=0 to ubound(aColumns)
			bFlag = False
			For iCount = 0 to 20								
				If Fn_UI_JavaTable_Operations("RAC_Common_DispatcherAdministrationConsoleOperations","GetCellData",objDispatcherRequestAdministrationConsole,"jtbl_Requests",iRowIndex,aColumns(iCounter),"","","")  = Trim(aValues(iCounter)) Then
					bFlag = True
					Exit For
				End If				
				If Fn_UI_JavaToolbar_Operations("RAC_Common_DispatcherAdministrationConsoleOperations", "Click", objDispatcherRequestAdministrationConsole, "jtlbr_DispatcherToolbar", "Refresh Request (F5)", "", "", "")= False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Refresh Requests (F5) ] button from Dispatcher Administration Console dialog","","","","","")
					Call Fn_ExitTest()
				End If
				wait 6
				Call Fn_RAC_ReadyStatusSync(3)
			Next
			If bFlag = False Then
				Exit For
			End IF
		Next
		
		If bFlag = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as column [ " & Cstr(dictDispatcher("VerifyColumn")) & " ] does not contain  [ " & Cstr(dictDispatcher("VerifyValues")) & " ] value on Dispatcher Administration Console dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sButton<>"" Then
			'Click on button
			If Fn_UI_JavaButton_Operations("RAC_Common_DispatcherAdministrationConsoleOperations","Click",objDispatcherRequestAdministrationConsole,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button from Dispatcher Administration Console dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
		
		'Capturing execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_DispatcherAdministrationConsoleOperations",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operations on [ Dispatcher Administration Console ] dialog due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform operation [ " & Cstr(sAction) & " ] on [ Dispatcher Administration Console ] dialog","","","","","")
		End If		
End Select

'Creating object of [ Dispatcher Administration Console ] dialog
Set objDispatcherRequestAdministrationConsole=Nothing

Function Fn_ExitTest()
	'Creating object of [ Dispatcher Administration Console ] dialog
	Set objDispatcherRequestAdministrationConsole=Nothing
	ExitTest
End Function

