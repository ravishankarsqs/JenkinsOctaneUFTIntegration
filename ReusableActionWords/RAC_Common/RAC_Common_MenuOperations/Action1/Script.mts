'! @Name RAC_Common_MenuOperations
'! @Details This actionword is used to perform menu operations in Teamcenter application.
'! @InputParam1. sAction = Action to be performed
'! @InputParam2. sMenuLabel = Menu label tag
'! @Author Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer Kundan Kudale kundan.kudale@sqs.com
'! @Date 25 Mar 2016
'! @Version 1.0
'! @Example  LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileNew","RAC_Common_Menu"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction, sMenuLabel
Dim objDefaultWindow,objWinMenu
Dim aMenuLabel
Dim bReturn,bTimeCaptureFlag
Dim sTempMenuLabel

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sMenuLabel = Parameter("sMenuLabel")

'Creating object of [ teamcenter default ] Window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jwnd_DefaultWindow","")

If sMenuLabel<>"" Then
	sTempMenuLabel=sMenuLabel
	'Storing menu label
	sMenuLabel=Fn_RAC_GetXMLNodeValue("RAC_Common_MenuOperations","",sMenuLabel)	
	If sMenuLabel=False Then
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to fetch value of menu label [ " & Cstr(sMenuLabel) & " ] from XML while performing menu operation","","","","DONOTSYNC","")
		Call Fn_ExitTest()
	End If
End If

'Capturing execution start time
bTimeCaptureFlag=False
IF GBL_FUNCTION_EXECUTION_START_TIME="" Then
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Menu operation",sAction,"Menu Name",sMenuLabel)
	bTimeCaptureFlag=True
End If

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Code to handle File : File -> Close menu
If Trim(sMenuLabel) = "File:Close" or Trim(sMenuLabel) = "FileClose" Then
	'Checking existance of [ teamcenter default  ] window
	 If Fn_UI_Object_Operations("RAC_Common_MenuOperations", "Exist",objDefaultWindow,"","","") Then
		If Fn_UI_Object_Operations("RAC_Common_MenuOperations", "getroproperty",objDefaultWindow,"","enabled","")=1 Then
			If InStr(Fn_UI_Object_Operations("RAC_Common_MenuOperations", "getroproperty",objDefaultWindow,"","title",""), "My Teamcenter")  > 0 Or InStr(Fn_UI_Object_Operations("RAC_Common_MenuOperations", "getroproperty",objDefaultWindow,"","title",""), "Organization")  > 0 Or InStr(Fn_UI_Object_Operations("RAC_Common_MenuOperations", "getroproperty",objDefaultWindow,"","title",""), "Lifecycle Viewer")  > 0 Then
				LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","Back","",""
				ExitAction
			End If
		Else
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select menu [ " & Cstr(sMenuLabel) & " ] as teamcenter main window is in disabled state","","","","","")
			Call Fn_ExitTest()
		End If
	Else
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select menu [ " & Cstr(sMenuLabel) & " ] as teamcenter application does not exist","","","","","")
		Call Fn_ExitTest()
	End If
End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select menu
	Case "Select","SelectAndDoNotPrintLog","SelectWithAdditionalSync"
		On Error Resume Next
		Call Fn_Setup_ReporterFilter("disableall")
		If Fn_UI_JavaMenu_Operations("RAC_Common_MenuOperations","Select",objDefaultWindow, sMenuLabel)=False Then
			objDefaultWindow.Click 100,20
			Call Fn_Setup_ReporterFilter("enableall")
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"KeyPress",sTempMenuLabel
'			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select menu [ " & Cstr(sMenuLabel) & " ]","","","","","")
'			Call Fn_ExitTest()
		Else
			Call Fn_Setup_ReporterFilter("enableall")
			If sAction="SelectWithAdditionalSync" Then
				Call Fn_RAC_ReadyStatusSync(GBL_MAX_SYNC_ITERATIONS)
			Else
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
			End If
			If bTimeCaptureFlag=True Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
			End If
			If sAction<>"SelectAndDoNotPrintLog" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected menu [ " & Cstr(sMenuLabel) & " ]","","","","DONOTSYNC","")	
			End If			
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select menu
	Case "SelectWithoutSync","SelectWithoutSyncAndDoNotPrintLog"
		If Fn_UI_JavaMenu_Operations("RAC_Common_MenuOperations","Select",objDefaultWindow, sMenuLabel)=False Then
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select menu [ " & Cstr(sMenuLabel) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			If bTimeCaptureFlag=True Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
			End If
			If sAction<>"SelectWithoutSyncAndDoNotPrintLog" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected menu [ " & Cstr(sMenuLabel) & " ]","","","","DONOTSYNC","")
			End If		
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'This case is used to select the menu option using key press method
	Case "KeyPress"
		'Split Menu String
		aMenuLabel=Split(sMenuLabel,":") 
		'This is a Special case to operate menu by KeyPress method - Few menus are not getting selected by traditional way
		Select Case  Ubound(aMenuLabel)
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "1"
				objDefaultWindow.PressKey Left(aMenuLabel(0), 1), micAlt
				objDefaultWindow.PressKey Left(aMenuLabel(1), 1)
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "2"
				objDefaultWindow.PressKey Left(aMenuLabel(0), 1), micAlt
				objDefaultWindow.PressKey Left(aMenuLabel(1), 1)
				objDefaultWindow.PressKey Left(aMenuLabel(2), 1)
		End Select
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		If bTimeCaptureFlag=True Then
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected menu [ " & Cstr(sMenuLabel) & " ]","","","","DONOTSYNC","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'This case is used to check existance the menu
	Case "Exist","VerifyExist","VerifyNonExist"		
		sMenuLabel=Replace(sMenuLabel,";",":")
		aMenuLabel=Split(sMenuLabel,":")
		
		Select Case Ubound(aMenuLabel)
			Case "0"								
				bReturn =objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0", "path:=MenuItem;Shell;").Exist(GBL_DEFAULT_MIN_TIMEOUT)                  		
			Case "1"								
				bReturn =objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)                    				
			Case "2"
				bReturn = objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)
									
			 Case "3"
				bReturn = objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)                       		 						
			Case "4"
				bReturn = objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0").JavaMenu("label:="&aMenuLabel(4)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)                    								
			Case Else
				bReturn =objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0", "path:=MenuItem;Shell;").Exist(GBL_DEFAULT_MIN_TIMEOUT)			
		End Select

		If sAction="Exist" Then
			If bReturn = False Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Menu [ " & Cstr(sMenuLabel) & " ] does not exist in application","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully verified menu [ " & Cstr(sMenuLabel) & " ] exist","","","","DONOTSYNC","")
			End If
		ElseIf sAction="VerifyExist" Then
			If bReturn = False Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as menu [ " & Cstr(sMenuLabel) & " ] does not exist","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified menu [ " & Cstr(sMenuLabel) & " ] exist","","","","DONOTSYNC","")
			End If
		ElseIf sAction="VerifyNonExist" Then
			If bReturn = True Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as menu [ " & Cstr(sMenuLabel) & " ] is exist","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified menu [ " & Cstr(sMenuLabel) & " ] is not exist","","","","DONOTSYNC","")
			End If
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'This case is used to check existance the menu
	Case "Exist_Ext" 		
		sMenuLabel=Replace(sMenuLabel,";",":")
		aMenuLabel=Split(sMenuLabel,":")

		Select Case ubound(aMenuLabel) 
			Case "0"								
				bReturn =objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0", "path:=MenuItem;Shell;").Exist(GBL_DEFAULT_MIN_TIMEOUT)
			Case "1"								
				bReturn =objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)
			Case "2"
				bReturn = objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)
			 Case "3"
				bReturn = objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)
			Case "4"
				bReturn = objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0").JavaMenu("label:="&aMenuLabel(4)&"","index:=0").Exist(GBL_DEFAULT_MIN_TIMEOUT)
			Case Else
				bReturn =objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0", "path:=MenuItem;Shell;").Exist(GBL_DEFAULT_MIN_TIMEOUT)				
		End Select
		If bTimeCaptureFlag=True Then
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
		End If
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_MenuOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= bReturn
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to check menus current state
	Case "VerifyMenuEnabled" ,"State"	
		aMenuLabel=Split(sMenuLabel,":")
		Select Case ubound(aMenuLabel)
			Case "1"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0"),"","enabled","")
			Case "2"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0"),"","enabled","")
			Case "3"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0"),"","enabled","")
			Case "4"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0").JavaMenu("label:="&aMenuLabel(4)&"","index:=0"),"","enabled","")
		End Select
		
		If sAction="State" Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1		
			DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_MenuOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= bReturn
		ElseIf sAction="VerifyMenuEnabled" Then
			If Cbool(bReturn) = False Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as menu [ " & Cstr(sMenuLabel) & " ] is in disabled state","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified menu [ " & Cstr(sMenuLabel) & " ] is in enabled state ","","","","DONOTSYNC","")
			End If
		End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to check menus current state
	Case "VerifyMenuNotEnabled"
		aMenuLabel=Split(sMenuLabel,":")
		Select Case ubound(aMenuLabel)
			Case "1"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0"),"","enabled","")
			Case "2"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0"),"","enabled","")
			Case "3"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0"),"","enabled","")
			Case "4"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0").JavaMenu("label:="&aMenuLabel(4)&"","index:=0"),"","enabled","")
		End Select
		
		If sAction="VerifyMenuNotEnabled" Then
			If Cbool(bReturn) = False Then				
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified menu [ " & Cstr(sMenuLabel) & " ] is in disabled\not enabled state ","","","","DONOTSYNC","")
			Else
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as menu [ " & Cstr(sMenuLabel) & " ] is in enabled state","","","","","")
				Call Fn_ExitTest()
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'This case is used to select win menu
	Case "WinMenuSelect"
		sMenuLabel=Replace(sMenuLabel,":",";")
		'Creating win menu object
		Set objWinMenu=Nothing
		Set objWinMenu=Description.Create()
		objWinMenu("menuobjtype").value=2
		objDefaultWindow.WinMenu(objWinMenu).Select(sMenuLabel)
		Set objWinMenu=Nothing
		
		If Err.Number < 0 Then
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select menu [ " & Cstr(sMenuLabel) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
			If bTimeCaptureFlag=True Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected menu [ " & Cstr(sMenuLabel) & " ]","","","","DONOTSYNC","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'This case is used to select win menu without syncing the application
	Case "WinMenuSelectWithoutSync"
		sMenuLabel=Replace(sMenuLabel,":",";")
		'Creating win menu object
		Set objWinMenu=Nothing
		Set objWinMenu=Description.Create()
		objWinMenu("menuobjtype").value=2
		objDefaultWindow.WinMenu(objWinMenu).Select(sMenuLabel)
		Set objWinMenu=Nothing
		
		If Err.Number < 0 Then
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select menu [ " & Cstr(sMenuLabel) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			If bTimeCaptureFlag=True Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected menu [ " & Cstr(sMenuLabel) & " ]","","","","DONOTSYNC","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'This case is used to verify win menu current state
	Case "WinMenuState","VerifyWinMenuState"
		sMenu=Replace(sMenuLabel,":",";")
		set objWinMenu=description.create()
		objWinMenu("menuobjtype").value=2
		bReturn=objDefaultWindow.winmenu(objWinMenu).GetItemProperty(sMenu,"enabled")
		Set objWinMenu=Nothing
		
		If sAction="WinMenuState" Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1		
			DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_MenuOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= bReturn
		ElseIf sAction="VerifyWinMenuState" Then
			If bReturn = False Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as menu [ " & Cstr(sMenuLabel) & " ] is in disabled state","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified menu [ " & Cstr(sMenuLabel) & " ] is in enabled state","","","","DONOTSYNC","")
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'This case is used to check existance win menu
	Case "WinMenuExist","WinMenuExist_Ext","VerifyWinMenuExist","VerifyWinMenuNonExist"
		sMenuLabel=Replace(sMenuLabel,":",";")
		Set objWinMenu=description.create()
		objWinMenu("menuobjtype").value=2			
		bReturn=objDefaultWindow.winmenu(objWinMenu).GetItemProperty (sMenuLabel,"exists")
		Set objWinMenu=Nothing
		
		If sAction="WinMenuExist" Then
			If bReturn = False Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence of menu [ " & Cstr(sMenuLabel) & " ]","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully verified menu [ " & Cstr(sMenuLabel) & " ] is exist","","","","DONOTSYNC","")
			End If
		ElseIf sAction="WinMenuExist_Ext" Then
			If bTimeCaptureFlag=True Then
				Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
			End If
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1		
			DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_MenuOperations"
			DataTable.Value("ReusableActionWordReturnValue","Global")= bReturn
		ElseIf sAction="VerifyWinMenuExist" Then
			If bReturn = False Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as menu [ " & Cstr(sMenuLabel) & " ] does not exist","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified menu [ " & Cstr(sMenuLabel) & " ] is exist","","","","DONOTSYNC","")
			End If
		ElseIf sAction="VerifyWinMenuNonExist" Then
			If bReturn = True Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as menu [ " & Cstr(sMenuLabel) & " ] is exist","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified menu [ " & Cstr(sMenuLabel) & " ] does not exist","","","","DONOTSYNC","")
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'This case is used to verify win menu check state
	Case "VerifyWinMenuCheck","VerifyWinMenuUncheck"
		sMenuLabel=Replace(sMenuLabel,":",";")
		Set objWinMenu=description.create()
		objWinMenu("menuobjtype").value=2
		bReturn=Cbool(objDefaultWindow.winmenu(objWinMenu).GetItemProperty(sMenuLabel,"Checked"))   
		
		If sAction="VerifyWinMenuUncheck" Then
			If bReturn = True Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as menu [ " & Cstr(sMenuLabel) & " ] is checked","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified menu [ " & Cstr(sMenuLabel) & " ] is not checked","","","","DONOTSYNC","")
			End If
		ElseIf sAction="VerifyWinMenuCheck" Then
			If bReturn = False Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as menu [ " & Cstr(sMenuLabel) & " ] is not checked","","","","","")
				Call Fn_ExitTest()
			Else
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified menu [ " & Cstr(sMenuLabel) & " ] is checked","","","","DONOTSYNC","")
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select menu if it is in enabled state
	Case "SelectIfEnabled"	
		aMenuLabel=Split(sMenuLabel,":")
		Select Case ubound(aMenuLabel)
			Case "1"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0"),"","enabled","")
			Case "2"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0"),"","enabled","")
			Case "3"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0"),"","enabled","")
			Case "4"	
				bReturn = Fn_UI_Object_Operations("RAC_Common_MenuOperations","getroproperty",objDefaultWindow.JavaMenu("label:="&aMenuLabel(0)&"","index:=0").JavaMenu("label:="&aMenuLabel(1)&"","index:=0").JavaMenu("label:="&aMenuLabel(2)&"","index:=0").JavaMenu("label:="&aMenuLabel(3)&"","index:=0").JavaMenu("label:="&aMenuLabel(4)&"","index:=0"),"","enabled","")
		End Select
		
		If Cbool(bReturn)=True Then
			If Fn_UI_JavaMenu_Operations("RAC_Common_MenuOperations","Select",objDefaultWindow, sMenuLabel)=False Then
				GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
				GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select menu [ " & Cstr(sMenuLabel) & " ]","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_RAC_ReadyStatusSync(GBL_MAX_SYNC_ITERATIONS)
				If bTimeCaptureFlag=True Then
					Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Menu operation",sAction,"Menu Name",sMenuLabel)
				End If
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected menu [ " & Cstr(sMenuLabel) & " ]","","","","DONOTSYNC","")
			End If
		End IF
End Select	

'Validating error while performing menu operation	
If Err.Number < 0 Then
	GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MenuOperations"
	GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform Menu Operation [ " & Cstr(sAction) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If
'Setting datatable row
DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
'Releasing object of [ teamcenter default ] Window
Set objDefaultWindow=Nothing

Function Fn_ExitTest()
	'Releasing object of [ teamcenter default ] Window
	Set objDefaultWindow=Nothing
	ExitTest
End Function

