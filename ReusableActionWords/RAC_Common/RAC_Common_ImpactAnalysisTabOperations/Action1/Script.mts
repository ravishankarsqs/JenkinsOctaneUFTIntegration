'! @Name 			RAC_Common_ImpactAnalysisTabOperations
'! @Details 		This actionword is used to perform Summary tab operations in Teamcenter application
'! @InputParam1 	sAction 			: Action to be performed
'! @InputParam2 	dictImpactAnalysisTabInfo 	: external dictionary parameter to pass additional details
'! @Author 			Kundan Kudale kundan.kudale@sqs.com
'! @Reviewer 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Date 			14 Jun 2017
'! @Version 		1.0
'! @Example 		dictImpactAnalysisTabInfo("Where") = "Referenced"
'! @Example 		dictImpactAnalysisTabInfo("RevisionNode") = "4501186/AA;1-AUT_55846"   
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ImpactAnalysisTabOperations","RAC_Common_ImpactAnalysisTabOperations",oneIteration,"VerifyRevisionNodeExist"
'! @Example 		dictImpactAnalysisTabInfo("RevisionNode") = "4501186/AA;1-AUT_55846"   
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ImpactAnalysisTabOperations","RAC_Common_ImpactAnalysisTabOperations",oneIteration,"VerifyNoObjectsFoundDialog"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction
Dim dictItems,dictKeys
Dim objDefaultWindow, objNoObjectsFound, objNoObjectsFound1,objPerformingAllLevelSearch
Dim iCounter
Dim bNoObjectsFoundWindow

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")

'Creating Object of Teamcenter main window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_DefaultWindow","")
Set objNoObjectsFound = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_NoObjectsFound","")
Set objNoObjectsFound1 = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_NoObjectsFound@1","")
Set objPerformingAllLevelSearch = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_PerformingAllLevelSearch","")

'Open impact analysis tab
LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select","Impact Analysis",""	
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ImpactAnalysisTabOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Verify if No Objcts dialog is displayed on select impact analysis tab
bNoObjectsFoundWindow = False
If Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations","Exist", objNoObjectsFound,"10", "", "") = True Then
	bNoObjectsFoundWindow = True	
ElseIf Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations","Exist", objNoObjectsFound1,"10", "", "") = True  Then
	bNoObjectsFoundWindow = True
	Set objNoObjectsFound = Nothing
	Set objNoObjectsFound = objNoObjectsFound1
	Set objNoObjectsFound1 = Nothing
End If

'If no objects found dialog is displayed then close it by clicking the OK button
If bNoObjectsFoundWindow Then

	If sAction = "VerifyNoObjectsFoundDialog" Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as [ No Objects Found ] dialog exists.","","","","","")
	End If
	
	If Fn_UI_JavaButton_Operations("RAC_Common_ImpactAnalysisTabOperations", "Click", objNoObjectsFound, "jbtn_OK") = False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to click on [OK] button on dialog [No Objects Found] which is displayed after selecting [Impact Analysis] tab","","","","","")
		Call Fn_ExitTest()
	Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on [OK] button on dialog [No Objects Found] which is displayed after selecting [Impact Analysis] tab","","","","","")							
	End If
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	If sAction = "VerifyNoObjectsFoundDialog" Then
		ExitAction
	End If
End If

'Capture business functionality start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Impact Analysis Tab Operations",sAction,"","")

Select Case Lcase(sAction)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify if revision node is displayed in impact analysis tab
	Case "verifyrevisionnodeexist"
	
		'Taking Items & Keys from dictionary
		dictItems = dictImpactAnalysisTabInfo.Items
		dictKeys = dictImpactAnalysisTabInfo.Keys
		
		'Loop to select values of drop downs from impact analysis tab
		For iCounter=0 to dictImpactAnalysisTabInfo.count-1
			Select Case Lcase(dictKeys(iCounter))
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "where", "depth", "rule", "display"		
				
					'Set the attached text property of java list displayed on Impact Analysis tab
					If Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations","settoproperty",objDefaultWindow.JavaList("jlst_ImpactAnalysisTab_DropDown"),"","attached text",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to verify existence of java list [" & Cstr(dictKeys(iCounter)) & "] on impact analysis tab","","","","","")
						Call Fn_ExitTest()
					End IF
					
					'Select the value of java list on impact analysis
					If Fn_UI_JavaList_Operations("RAC_Common_ImpactAnalysisTabOperations", "Select", objDefaultWindow, "jlst_ImpactAnalysisTab_DropDown", dictItems(iCounter), "", "") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to select value for dropdown [" & Cstr(dictKeys(iCounter)) & "] as [" & Cstr(dictItems(iCounter)) & "] on impact analysis tab","","","","","")
						Call Fn_ExitTest()
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected value for dropdown [" & Cstr(dictKeys(iCounter)) & "] as [" & Cstr(dictItems(iCounter)) & "] on impact analysis tab","","","","","")
					End If
					
					'If Lcase(dictKeys(iCounter)) = "where" And Lcase(dictItems(iCounter)) = "used" Then
					If Lcase(dictKeys(iCounter)) = "where" or Lcase(dictKeys(iCounter)) = "depth"  Then
						If Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations","Exist", objDefaultWindow.JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_NoObjectsFound"),"5", "", "") = True Then
							If Fn_UI_JavaButton_Operations("RAC_Common_ImpactAnalysisTabOperations", "Click", objDefaultWindow.JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_NoObjectsFound"), "jbtn_OK") = False Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to click on [OK] button on dialog [No Objects Found] which is displayed after setting value for [Where] drop down as [Used]","","","","","")
								Call Fn_ExitTest()
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on [OK] button on dialog [No Objects Found] which is displayed after setting value for [Where] drop down as [Used]","","","","","")							
							End If
						End If
						
						If objPerformingAllLevelSearch.Exist(5) Then
							If Fn_UI_JavaButton_Operations("RAC_Common_ImpactAnalysisTabOperations", "Click", objPerformingAllLevelSearch, "jbtn_Yes") = False Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to click on [ Yes ] button on dialog [ Performing All Level Search ] which is displayed after setting value [ All Levels ] from [ Depth ] drop down","","","","","")
								Call Fn_ExitTest()
							End If
							Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
						End If
					End If
					
			End Select
		Next	

		'Verify if revision node is displayed on Impact Analysis tab
		If Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations", "settoexistcheck", objDefaultWindow.JavaApplet("japt_JavaApplet").JavaStaticText("jstx_ImpactAnalysisText"), "5","label", dictImpactAnalysisTabInfo("RevisionNode")) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as text [ " & Cstr(dictImpactAnalysisTabInfo("RevisionNode")) & " ] is not displayed on impact analysis tab","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified text [ " & Cstr(dictImpactAnalysisTabInfo("RevisionNode")) & " ] is displayed on impact analysis tab","","","","","")
		End If
	
	Case "verifynoobjectsfounddialog"
	
		'Verify if revision node is displayed on Impact Analysis tab
		If Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations", "settoexistcheck", objDefaultWindow.JavaApplet("japt_JavaApplet").JavaStaticText("jstx_ImpactAnalysisText"), "5","label", dictImpactAnalysisTabInfo("RevisionNode")) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to find text [" & dictImpactAnalysisTabInfo("RevisionNode") &"] under Impact analysis tab","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Double click on the impact analysis text
		objDefaultWindow.JavaApplet("japt_JavaApplet").JavaStaticText("jstx_ImpactAnalysisText").DblClick 1, 1,"LEFT"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_TIMEOUT)
		
		'Verify existence of no objects found dialog
		bNoObjectsFoundWindow = False
		If Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations","Exist", objNoObjectsFound,"10", "", "") = True Then
			bNoObjectsFoundWindow = True	
		ElseIf Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations","Exist", objNoObjectsFound1,"10", "", "") = True  Then
			bNoObjectsFoundWindow = True
			Set objNoObjectsFound = objNoObjectsFound1
			Set objNoObjectsFound1 = Nothing
		End If
		
		If bNoObjectsFoundWindow Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as [ No Objects Found ] dialog exists after double clicking the Impact analysis text.","","","","","")
			If Fn_UI_JavaButton_Operations("RAC_Common_ImpactAnalysisTabOperations", "Click", objNoObjectsFound, "jbtn_OK") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to click on [OK] button on dialog [No Objects Found] which is displayed after selecting [Impact Analysis] tab","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on [OK] button on dialog [No Objects Found] which is displayed after selecting [Impact Analysis] tab","","","","","")							
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification failed as [ No Objects Found ] dialog doesn't exist after double clicking Impact analysis text.","","","","","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "selectnode"				
		If dictImpactAnalysisTabInfo("Where")<>"" Then
			'Set the attached text property of java list displayed on Impact Analysis tab
			If Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations","settoproperty",objDefaultWindow.JavaList("jlst_ImpactAnalysisTab_DropDown"),"","attached text","Where:")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to verify existence of java list [ Where ] on impact analysis tab","","","","","")
				Call Fn_ExitTest()
			End IF
			'Select the value of java list on impact analysis
			If Fn_UI_JavaList_Operations("RAC_Common_ImpactAnalysisTabOperations", "Select", objDefaultWindow, "jlst_ImpactAnalysisTab_DropDown", dictImpactAnalysisTabInfo("Where"), "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to select value for dropdown [ " & Cstr(dictImpactAnalysisTabInfo("Where")) & " ] from [ Where ] list on impact analysis tab","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			If objPerformingAllLevelSearch.Exist(1) Then
				If Fn_UI_JavaButton_Operations("RAC_Common_ImpactAnalysisTabOperations", "Click", objPerformingAllLevelSearch, "jbtn_Yes") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to click on [ Yes ] button on dialog [ Performing All Level Search ] which is displayed after setting value [ All Levels ] from [ Depth ] drop down","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End If
		End If
		
		If dictImpactAnalysisTabInfo("Depth")<>"" Then
			If Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations","settoproperty",objDefaultWindow.JavaList("jlst_ImpactAnalysisTab_DropDown"),"","attached text","Depth:")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to verify existence of java list [ Depth ] on impact analysis tab","","","","","")
				Call Fn_ExitTest()
			End IF
			'Select the value of java list on impact analysis
			If Fn_UI_JavaList_Operations("RAC_Common_ImpactAnalysisTabOperations", "Select", objDefaultWindow, "jlst_ImpactAnalysisTab_DropDown", dictImpactAnalysisTabInfo("Depth"), "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to select value for dropdown [ " & Cstr(dictImpactAnalysisTabInfo("Depth")) & " ] from [ Depth ] list on impact analysis tab","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			
			If dictImpactAnalysisTabInfo("Depth")="All Levels" Then				
				If objPerformingAllLevelSearch.Exist(1) Then
					If Fn_UI_JavaButton_Operations("RAC_Common_ImpactAnalysisTabOperations", "Click", objPerformingAllLevelSearch, "jbtn_Yes") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to click on [ Yes ] button on dialog [ Performing All Level Search ] which is displayed after setting value [ All Levels ] from [ Depth ] drop down","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
				End If
			End If
		End If
		Set objPerformingAllLevelSearch =Nothing
		
		If Fn_UI_Object_Operations("RAC_Common_ImpactAnalysisTabOperations", "settoexistcheck", objDefaultWindow.JavaApplet("japt_JavaApplet").JavaStaticText("jstx_ImpactAnalysisText"), "5","label", dictImpactAnalysisTabInfo("RevisionNode")) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to find revision [ " & dictImpactAnalysisTabInfo("RevisionNode") & " ] under Impact analysis tab","","","","","")
			Call Fn_ExitTest()
		End If
		
		'click on the impact analysis text
		objDefaultWindow.JavaApplet("japt_JavaApplet").JavaStaticText("jstx_ImpactAnalysisText").Click 1, 1,"LEFT"
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select revision [ " & CStr(dictImpactAnalysisTabInfo("RevisionNode")) & " ] from impacted analysis tab due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(1)
		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully revision [ " & CStr(dictImpactAnalysisTabInfo("RevisionNode")) & " ] from impacted analysis tab","","","","","")
End Select			

'Capture business functionality end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Impact Analysis Tab Operations",sAction,"","")

'Releasing teamcenter main window object
Set objDefaultWindow=Nothing

Function Fn_ExitTest()
	'Releasing teamcenter main window object
	Set objDefaultWindow=Nothing
	ExitTest
End Function


