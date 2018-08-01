'! @Name 			RAC_Common_SummaryTabOperations
'! @Details 		This actionword is used to perform Summary tab operations in Teamcenter application
'! @InputParam1 	sAction 			: Action to be performed
'! @InputParam4 	sPerspective 		: Perspective name
'! @InputParam5 	sInnerTabName 		: InnerTab name
'! @InputParam6 	dictSummaryTabInfo 	: external dictionary parameter to pass additional details
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			27 Jun 2016
'! @Version 		1.0
'! @Example 		dictSummaryTabInfo("ID") = "1234"
'! @Example 		dictSummaryTabInfo("Name") = "Testing" 
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_SummaryTabOperations","RAC_Common_SummaryTabOperations",oneIteration,"Verify","All"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInnerTabName
Dim dictItems,dictKeys
Dim objDefaultWindow
Dim aValue,aKey
Dim iCounter,iCount
Dim objInsightObject
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sInnerTabName = Parameter("sInnerTabName")

'Creating Object of Teamcenter main window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_DefaultWindow","")

'Click on inner tab
If sInnerTabName<>"" Then
	LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Summary",sInnerTabName
Else
	'Selecting summary tab
	LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select","Summary",""	
End If

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_SummaryTabOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	

'Capture business functionality start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Summary Tab Operations",sAction,"","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify value from summary tab
	Case "Verify"
		'Taking Items & Keys from dictionary
		dictItems = dictSummaryTabInfo.Items
		dictKeys = dictSummaryTabInfo.Keys
		For iCounter=0 to dictSummaryTabInfo.count-1
			Select Case dictKeys(iCounter)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "ID","Program Manager","Engineering Manager","Finance","Manufacturing","Quality","Sale snd Commercial","Buyer","Platform Director","Engineering Director"			
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","label",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End IF
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations", "Exist",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit"),"","","") Then
						If dictItems(iCounter)="{BLANK}" Then								
							If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","getroproperty",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit"),"","value","")="" Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property is empty\blank","","","","DONOTSYNC","")
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property is not empty\blank","","","","","")
								Call Fn_ExitTest()
							End If
						Else
							If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","getroproperty",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit"),"","value","")=dictItems(iCounter) Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] value","","","","DONOTSYNC","")
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] value","","","","","")
								Call Fn_ExitTest()
							End If
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name","Description","Description for Change","Short Description"
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","label",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End IF
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations", "Exist",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit"),"","","") Then
						If dictItems(iCounter)="{BLANK}" Then								
							If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","getroproperty",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit"),"","value","")="" Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property is empty\blank","","","","DONOTSYNC","")
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property is not empty\blank","","","","","")
								Call Fn_ExitTest()
							End If
						Else
							If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","getroproperty",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit"),"","value","")=dictItems(iCounter) Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] value","","","","DONOTSYNC","")
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] value","","","","","")
								Call Fn_ExitTest()
							End If
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Release Status","Type"
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","label",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End IF
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations", "Exist",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit_2"),"","","") Then
						If dictItems(iCounter)="{BLANK}" Then								
							If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","getroproperty",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit_2"),"","value","")="" Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property is empty\blank","","","","DONOTSYNC","")
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property is not empty\blank","","","","","")
								Call Fn_ExitTest()
							End If
						Else
							If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","getroproperty",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit_2"),"","value","")=dictItems(iCounter) Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] value","","","","DONOTSYNC","")
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] value","","","","","")
								Call Fn_ExitTest()
							End If
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If
					'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Last Modified Date"
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","label",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End IF
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations", "Exist",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit_2"),"","","") Then
						If Instr(1,Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","getroproperty",objDefaultWindow.JavaEdit("jedt_SummaryTabEdit_2"),"","value",""),dictItems(iCounter)) Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] value","","","","DONOTSYNC","")
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] value","","","","","")
							Call Fn_ExitTest()
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name_Text","Type_Text"
					aKey=Split(dictKeys(iCounter),"_")
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","label",aKey(0) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aKey(0)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End IF
					
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations", "Exist", objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","","") Then
						If Trim(objDefaultWindow.JavaStaticText("jstx_SummaryTabValue").GetROProperty("label"))=Trim(dictItems(iCounter)) Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aKey(0)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] value","","","","DONOTSYNC","")
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aKey(0)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] value","","","","","")
							Call Fn_ExitTest()
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aKey(0)) & " ] property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If	
			End Select
		Next			
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify value from summary tab
	Case "VerifyLOVValues"
		If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","label",dictSummaryTabInfo("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictSummaryTabInfo("PropertyName")) & " ] property does not exist\available on summary tab","","","","","")
			Call Fn_ExitTest()
		End IF
		If Fn_UI_JavaButton_Operations("RAC_Common_SummaryTabOperations", "Click", objDefaultWindow, "jbtn_SummaryTabLOVDropDown")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as fail to click on LOV drop down button of [ " & Cstr(dictSummaryTabInfo("PropertyName")) & " ] on summary tab","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		aValue=Split(dictSummaryTabInfo("PropertyValue"),"~")
		For iCounter=0 to Ubound(aValue)
			bFlag=False
			For iCount=0 to objDefaultWindow.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtree_LOVTree").GetROProperty("items count")-1
				If aValue(iCounter)=objDefaultWindow.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtree_LOVTree").GetItem(iCount) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictSummaryTabInfo("PropertyName")) & " ] property contains [ " & Cstr(aValue(iCounter)) & " ] value on summary tab","","","","DONOTSYNC","")
					bFlag=True
					Exit For	
				End If
			Next
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictSummaryTabInfo("PropertyName")) & " ] property does not contain value [ " & Cstr(aValue(iCounter)) & " ] on summary tab","","","","","")
				Call Fn_ExitTest()
				Exit For
			End If
		Next
		If Fn_UI_JavaButton_Operations("RAC_Common_SummaryTabOperations", "Click", objDefaultWindow, "jbtn_SummaryTabLOVDropDown")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as fail to click on LOV drop down button of [ " & Cstr(dictSummaryTabInfo("PropertyName")) & " ] on summary tab","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'dictSummaryTabInfo("PropertyName")="ID~Name~Description"
	Case "VerifyHeader"
		aValue=Split(dictSummaryTabInfo("PropertyName"),"~")
		For iCounter=0 to Ubound(aValue)
			If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","attached text",aValue(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aValue(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations", "Exist", objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","","") Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aValue(iCounter)) & " ] property exist\available on summary tab","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aValue(iCounter)) & " ] property does not exist\available on summary tab","","","","","")
				Call Fn_ExitTest()
			End If
		Next									
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
	Case "Modify"
		'Taking Items & Keys from dictionary
		dictItems = dictSummaryTabInfo.Items
		dictKeys = dictSummaryTabInfo.Keys
		For iCounter=0 to dictSummaryTabInfo.count-1
			Select Case dictKeys(iCounter)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name","Description"
					If  dictItems(iCounter)="{BLANK}"  Then
						dictItems(iCounter) = ""
					End If
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","label",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End IF
					
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations", "Exist",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","","") Then
						If Fn_UI_JavaEdit_Operations("RAC_Common_SummaryTabOperations", "Set", objDefaultWindow, "jedt_SummaryTabEdit", dictItems(iCounter)) Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully modified [ " & Cstr(dictKeys(iCounter)) & " ] property value to [ " & Cstr(dictItems(iCounter)) & " ] from summary tab","","","","","")
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property value to [ " & Cstr(dictItems(iCounter)) & " ] from summary tab","","","","","")
							Call Fn_ExitTest()
							
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Gate/Event"
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","label",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End IF
					
					If Fn_UI_JavaButton_Operations("RAC_Common_SummaryTabOperations", "Click", objDefaultWindow, "jbtn_SummaryTabLOVDropDown")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
					If Fn_UI_JavaTree_Operations("RAC_Common_SummaryTabOperations","Select",objDefaultWindow.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtree_LOVTree"),"",dictItems(iCounter),"","") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property value to [ " & Cstr(dictItems(iCounter)) & " ] from summary tab","","","","","")
						Call Fn_ExitTest()
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully modified [ " & Cstr(dictKeys(iCounter)) & " ] property value to [ " & Cstr(dictItems(iCounter)) & " ] from summary tab","","","","","")
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					
					If Fn_UI_JavaButton_Operations("RAC_Common_SummaryTabOperations", "Click", objDefaultWindow, "jbtn_SummaryTabLOVDropDown")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name_LOVDropDown"
					If Instr(1,dictKeys(iCounter),"_") Then
						dictKeys(iCounter) = Split(dictKeys(iCounter),"_")(0)
					End If
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_SummaryTabText"),"","label",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End IF
					
					If Fn_UI_JavaButton_Operations("RAC_Common_SummaryTabOperations", "Click", objDefaultWindow, "jbtn_SummaryTabLOVDropDown")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
					If Fn_UI_JavaTree_Operations("RAC_Common_SummaryTabOperations","Select",objDefaultWindow.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtree_LOVTree"),"",dictItems(iCounter),"","") = False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property value to [ " & Cstr(dictItems(iCounter)) & " ] from summary tab","","","","","")
						Call Fn_ExitTest()
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully modified [ " & Cstr(dictKeys(iCounter)) & " ] property value to [ " & Cstr(dictItems(iCounter)) & " ] from summary tab","","","","","")
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					
					If Fn_UI_JavaButton_Operations("RAC_Common_SummaryTabOperations", "Click", objDefaultWindow, "jbtn_SummaryTabLOVDropDown")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on summary tab","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			End Select
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
	'dictSummaryTabInfo("HyperlinkName")="More Properties..."jobj_ImageHyperlink
	Case "HyperlinkClick"
'		Call Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaObject("jobj_Hyperlink"),"","text",dictSummaryTabInfo("HyperlinkName"))
		If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaObject("jobj_Hyperlink"),"","text",dictSummaryTabInfo("HyperlinkName"))=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(dictSummaryTabInfo("HyperlinkName")) & " ] link as link does not exist\available on summary tab","","","","","")
			Call Fn_ExitTest()
		End IF
		If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations", "Exist",objDefaultWindow.JavaObject("jobj_Hyperlink"),"","","") Then
			If Fn_UI_JavaObject_Operations("RAC_Common_SummaryTabOperations","Click",objDefaultWindow,"jobj_Hyperlink","10","10","")=True Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on [ " & Cstr(dictSummaryTabInfo("HyperlinkName")) & " ] link from summary tab","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(dictSummaryTabInfo("HyperlinkName")) & " ] link from summary tab","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(dictSummaryTabInfo("HyperlinkName")) & " ] link as link does not exist\available on summary tab","","","","","")
			Call Fn_ExitTest()
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
	'dictSummaryTabInfo("ImageName")="Preview2DCircle"
	Case "VerifyPreviewImageExist"
		Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_" & dictSummaryTabInfo("ImageName"),"")
		bFlag = True
		If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,"","","") = False Then
			LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"DoubleClick","Summary",""
			
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_SummaryTabOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			
			If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,"","","") = False Then
				Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_" & dictSummaryTabInfo("ImageName") & "@2","")
				If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,"","","") = False Then
					bFlag = False
				End If
			End If
			
			If dictSummaryTabInfo("ImageName")="Preview3DSquareNX" And bFlag = False Then
				Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_Preview3DSquareNX","")
				If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,1,"","") Then
					bFlag = True
				End If
				If bFlag = False Then
					Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_Preview3DSquareNX@2","")
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,1,"","") Then
						bFlag = True
					End If	
				End If
				If bFlag = False Then
					Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_Preview3DSquareNX@3","")
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,1,"","") Then
						bFlag = True
					End If	
				End If
				If bFlag = False Then
					If Environment.Value("PLMLauncher_NXEnv")="NX - FCA" Then												
						If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", JavaWindow("jwnd_DefaultWindow").InsightObject("iobj_Preview3DSquareNX@1FCA"),1,"","") Then
							bFlag = True
						End If
					End If
				End If
			End If
			
			If dictSummaryTabInfo("ImageName")="PreviewSquareNX" And bFlag = False Then
				Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_PreviewSquareNX","")
				If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,1,"","") Then
					bFlag = True
				End If
				If bFlag = False Then
					Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_PreviewSquareNX@2","")
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,1,"","") Then
						bFlag = True
					End If	
				End If
				If bFlag = False Then
					Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_PreviewSquareNX@3","")
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,1,"","") Then
						bFlag = True
					End If	
				End If
				If bFlag = False Then
					Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_PreviewSquareNX@4","")
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,1,"","") Then
						bFlag = True
					End If	
				End If
				If bFlag = False Then
					If Environment.Value("PLMLauncher_NXEnv")="NX - FCA" Then												
						If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", JavaWindow("jwnd_DefaultWindow").InsightObject("iobj_PreviewRectangleNX@1FCA"),1,"","") Then
							bFlag = True
						End If
					End If
				End If
			End If
			
			If dictSummaryTabInfo("ImageName")="PreviewRectangleNX" Then
				bFlag = False
				If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", objInsightObject,1,"","") Then
					bFlag = True
				End If
				If bFlag = False Then
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", JavaWindow("jwnd_DefaultWindow").InsightObject("iobj_PreviewRectangleNX@3"),1,"","") Then
						bFlag = True
					End If	
				End If
				If bFlag = False Then
					If Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","Exist", JavaWindow("jwnd_DefaultWindow").InsightObject("iobj_PreviewRectangleNX@4"),1,"","") Then
						bFlag = True
					End If	
				End If
			End If
			LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"DoubleClick","Summary",""
			
			GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_SummaryTabOperations"
			GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		End If
		If bFlag = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictSummaryTabInfo("ImageName")) & " ] image does not exist under Summary tab preview","","","","","")
			Call Fn_ExitTest()
'			Reporter.ReportEvent micFail, "PreviewImageVerification", "verification fail as [ " & Cstr(dictSummaryTabInfo("ImageName")) & " ] image does not exist under Summary tab preview"
'			Desktop.CaptureBitmap Environment.Value("BatchFolderName") +"\"   & Environment.Value("TestName") & ".png", True
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictSummaryTabInfo("ImageName")) & " ] image exist under Summary tab preview","","","","DONOTSYNC","")
		End IF
		Set objInsightObject=Nothing
		
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
		Case "HyperlinkNonExist"
			Select Case dictSummaryTabInfo("HyperlinkName")
				Case "NewWorkflowProcess"
					dictSummaryTabInfo("HyperlinkName")="New Workflow Process..."
			End Select
			
			Call Fn_UI_Object_Operations("RAC_Common_SummaryTabOperations","settoproperty",objDefaultWindow.JavaObject("jobj_Hyperlink"),"","text",dictSummaryTabInfo("HyperlinkName"))
			If objDefaultWindow.JavaObject("jobj_Hyperlink").Exist(1)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictSummaryTabInfo("HyperlinkName")) & " ] option not available for Action in " & Cstr(sInnerTabName) &" tab","","","","DONOTSYNC","")
			Else	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictSummaryTabInfo("HyperlinkName")) & " ] option  available for Action in " & Cstr(sInnerTabName) &" tab","","","","","")
				Call Fn_ExitTest()
			End IF
End Select			

'Capture business functionality end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Summary Tab Operations",sAction,"","")

'Releasing teamcenter main window object
Set objDefaultWindow=Nothing

Function Fn_ExitTest()
	'Releasing teamcenter main window object
	Set objDefaultWindow=Nothing
	ExitTest
End Function


