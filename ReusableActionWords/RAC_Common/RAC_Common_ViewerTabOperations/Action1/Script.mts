'! @Name 			RAC_Common_ViewerTabOperations
'! @Details 		This actionword is used to perform Viewer tab operations in Teamcenter application
'! @InputParam1 	sAction 				: Action to be performed
'! @InputParam2 	sInnerTabName 			: InnerTab name
'! @InputParam3 	sObjectCheckOutOption 	: Object checkout option
'! @InputParam4 	dictViewerTabInfo 		: external dictionary parameter to pass additional details
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			03 Aug 2016
'! @Version 		1.0
'! @Example 		dictViewerTabInfo("Name")="AUT_Item"
'! @Example 		dictViewerTabInfo("Description")="AUT_Item Description"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ViewerTabOperations","RAC_Common_ViewerTabOperations",OneIteration,"Verify","General","Menu"
'! @Example 		dictViewerTabInfo("PropertyName")="ID~Name~Description"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ViewerTabOperations","RAC_Common_ViewerTabOperations",OneIteration,"VerifyHeader","General","Menu"
'! @Example 		dictViewerTabInfo("TextToVerify")="Approved~Approver: a400237~Confidential"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ViewerTabOperations","RAC_Common_ViewerTabOperations",OneIteration,"VerifyPDFText","",""
'! @Example 		dictViewerTabInfo("ImageName")="ViewerTabDatasetImage_Circle"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ViewerTabOperations","RAC_Common_ViewerTabOperations",OneIteration,"VerifyDatasetImageExist","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sPerspective,sInnerTabName,sObjectCheckOutOption, sPDFText
Dim dictItems,dictKeys
Dim objDefaultWindow
Dim aValue,aKey, aTextToVerify
Dim iCounter
Dim objShell, objClipboard, objInsightObject

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sInnerTabName = Parameter("sInnerTabName")
sObjectCheckOutOption = Parameter("sObjectCheckOutOption")

'Creating Object of Teamcenter main window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_DefaultWindow","")
'If dictViewerTabInfo("WindowName") = "TC_Visualization_Professional" Then
'	
'Else
	If sInnerTabName<>"" Then
		'Click on inner tab
		LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Viewer",sInnerTabName
	Else
		'Selecting Viewer tab
		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select","Viewer",""
	End If
	'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
'End If

Select Case Lcase(Cstr(sObjectCheckOutOption))
	Case "menu"
        'Perform CheckOut Operation by menu operation
		LoadAndRunAction "RAC_Common\RAC_Common_ObjectCheckOut","RAC_Common_ObjectCheckOut",OneIteration,"CheckOut","menu","","","","","",""
		sObjectCheckOutOption = True
	Case "viewertabtoolbar"
        'Perform CheckOut Operation by menu operation
		LoadAndRunAction "RAC_Common\RAC_Common_ObjectCheckOut","RAC_Common_ObjectCheckOut",OneIteration,"CheckOut","viewertabtoolbar","","","","","",""
		sObjectCheckOutOption = True
	Case Else
		sObjectCheckOutOption = False
End Select

'Capture business functionality start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Viewer Tab Operations",sAction,"","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ViewerTabOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify value from Viewer tab
	Case "Verify"
		'Taking Items & Keys from dictionary
		dictItems = dictViewerTabInfo.Items
		dictKeys = dictViewerTabInfo.Keys
		For iCounter=0 to dictViewerTabInfo.count-1
			Select Case dictKeys(iCounter)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name","Description"
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","label",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on Viewer tab","","","","","")
						Call Fn_ExitTest()
					End IF
					If objDefaultWindow.JavaStaticText("jstx_ViewerTabText").Exist(1)=False Then
						Call Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","label",dictKeys(iCounter))
					End If
					
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations", "Exist",objDefaultWindow.JavaEdit("jedt_ViewerTabEdit"),"","","") Then
						If dictItems(iCounter)="{BLANK}" Then								
							If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","getroproperty",objDefaultWindow.JavaEdit("jedt_ViewerTabEdit"),"","value","")="" Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property is empty\blank","","","","DONOTSYNC","")
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property is not empty\blank","","","","","")
								Call Fn_ExitTest()
							End If
						Else
							If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","getroproperty",objDefaultWindow.JavaEdit("jedt_ViewerTabEdit"),"","value","")=dictItems(iCounter) Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] value","","","","DONOTSYNC","")
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] value","","","","","")
								Call Fn_ExitTest()
							End If
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on Viewer tab","","","","","")
						Call Fn_ExitTest()
					End If				
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name_Text"
					aKey=Split(dictKeys(iCounter),"_")
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","label",aKey(0) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aKey(0)) & " ] property does not exist\available on Viewer tab","","","","","")
						Call Fn_ExitTest()
					End IF
					
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations", "Exist", objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","","") Then
						If Trim(objDefaultWindow.JavaStaticText("jstx_ViewerTabValue").GetROProperty("label"))=Trim(dictItems(iCounter)) Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aKey(0)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] value","","","","DONOTSYNC","")
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aKey(0)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] value","","","","","")
							Call Fn_ExitTest()
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aKey(0)) & " ] property does not exist\available on Viewer tab","","","","","")
						Call Fn_ExitTest()
					End If	
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Projects"
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","label",dictKeys(iCounter) &":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on Viewer tab","","","","","")
						Call Fn_ExitTest()
					End IF
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations", "Exist",objDefaultWindow.JavaList("jlst_ViewerTabList"),"","","") Then						
						If Trim(objDefaultWindow.JavaList("jlst_ViewerTabList").GetItem(0))=dictItems(iCounter) Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] value","","","","DONOTSYNC","")
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] value","","","","","")
							Call Fn_ExitTest()							
						End If						
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on Viewer tab","","","","","")
						Call Fn_ExitTest()
					End If
			End Select
		Next			
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'dictViewerTabInfo("PropertyName")="ID~Name~Description"
	Case "VerifyHeader"
		aValue=Split(dictViewerTabInfo("PropertyName"),"~")
		For iCounter=0 to Ubound(aValue)
			If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","label",aValue(iCounter))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aValue(iCounter)) & " ] property does not exist\available on Viewer tab","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations", "Exist", objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","","") Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aValue(iCounter)) & " ] property exist\available on Viewer tab","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aValue(iCounter)) & " ] property does not exist\available on Viewer tab","","","","","")
				Call Fn_ExitTest()
			End If
		Next									
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
	Case "Modify"
		'Taking Items & Keys from dictionary
		dictItems = dictViewerTabInfo.Items
		dictKeys = dictViewerTabInfo.Keys
		For iCounter=0 to dictViewerTabInfo.count-1
			Select Case dictKeys(iCounter)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name","Description"
					If  dictItems(iCounter)="{BLANK}"  Then
						dictItems(iCounter) = ""
					End If
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","label",dictKeys(iCounter) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on Viewer tab","","","","","")
						Call Fn_ExitTest()
					End IF
					
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations", "Exist",objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","","") Then
						If Fn_UI_JavaEdit_Operations("RAC_Common_ViewerTabOperations", "Set", objDefaultWindow, "jedt_ViewerTabEdit", dictItems(iCounter)) Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully modified [ " & Cstr(dictKeys(iCounter)) & " ] property value to [ " & Cstr(dictItems(iCounter)) & " ] from Viewer tab","","","","","")
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property value to [ " & Cstr(dictItems(iCounter)) & " ] from Viewer tab","","","","","")
							Call Fn_ExitTest()
							
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictKeys(iCounter)) & " ] property as property does not exist\available on Viewer tab","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End Select
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
	'dictViewerTabInfo("HyperlinkName")="More Properties..."
	Case "HyperlinkClick"
		If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","settoproperty",objDefaultWindow.JavaObject("jobj_ImageHyperlink"),"","text",dictViewerTabInfo("HyperlinkName"))=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(dictViewerTabInfo("HyperlinkName")) & " ] link as link does not exist\available on Viewer tab","","","","","")
			Call Fn_ExitTest()
		End IF
		If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations", "Exist",objDefaultWindow.JavaObject("jobj_ImageHyperlink"),"","","") Then
			If Fn_UI_JavaObject_Operations("RAC_Common_ViewerTabOperations","Click",objDefaultWindow,"jobj_ImageHyperlink","10","10","")=True Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on [ " & Cstr(dictViewerTabInfo("HyperlinkName")) & " ] link from Viewer tab","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(dictViewerTabInfo("HyperlinkName")) & " ] link from Viewer tab","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(dictViewerTabInfo("HyperlinkName")) & " ] link as link does not exist\available on Viewer tab","","","","","")
			Call Fn_ExitTest()
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
	'dictViewerTabInfo("Modified Weight")="g"
	Case "VerifyUnits"
		'Taking Items & Keys from dictionary
		dictItems = dictViewerTabInfo.Items
		dictKeys = dictViewerTabInfo.Keys
		For iCounter=0 to dictViewerTabInfo.count-1
			Select Case dictKeys(iCounter)
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Modified weight","Weight","Density"
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","settoproperty",objDefaultWindow.JavaStaticText("jstx_ViewerTabText"),"","label",dictKeys(iCounter))=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on Viewer tab","","","","","")
						Call Fn_ExitTest()
					End IF
					If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations", "Exist",objDefaultWindow.JavaStaticText("jstx_ViewerTabValue"),"","","") Then
						If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","getroproperty",objDefaultWindow.JavaStaticText("jstx_ViewerTabValue"),"","label","")=dictItems(iCounter) Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] as unit value","","","","DONOTSYNC","")
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] as unit value","","","","","")
							Call Fn_ExitTest()
						End If
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] as unit value","","","","","")
						Call Fn_ExitTest()
					End If
			End Select
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
	'Case to verify PDF text displayed in Viewer tab
	Case "VerifyPDFText"
		
		JavaWindow("jwnd_DefaultWindow").JavaObject("jobj_ViewerTabObject").Click 50,50
		wait(GBL_MIN_MICRO_TIMEOUT)
		
		'Create object of Shell scripting
		Set objShell = CreateObject("WScript.Shell")
		
		'Create object of mercury clipboard
		Set objClipboard = CreateObject("Mercury.Clipboard")
		
		'Select all and copy the pdf content
	    objShell.SendKeys "^(a)"
	    wait(GBL_MIN_MICRO_TIMEOUT)
	    objShell.SendKeys "^(c)"
	    
	    'Get data from clipboard in variable
	    sPDFText = objClipboard.GetText
	    
	    'Verify each text value is present in PDF
	    aTextToVerify = Split(dictViewerTabInfo("TextToVerify"), "~")
		For iCounter = 0 To Ubound(aTextToVerify) Step 1
			If Instr(sPDFText, aTextToVerify(iCounter)) > 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification passed as [ " & Cstr(aTextToVerify(iCounter)) & " ] was found in viewer tab text [" & sPDFText & "]","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aTextToVerify(iCounter)) & " ] was not found in viewer tab text [" & sPDFText & "]","","","","","")
				Call Fn_ExitTest()
			End If
		Next
		Set objShell = Nothing
		Set objClipboard = Nothing
		
	Case "VerifyDatasetImageExist"
	
		Set objInsightObject = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_" & dictViewerTabInfo("ImageName"),"")
		
'		If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","Exist",Window("wwnd_TeamcenterVisualization"),"","","")=False Then
'			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify dataset image as Visualization window does not exist","","","","","")
'			Call Fn_ExitTest()
'		End If
	
'		If dictViewerTabInfo("WindowName") = "TC_Visualization_Professional" Then
'			Window("wwnd_TeamcenterVisualization").Activate
'			Wait 3
'		End If
		
		If Fn_UI_Object_Operations("RAC_Common_ViewerTabOperations","Exist", objInsightObject,"", "", "") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as image [ " & Cstr(dictViewerTabInfo("ImageName")) & " ] was not found under Viewer tab","","","","DONOTSYNC","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as image [ " & Cstr(dictViewerTabInfo("ImageName")) & " ] was found under Viewer tab","","","","","")
		End If
		
'		If dictViewerTabInfo("CloseWindow") = True Then
'			Window("wwnd_TeamcenterVisualization").Close
'		End If

End Select			

dictViewerTabInfo.RemoveAll

'Capture business functionality end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Viewer Tab Operations",sAction,"","")

If sObjectCheckOutOption = True Then
	LoadAndRunAction "RAC_Common\RAC_Common_ObjectCheckIn","RAC_Common_ObjectCheckIn",OneIteration,"CheckIn","menu",""
End If

'Added code to handle VFFrame dialog- 8-Sept-2017 - Sandeep Navghane
LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select","Summary",""
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'Releasing teamcenter main window object
Set objDefaultWindow=Nothing

Function Fn_ExitTest()
	'Releasing teamcenter main window object
	Set objDefaultWindow=Nothing
	ExitTest
End Function

