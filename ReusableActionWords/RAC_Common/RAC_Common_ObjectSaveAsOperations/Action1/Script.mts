'! @Name 			RAC_Common_ObjectSaveAsOperations
'! @Details 		Action word to perform operations on Save As dialog
'! @InputParam1 	sAction 		: Action to be performed e.g. Autosave asBasic
'! @InputParam2 	sInvokeOption 	: Method to invoke save as dialog e.g. menu
'! @InputParam3 	sButton		 	: Button Name
'! @InputParam4 	dictSaveAsInfo 	: External parameter to pass save as object additional information
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			13 Dec 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ObjectSaveAsOperations","RAC_Common_ObjectSaveAsOperations",OneIteration,"AutoSaveAsItemBasic","Menu",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ObjectSaveAsOperations","RAC_Common_ObjectSaveAsOperations",OneIteration,"AutoSaveAsAndCopyToClipboard","Menu",""
'! @Example 		dictSaveAsInfo("Description") = "asa"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ObjectSaveAsOperations","RAC_Common_ObjectSaveAsOperations",OneIteration,"verifypropertyvalues","Menu",""

Option Explicit
Err.Clear

'Declaring varaibles	
Dim sAction,sInvokeOption,sButton
Dim sID,sName,sPerspective,sObjectInformation
Dim iSaveAsObjectCount
Dim objSaveAs
Dim dictItems, dictKeys
Dim aNode,aObjectType,aObjectTypeValue,aObjectName,aCopyOption,aColumnName,aColumnValue
Dim iCounter,iPath,iInstanceHandler,iItemCount,iCount
Dim iY,iX,iColumnWidth,iNodeIndex,iItemHeight
Dim objDescription,objChildObjects
Dim sChildNodePath
Dim sNodePath
Dim bFlag
Dim aName
Dim iNameCount
Dim sNode

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction=Parameter("sAction")
sInvokeOption=Parameter("sInvokeOption")
sButton = Parameter("sButton")

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating object of save as dialog
Select Case LCase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter","","structuremanager"
		Set objSaveAs=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_SaveAs","")
End Select

'inoke save as dialog
Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileSaveAs"
		'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "summarytablink"
		LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Summary","Overview"		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
		JavaWindow("jwnd_DefaultWindow").JavaObject("to_class:=JavaObject","text:=Save As").Click 2,2,"LEFT"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'Use this invoke option when user wants to invoke save as dialog from outside function
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ObjectSaveAsOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of save as dialog
If Fn_UI_Object_Operations("RAC_Common_ObjectSaveAsOperations", "Exist", objSaveAs, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ save as ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

objSaveAs.maximize
'Setting save as object count
If Lcase(sAction)= Lcase("autosaveasitembasic") or sAction = "AutoSaveAsAndCopyToClipboard" or Lcase(sAction) ="saveasitemdetails" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectCount","","")
	iSaveAsObjectCount=Fn_CommonUtil_DataTableOperations("GetValue","RACSaveAsObjectCount","","")
	If iSaveAsObjectCount="" Then
		iSaveAsObjectCount=1
	Else
		iSaveAsObjectCount=iSaveAsObjectCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACSaveAsObjectCount",iSaveAsObjectCount,"")
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Object save as Operations",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to save as obejct with auto populated values
	Case "autosaveasitembasic"
		'Setting revision		
		If Lcase(sAction)="autosaveasitembasic" Then
			IF Fn_UI_Object_Operations("RAC_Common_ObjectSaveAsOperations","GetROProperty",objSaveAs.JavaButton("jbtn_Assign"),"","enabled","")=1 Then
				If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations", "Click", objSaveAs,"jbtn_Assign")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object as fail to click on [ Assign ] button to assign new id","","","","","")	
					Call Fn_ExitTest()
				End If			
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End IF
		End If
		
		'Getting New ID
		sID=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectSaveAsOperations", "GetText",objSaveAs,"jedt_ID", "")
		sObjectInformation=Fn_UI_Object_Operations("RAC_Common_ObjectSaveAsOperations","GetROProperty",objSaveAs.JavaStaticText("jstx_ObjectInformation"),"","label","")
		sObjectInformation=Split(sObjectInformation,"-")
		sName=sObjectInformation(1)
				
		'Clicking on Finish Button
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations","click",objSaveAs,"jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object as fail to click on [ Finish ] button","","","","","")	
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		DataTable.SetCurrentRow iSaveAsObjectCount
		
		'Setting ID in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectID","","")					
		DataTable.Value("RACSaveAsObjectID","Global")= sID
		
		'Setting Revision Name in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectName","","")					
		DataTable.Value("RACSaveAsObjectName","Global")= sName	
		
		'Setting save asObject Revision Node value in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectNode","","")
		DataTable.Value("RACSaveAsObjectNode","Global")=GBL_NEWSTUFF_FOLDER_PATH & "~" & DataTable.Value("RACSaveAsObjectID","Global") & "-" & DataTable.Value("RACSaveAsObjectName","Global")
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object save as Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully save as selected object to Id [ " & Cstr(Datatable.Value("RACSaveAsObjectID", "Global")) & " ]","","","","DONOTSYNC","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to click on specific button of save as dialog
	Case "ClickButton"
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations", "Click", objSaveAs,sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of save as dialog","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object save as Operations",sAction,"Button Name",sButton)
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully click on [ " & Cstr(sButton) & " ] button of save as dialog","","","","","")
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to save as obejct with auto populated values, copy the new revision to clipboard and return the newly saved value
	Case "autosaveasandcopytoclipboard"
		'Select Copy to Clipboard checkbox
		If Fn_UI_JavaCheckBox_Operations("RAC_Common_ObjectSaveAsOperations", "Set", objSaveAs, "jckbCopyToClipboard", "ON") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Copy To Clipboard ] chekbox of save as dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on Finish button
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations", "Click", objSaveAs,"jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Finish ] button of save as dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(5)
		
		'Get the saved as item details from clipboard
		LoadAndRunAction "RAC_Common\RAC_Common_ClipboardOperations","RAC_Common_ClipboardOperations", oneIteration, "GetClipBoardContents","",""
		Datatable.SetCurrentRow 1
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ObjectSaveAsOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		sID = DataTable.Value("ReusableActionWordReturnValue","Global")
		DataTable.SetCurrentRow iSaveAsObjectCount
		
		'Setting ID in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectID","","")					
		DataTable.Value("RACSaveAsObjectID","Global") = Split(sID,"-")(0)
		
		'Setting Revision Name in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectName","","")					
		DataTable.Value("RACSaveAsObjectName","Global")= Split(sID,"-")(1)	
		
		If Lcase(sAction) = "autosaverawmaterialrevisionas" or Lcase(sAction) ="autosaveengineeredpartrevisionas" Then
			'Setting Revision ID in datatable
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectRevisionID","","")					
			DataTable.Value("RACSaveAsObjectRevisionID","Global")="AA"
		End If
		
		'Setting save asObject Revision Node value in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectNode","","")
		DataTable.Value("RACSaveAsObjectNode","Global") = GBL_NEWSTUFF_FOLDER_PATH & "~" & sID
		
		If Lcase(sAction) = "autosaverawmaterialrevisionas" or Lcase(sAction) ="autosaveengineeredpartrevisionas" Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectRevisionNode","","")
			DataTable.Value("RACSaveAsObjectRevisionNode","Global") = DataTable.Value("RACSaveAsObjectNode","Global") & "~" & DataTable.Value("RACSaveAsObjectID","Global") & "/" & DataTable.Value("RACSaveAsObjectRevisionID","Global") & "-" & DataTable.Value("RACSaveAsObjectName","Global")
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object save as Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully save as selected object to Id [ " & Cstr(Datatable.Value("RACSaveAsObjectID", "Global")) & " ]","","","","","")
		End If
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER	
	Case "verifypropertyvalues"
			dictItems = dictSaveAsInfo.Items
			dictKeys = dictSaveAsInfo.Keys
			For iCounter=0 to dictSaveAsInfo.count-1
				Select Case dictKeys(iCounter)
					
					Case "Description", "Is Serviceable", "Customer Master ID", "Customer", "Legacy OEM Name", "Customer Part Number", "Customer Part Name" , "Customer Part Revision"
						'Set the label property of static text or field label
						If Fn_UI_Object_Operations("RAC_Common_ObjectSavesOperations","settoproperty",objSaveAs.JavaStaticText("jstx_SaveAsLabel"),"","label",dictKeys(iCounter) & ":")=False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on save as dialog","","","","","")
							Call Fn_ExitTest()
						End IF
						
						'Verify existence of edit box
						If Fn_UI_Object_Operations("RAC_Common_ObjectSavesOperations", "Exist",objSaveAs.JavaEdit("jedt_SaveAsEdit"),"","","") Then
							If dictItems(iCounter)="{BLANK}" Then								
								If Fn_UI_Object_Operations("RAC_Common_ObjectSavesOperations","getroproperty",objSaveAs.JavaEdit("jedt_SaveAsEdit"),"","value","")="" Then
									Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property is empty\blank","","","","DONOTSYNC","")
								Else
									Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property is not empty\blank","","","","","")
									Call Fn_ExitTest()
								End If
							Else
								If Fn_UI_Object_Operations("RAC_Common_ObjectSavesOperations","getroproperty",objSaveAs.JavaEdit("jedt_SaveAsEdit"),"","value","")=dictItems(iCounter) Then
									Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictKeys(iCounter)) & " ] property contains [ " & Cstr(dictItems(iCounter)) & " ] value","","","","DONOTSYNC","")
								Else
									Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not contain [ " & Cstr(dictItems(iCounter)) & " ] value","","","","","")
									Call Fn_ExitTest()
								End If
							End If
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictKeys(iCounter)) & " ] property does not exist\available on save as dialog","","","","","")
							Call Fn_ExitTest()
						End If
					End Select
			Next
			
			If sButton <> "" Then
				If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations","click",objSaveAs,"jbtn_" & sButton)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties on save as dialog as fail to click on [ Finish ] button","","","","","")	
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to save as obejct with detail information
	Case "saveasitemdetails"
		'Setting revision		
		If Lcase(sAction)="saveasitemdetails" Then
			IF Fn_UI_Object_Operations("RAC_Common_ObjectSaveAsOperations","GetROProperty",objSaveAs.JavaButton("jbtn_Assign"),"","enabled","")=1 Then
				If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations", "Click", objSaveAs,"jbtn_Assign")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object as fail to click on [ Assign ] button to assign new id","","","","","")	
					Call Fn_ExitTest()
				End If			
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			End IF
		End If
		
		If Lcase(sAction)="saveasitemdetails" Then
			'Getting New ID
			sID=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectSaveAsOperations", "GetText",objSaveAs,"jedt_ID", "")
			sObjectInformation=Fn_UI_Object_Operations("RAC_Common_ObjectSaveAsOperations","GetROProperty",objSaveAs.JavaStaticText("jstx_ObjectInformation"),"","label","")
			sObjectInformation=Split(sObjectInformation,"-")
			sName=sObjectInformation(1)
		End If
		
		If dictSaveAsInfo("DefineAttachedObjects")=True Then
			'Clicking on Next Button
			If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations","click",objSaveAs,"jbtn_Next")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object as fail to click on [ Next ] button","","","","","")	
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			
			aNode=Split(dictSaveAsInfo("DefineAttachedObjectsParentNode"),"^")
			aObjectType=Split(dictSaveAsInfo("ObjectType"),"^")
			aObjectTypeValue=Split(dictSaveAsInfo("ObjectTypeValue"),"^")
			aObjectName=Split(dictSaveAsInfo("ObjectName"),"^")
			aCopyOption=Split(dictSaveAsInfo("CopyOption"),"^")
			
			For iCounter=0 to Ubound(aNode)
				iPath = Fn_RAC_GetJavaTreeNodePath(objSaveAs.JavaTree("jtree_DefineAttachedObjects"), aNode(iCounter) , "~", "@")
				If iPath=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as Node [" & Cstr(aNode(iCounter)) & "] does not exist in Define Attached Objects tree on SaveAs dialog","","","","","")
					Call Fn_ExitTest()
				End If
				
				If aObjectType(iCounter)="Dataset" Then
					aObjectTypeValue(iCounter) = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewDatasetValues_APL",aObjectTypeValue(iCounter),"")
				ElseIf aObjectType(iCounter)="Form" Then
					aObjectTypeValue(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewFormValues_APL",aObjectTypeValue(iCounter),"")
				Else
					aObjectTypeValue(iCounter)="False"
				End If
				
				If Cstr(aObjectTypeValue(iCounter))="False" Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as invalid Object Type [ " & Cstr(aObjectType(iCounter)) & " ] passed to SaveAs action word","","","","","")
					Call Fn_ExitTest()
				End If
				
				iInstanceHandler=1		
				iItemCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
				sChildNodePath=aNode(iCounter) & "~" & aObjectName(iCounter)
				bFlag=False
				For iCount=0 to iItemCount
					sNodePath = sChildNodePath & "@" & iInstanceHandler
					iPath = Fn_RAC_GetJavaTreeNodePath(objSaveAs.JavaTree("jtree_DefineAttachedObjects"),sNodePath , "~", "@")
					If iPath=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as Node [" & Cstr(sNodePath) & "] does not exist in Define Attached Objects tree on SaveAs dialog","","","","","")
						Call Fn_ExitTest()
					End If
					If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(aObjectTypeValue(iCounter)) Then
						bFlag = True
						Exit For
					End If
					iInstanceHandler=iInstanceHandler+1
				Next
				If bFlag = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as [ " & Cstr(aObjectType(iCounter)) & " ] [ " & Cstr(aObjectName(iCounter)) & " ] of Type [ " & Cstr(aObjectTypeValue(iCounter)) & " ] does not exist in Define Attached Objects tree on SaveAs dialog","","","","","")
					Call Fn_ExitTest()
				End If
				
				iY = 0
				iX = 0					
				For iCount = 0 To 2
					iColumnWidth=objSaveAs.JavaTree("jtree_DefineAttachedObjects").Object.getColumn(1).getWidth()
					iX = iX + iColumnWidth
				Next
				iX = iX - iColumnWidth/2
				
				iNodeIndex=Fn_RAC_GetJavaTreeNodeIndex(objSaveAs.JavaTree("jtree_DefineAttachedObjects"),sNodePath,"","")
				For iCount = 0 to iNodeIndex
					iItemHeight = objSaveAs.JavaTree("jtree_DefineAttachedObjects").Object.getItemHeight()
					iY = iY + iItemHeight
				Next
				iY = iY - iItemHeight/2
				
				objSaveAs.JavaTree("jtree_DefineAttachedObjects").Click iX, iY,"LEFT"
				wait 1
				objSaveAs.JavaTree("jtree_DefineAttachedObjects").Click iX, iY,"LEFT"
				wait 1
				Set objDescription=Description.Create()
				objDescription("Class Name").value = "JavaList"
				objDescription("toolkit class").value = "org.eclipse.swt.custom.CCombo"
				objDescription("tagname").value = "CCombo"
				Set objChildObjects = objSaveAs.ChildObjects(objDescription)
				For iCount = 0 to objChildObjects.count-1
					objChildObjects(iCount).Select aCopyOption(iCounter)
					Call Fn_CommonUtil_KeyBoardOperation("SendKeys", "{TAB}")
				Next
				Set objChildObjects =Nothing
				Set objDescription=Nothing
				If Err.Number<>0 then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
			Next
		End If
		
		'Clicking on Finish Button
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations","click",objSaveAs,"jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object as fail to click on [ Finish ] button","","","","","")	
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		DataTable.SetCurrentRow iSaveAsObjectCount
		
		
		'Setting ID in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectID","","")					
		DataTable.Value("RACSaveAsObjectID","Global")= sID
		
		'Setting Revision Name in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectName","","")					
		DataTable.Value("RACSaveAsObjectName","Global")= sName
		
		'Setting save asObject Revision Node value in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectNode","","")
		DataTable.Value("RACSaveAsObjectNode","Global")=GBL_NEWSTUFF_FOLDER_PATH & "~" & DataTable.Value("RACSaveAsObjectID","Global") & "-" & DataTable.Value("RACSaveAsObjectName","Global")
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object save as Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully save as selected object to Id [ " & Cstr(Datatable.Value("RACSaveAsObjectID", "Global")) & " ]","","","","DONOTSYNC","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "verifydefineattachedobjects"				
		'Clicking on Next Button
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations","click",objSaveAs,"jbtn_Next")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object as fail to click on [ Next ] button","","","","","")	
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		aNode=Split(dictSaveAsInfo("DefineAttachedObjectsParentNode"),"^")
		aObjectType=Split(dictSaveAsInfo("ObjectType"),"^")
		aObjectTypeValue=Split(dictSaveAsInfo("ObjectTypeValue"),"^")
		aObjectName=Split(dictSaveAsInfo("ObjectName"),"^")
		If dictSaveAsInfo("ColumnName")<>"" Then
			aColumnName=Split(dictSaveAsInfo("ColumnName"),"^")
			aColumnValue=Split(dictSaveAsInfo("ColumnValue"),"^")
		End If
		For iCounter=0 to Ubound(aNode)
			iPath = Fn_RAC_GetJavaTreeNodePath(objSaveAs.JavaTree("jtree_DefineAttachedObjects"), aNode(iCounter) , "~", "@")
			If iPath=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as Node [" & Cstr(aNode(iCounter)) & "] does not exist in Define Attached Objects tree on SaveAs dialog","","","","","")
				Call Fn_ExitTest()
			End If
			
			If aObjectType(iCounter)="Dataset" Then
				aObjectTypeValue(iCounter) = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewDatasetValues_APL",aObjectTypeValue(iCounter),"")
			ElseIf aObjectType(iCounter)="Form" Then
				aObjectTypeValue(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewFormValues_APL",aObjectTypeValue(iCounter),"")
			Else
				aObjectTypeValue(iCounter)="False"
			End If
			
			If Cstr(aObjectTypeValue(iCounter))="False" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as invalid Object Type [ " & Cstr(aObjectType(iCounter)) & " ] passed to SaveAs action word","","","","","")
				Call Fn_ExitTest()
			End If
			
			iInstanceHandler=1		
			iItemCount=Cint(GBL_JAVATREE_CURRENTNODE_OBJECT.getItemCount())-1
			sChildNodePath=aNode(iCounter) & "~" & aObjectName(iCounter)
			bFlag=False
			For iCount=0 to iItemCount-1
				sNodePath = sChildNodePath & "@" & iInstanceHandler
				iPath = Fn_RAC_GetJavaTreeNodePath(objSaveAs.JavaTree("jtree_DefineAttachedObjects"),sNodePath , "~", "@")
				If iPath=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as Node [" & Cstr(sNodePath) & "] does not exist in Define Attached Objects tree on SaveAs dialog","","","","","")
					Call Fn_ExitTest()
				End If
				If lCase(GBL_JAVATREE_CURRENTNODE_OBJECT.getData().getComponent().getType())=lcase(aObjectTypeValue(iCounter)) Then
					bFlag = True
					Exit For
				End If
				iInstanceHandler=iInstanceHandler+1
			Next
			If bFlag = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as [ " & Cstr(aObjectType(iCounter)) & " ] [ " & Cstr(aObjectName(iCounter)) & " ] of Type [ " & Cstr(aObjectTypeValue(iCounter)) & " ] does not exist in Define Attached Objects tree on SaveAs dialog","","","","","")
				Call Fn_ExitTest()
			End If
			
			If dictSaveAsInfo("ColumnName")<>"" Then
				If aColumnValue(iCounter)=objSaveAs.JavaTree("jtree_DefineAttachedObjects").GetColumnValue(iPath,aColumnName(iCounter)) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aObjectType(iCounter)) & " ] [ " & Cstr(sNodePath) & " ] contains [ " & Cstr(aColumnValue(iCounter)) & " ] value under column [ " & Cstr(aColumnName(iCounter)) & " ] under [ DefineAttachedObjects ] on SaveAs dialog","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aObjectType(iCounter)) & " ] [ " & Cstr(sNodePath) & " ] does not contains [ " & Cstr(aColumnValue(iCounter)) & " ] value under column [ " & Cstr(aColumnName(iCounter)) & " ] under [ DefineAttachedObjects ] on SaveAs dialog","","","","","")
					Call Fn_ExitTest()
				End IF
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aObjectType(iCounter)) & " ] [ " & Cstr(sNodePath) & " ] exist under [ DefineAttachedObjects ] on SaveAs dialog","","","","DONOTSYNC","")	
			End If
			If Err.Number<>0 then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to complete operation [ " & Cstr(sAction) & " ] selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		Next
		If sButton <> "" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations","click",objSaveAs,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to complete operation [ " & Cstr(sAction) & " ] on save as dialog as fail to click on [ " & Cstr(sButton) & " ] button","","","","","")	
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		
	'Case to save as object revision after clicking on Open on Create checkbox
	Case "autosavepartrevisionas_openoncreate"
		'Select Copy to Clipboard checkbox
		If Fn_UI_JavaCheckBox_Operations("RAC_Common_ObjectSaveAsOperations", "Set", objSaveAs, "jckbOpenOnCreate", "ON") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Open On Create ] chekbox of save as dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		If Fn_UI_JavaCheckBox_Operations("RAC_Common_ObjectSaveAsOperations", "Set", objSaveAs, "jckbCopyToClipboard", "OFF") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ Copy To Clipboard ] chekbox of save as dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on Finish button
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations", "Click", objSaveAs,"jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Finish ] button of save as dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Get the node name from opened item 
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "GetFirstNodeName", "",""
		Datatable.SetCurrentRow 1
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ObjectSaveAsOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		sID = Split(DataTable.Value("ReusableActionWordReturnValue","Global"),"-")
		
		DataTable.SetCurrentRow iSaveAsObjectCount
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectID","","")					
		DataTable.Value("RACSaveAsObjectID","Global") = sID(0)
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectNode","","")	
		DataTable.Value("RACSaveAsObjectNode","Global") = DataTable.Value("ReusableActionWordReturnValue","Global")
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectName","","")					
		DataTable.Value("RACSaveAsObjectName","Global")= sID(1)
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectRevisionID","","")					
		DataTable.Value("RACSaveAsObjectRevisionID","Global")="AA"
		
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "GetChildrenByName", DataTable.Value("RACSaveAsObjectNode","Global") & "^" & sID(0),""
		Datatable.SetCurrentRow 1
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ObjectSaveAsOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		sID = DataTable.Value("ReusableActionWordReturnValue","Global")
		
		DataTable.SetCurrentRow iSaveAsObjectCount
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACSaveAsObjectRevisionNode","","")
		DataTable.Value("RACSaveAsObjectRevisionNode","Global") = DataTable.Value("ReusableActionWordReturnValue","Global")
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object save as Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to save as selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully save as selected object to Id [ " & Cstr(Datatable.Value("RACSaveAsObjectID", "Global")) & " ]","","","","","")
		End If
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "autopopulatesaveasandverifyfinishbuttonenabled"
		'Click on Finish button
		If Cint(objSaveAs.JavaButton("jbtn_Finish").GetROProperty("enabled"))=0  Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as [ Finish ] button is disabled on Save As dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sButton <> "" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_ObjectSaveAsOperations","click",objSaveAs,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of [ Save As ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object save as Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify [ Finish ] button is enabled on [ Save As ] dialog due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ Finish ] button is enabled on [ Save As ] dialog","","","","","")
		End If		
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing object
Set objSaveAs=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objSaveAs=Nothing
	ExitTest
End Function
