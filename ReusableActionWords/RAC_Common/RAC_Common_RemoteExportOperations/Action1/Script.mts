'! @Name 			RAC_Common_RemoteExportOperations
'! @Details 		Action word to perform remote export opertations
'! @InputParam1 	sAction 						: String to indicate what action is to be performed
'! @InputParam2 	sInvokeOption					: Export To Excel dialog invoke option
'! @InputParam3 	sButton 						: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Mohini Deshmukh Mohini.Deshmukh@sqs.com
'! @Date 			22 Aug 2017
'! @Version 		1.0
'! @Example 		dictRemoteExportInfo("TargetSites")="YFJC_QTEST11"
'! @Example 		dictRemoteExportInfo("RemoteExportOptions")="true"
'! @Example 		dictRemoteExportInfo("IncludeAllVersions")="ON"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_RemoteExportOperations","RAC_Common_RemoteExportOperations",OneIteration,"Export","Menu",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sButton
Dim objOptionsSettings,objRemoteExport,objRemoteExportOptions,objRemoteSiteSelection,objOrganizationSelection
Dim sOptionSettingsHeader,sOptionSettingsValue
Dim aSites,aTemp
Dim iCounter
Dim sNewOwningUser,sNode
Dim aNode

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sButton = Parameter("sButton")

'Invoking [ Remote ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ToolsExportRemote"
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Creating object of [ Remote Export ] dialog
Set objRemoteExport =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_RemoteExport","")
Set objRemoteSiteSelection =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_RemoteSiteSelection","")
Set objRemoteExportOptions =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_RemoteExportOptions","")
Set objOptionsSettings =Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_OptionsSettings","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_RemoteExportOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of [  Remote ] dialog
If Fn_UI_Object_Operations("RAC_Common_RemoteExportOperations","Exist", objRemoteExport,"","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Remote Export ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_RemoteExportOperations",sAction,"","")

Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Remote
	Case "Export", "ExportAndVerifyOptionSettings","ExportToSpecificOwner"
		If dictRemoteExportInfo("TargetSites")<>"" Then
			'Click on Select Target Remote Site button
			If Fn_UI_JavaButton_Operations("RAC_Common_RemoteExportOperations","Click",objRemoteExport,"jbtn_SelectTargetRemoteSite")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to click on [ Select Target Remote Site ] button from Remote Export dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			
			aSites=Split(dictRemoteExportInfo("TargetSites"),"~")
			For iCounter=0 to ubound(aSites)
				If Fn_UI_JavaList_Operations("RAC_Common_RemoteExportOperations", "Select", objRemoteSiteSelection,"jlst_AvailableSite",aSites(iCounter),"", "")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to select target site [ " & Cstr(aSites(iCounter)) & " ] from Remote Site Selection dialog","","","","","")
					Call Fn_ExitTest()
				End If
				'Click on Add button
				If Fn_UI_JavaButton_Operations("RAC_Common_RemoteExportOperations","Click",objRemoteSiteSelection,"jbtn_Add")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to click on [ Add ] button from Remote Site Selection dialog","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			Next
			
			'Click on OK button
			If Fn_UI_JavaButton_Operations("RAC_Common_RemoteExportOperations","Click",objRemoteSiteSelection,"jbtn_OK")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to click on [ OK ] button from Remote Site Selection dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
		
		If LCase(dictRemoteExportInfo("RemoteExportOptions"))="true" Then
			'Click on Set Remote Export Options button
			If Fn_UI_JavaButton_Operations("RAC_Common_RemoteExportOperations","Click",objRemoteExport,"jbtn_SetRemoteExportOptions")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to click on [ Set Remote Export Options ] button from Remote Export dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			
			sOptionSettingsHeader=""
			sOptionSettingsValue=""
			
			If dictRemoteExportInfo("TransferOwnership")<>"" Then
				sOptionSettingsHeader="Transfer Options"
				sOptionSettingsValue="Transfer ownership"
				
				objRemoteExportOptions.JavaCheckBox("jckb_ExportOption").SetTOProperty "attached text","Transfer ownership"
				If Fn_UI_JavaCheckBox_Operations("RAC_Common_RemoteExportOperations", "Set", objRemoteExportOptions, "jckb_ExportOption", "ON") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to select option [ Transfer ownership ] from Remote Export Options dialog","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
				If sAction="ExportToSpecificOwner" Then
					If dictRemoteExportInfo("NewOwningUser")<>"" Then
						objRemoteExportOptions.JavaTab("jtab_ExportOption").Select "Advanced"
						Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
						If Fn_UI_JavaButton_Operations("RAC_Common_RemoteExportOperations","Click",objRemoteExportOptions,"jbtn_SelectNewOwningUser")=False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to click on [ Select New Owning User ] button from Remote Export Options dialog","","","","","")
							Call Fn_ExitTest()
						End If
						Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
						
						Set objOrganizationSelection=JavaWindow("jwnd_DefaultWindow").JavaWindow("jwnd_TcDefaultApplet").JavaDialog("jdlg_OrganizationSelection")
						'Checking existance 
						If objOrganizationSelection.Exist(20)=False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to select new owning user from [ Organization Selection ] dialog as [ Organization Selection ] dialog does not exist after clicking on [ Select New Owning User ] button","","","","","")
							Call Fn_ExitTest()
						End If
						'Fetching New Owning user details
						sNewOwningUser= Fn_Setup_GetTestUserDetailsFromExcelOperations("getorganizationtreeusernodepath","",dictRemoteExportInfo("NewOwningUser"))
						
						'Selecting New Owning User from Organization chart
						aNode = Split(sNewOwningUser,"~")
						If Ubound(aNode) > 1 Then
							For iCounter = 0 to Ubound(aNode) - 1
								If iCounter = 0 Then
									sNode = aNode(0)
								Else
									sNode = sNode & "~" & aNode(iCounter)
								End If
								'expanding node
								Call Fn_UI_JavaTree_Operations("RAC_Common_RemoteExportOperations", "Expand",objOrganizationSelection, "jtree_Organization",sNode,"","")
								Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)					
							Next
						End If
							
						If Fn_UI_JavaTree_Operations("RAC_Common_RemoteExportOperations", "Select",objOrganizationSelection, "jtree_Organization",sNewOwningUser,"","")=False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Change Ownership of selected object as fail to select organization tree node [ " & Cstr(sNewOwningUser) & " ] on Change Ownership dialog","","","","","")
							Call Fn_ExitTest()
						End IF
						Call Fn_RAC_ReadyStatusSync(GBL_MAX_SYNC_ITERATIONS)
						wait 3
						
						For iCounter = 1 To 10 Step 1
							'If Fn_UI_Object_Operations("RAC_Common_RemoteExportOperations", "Enabled", objOrganizationSelection.JavaButton("jbtn_OK"),"", "", "") = False Then
							If objOrganizationSelection.JavaButton("jbtn_OK").GetROProperty("enabled") = "0"  OR objOrganizationSelection.JavaButton("jbtn_OK").GetROProperty("enabled") = False Then
								wait 1
							End If
						Next
						
						'Click on OK button
						If Fn_UI_JavaButton_Operations("RAC_Common_RemoteExportOperations", "Click", objOrganizationSelection,"jbtn_OK") = False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to change ownership of selected object as fail to click on [ OK ] button of Organization Selection dialog","","","","","")
							Call Fn_ExitTest()
						End If
						Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
						
						objRemoteExportOptions.JavaTab("jtab_ExportOption").Select "General"
						Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
					End If
				End If
			End If
						
			If dictRemoteExportInfo("UseDefaultUserGroupOwnershipRules")=True Then
				objRemoteExportOptions.JavaTab("jtab_ExportOption").Select "Advanced"
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
				
				objRemoteExportOptions.JavaCheckBox("jckb_ExportOption").SetTOProperty "attached text","Use default user/group ownership rules"
				If Fn_UI_JavaCheckBox_Operations("RAC_Common_RemoteExportOperations", "Set", objRemoteExportOptions, "jckb_ExportOption", "ON") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to select option [ Use default user/group ownership rules ] from Remote Export Options dialog","","","","","")
					Call Fn_ExitTest()
				End If
				
				objRemoteExportOptions.JavaTab("jtab_ExportOption").Select "General"
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			End If
			
			If dictRemoteExportInfo("IncludeAllFiles")<>"" Then
				If sOptionSettingsHeader="" Then
					sOptionSettingsHeader="Dataset Options"
					sOptionSettingsValue="Include all files"
				Else
					sOptionSettingsHeader=sOptionSettingsHeader & "^Dataset Options"
					sOptionSettingsValue=sOptionSettingsValue & "^Include all files"	
				End If
				
				objRemoteExportOptions.JavaCheckBox("jckb_ExportOption").SetTOProperty "attached text","Include all files"
				If Fn_UI_JavaCheckBox_Operations("RAC_Common_RemoteExportOperations", "Set", objRemoteExportOptions, "jckb_ExportOption", "ON") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to select option [ Include all files ] from Remote Export Options dialog","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			End If
			
			If dictRemoteExportInfo("IncludeAllVersions")<>"" Then
				If sOptionSettingsHeader="" Then
					sOptionSettingsHeader="Dataset Options"
					sOptionSettingsValue="Include all versions"
				Else
					sOptionSettingsHeader=sOptionSettingsHeader & "^Dataset Options"
					sOptionSettingsValue=sOptionSettingsValue & "^Include all versions"					
				End If
				
				objRemoteExportOptions.JavaCheckBox("jckb_ExportOption").SetTOProperty "attached text","Include all versions"
				If Fn_UI_JavaCheckBox_Operations("RAC_Common_RemoteExportOperations", "Set", objRemoteExportOptions, "jckb_ExportOption", "ON") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to select option [ Include all versions ] from Remote Export Options dialog","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			End If
			
			If dictRemoteExportInfo("IncludeEntireBOM")<>"" Then
				If sOptionSettingsHeader="" Then
					sOptionSettingsHeader="Structure Manager Options"
					sOptionSettingsValue="Include entire BOM"
				Else
					sOptionSettingsHeader=sOptionSettingsHeader & "^Structure Manager Options"		
					sOptionSettingsValue=sOptionSettingsValue & "^Include entire BOM"							
				End If
				
				objRemoteExportOptions.JavaCheckBox("jckb_ExportOption").SetTOProperty "attached text","Include entire BOM"
				If Fn_UI_JavaCheckBox_Operations("RAC_Common_RemoteExportOperations", "Set", objRemoteExportOptions, "jckb_ExportOption", "ON") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to select option [ Include entire BOM ] from Remote Export Options dialog","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			End If
			
			If dictRemoteExportInfo("IncludeAllRevisions")<>"" Then
				If sOptionSettingsHeader="" Then
					sOptionSettingsHeader="Item Options"
					sOptionSettingsValue="Include all revisions"
				Else
					sOptionSettingsHeader=sOptionSettingsHeader & "^Item Options"					
					sOptionSettingsValue=sOptionSettingsValue & "^Include all revisions"	
				End If
				
				objRemoteExportOptions.JavaCheckBox("jckb_ExportOption").SetTOProperty "attached text","Include all revisions"
				If Fn_UI_JavaCheckBox_Operations("RAC_Common_RemoteExportOperations", "Set", objRemoteExportOptions, "jckb_ExportOption", "ON") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to select option [ Include all revisions ] from Remote Export Options dialog","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
			End If
			
			If dictRemoteExportInfo("IncludeReference")<>"" Then
				aTemp=Split(dictRemoteExportInfo("IncludeReference"),"^")
				For iCounter = 0 To Ubound(aTemp)
					If sOptionSettingsHeader="" Then
						sOptionSettingsHeader="Include Reference"
						sOptionSettingsValue=aTemp(iCounter)
					Else
						sOptionSettingsHeader=sOptionSettingsHeader & "^Include Reference"					
						sOptionSettingsValue=sOptionSettingsValue & "^"	& aTemp(iCounter)
					End If
				Next
			End If
			
			'Click on OK button
			If Fn_UI_JavaButton_Operations("RAC_Common_RemoteExportOperations","Click",objRemoteExportOptions,"jbtn_OK")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to click on [ OK ] button from Remote Export Options dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)			
		End If
		
		'Click on Yes button
		If Fn_UI_JavaButton_Operations("RAC_Common_RemoteExportOperations","Click",objRemoteExport,"jbtn_Yes")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to click on [ Yes ] button from Remote Export dialog","","","","","")
			Call Fn_ExitTest()
		End If
'		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		If objOptionsSettings.Exist(30) Then
			If sAction="ExportAndVerifyOptionSettings" Then
				If sOptionSettingsHeader<>"" Then
					sOptionSettingsHeader=Split(sOptionSettingsHeader,"^")
					sOptionSettingsValue=Split(sOptionSettingsValue,"^")
					For iCounter = 0 To Ubound(sOptionSettingsHeader)
						objOptionsSettings.JavaObject("jobj_OptionSettingValues").SetTOProperty "attached text",sOptionSettingsHeader(iCounter)
						If objOptionsSettings.JavaObject("jobj_OptionSettingValues").Exist(1) Then
							If Instr(1,Lcase(objOptionsSettings.JavaObject("jobj_OptionSettingValues").GetROProperty("text")),Lcase(sOptionSettingsValue(iCounter))) Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Sucessfully verified Option setting [ " & Cstr(sOptionSettingsHeader(iCounter)) & " ] show correct option setting value [ " & Cstr(sOptionSettingsValue(iCounter)) & " ] on [ Remote Export Options Settings ] dialog","","","","DONOTSYNC","")
							Else
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as Option setting [ " & Cstr(sOptionSettingsHeader(iCounter)) & " ] does not show correct option setting value [ " & Cstr(sOptionSettingsValue(iCounter)) & " ] on [ Remote Export Options Settings ] dialog","","","","","")
								Call Fn_ExitTest()
							End If	
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as Options Setting [ " & Cstr(sOptionSettingsHeader(iCounter)) & " ] does not available on Options Settings dialog","","","","","")
							Call Fn_ExitTest()
						End If
					Next
				End If	
			End If
			'Click on Yes button
			If Fn_UI_JavaButton_Operations("RAC_Common_RemoteExportOperations","Click",objOptionsSettings,"jbtn_Yes")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects as fail to click on [ Yes ] button from Options Settings dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
		
		If objRemoteExport.Exist(6) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects","","","","","")
			Call Fn_ExitTest()
		End If

		'Capturing execution end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_RemoteExportOperations",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to remote export selected objects due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully remote exported selected objects","","","","","")
		End If		
End Select

'Creating object of [ Remote Export ] dialog
Set objRemoteExport=Nothing
Set objOptionsSettings=Nothing
Set objRemoteExportOptions=Nothing
Set objRemoteSiteSelection=Nothing

Function Fn_ExitTest()
	'Creating object of [ Remote Export ] dialog
	Set objRemoteExport=Nothing
	Set objOptionsSettings=Nothing
	Set objRemoteExportOptions=Nothing
	Set objRemoteSiteSelection=Nothing
	ExitTest
End Function


