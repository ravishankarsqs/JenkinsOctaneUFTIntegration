'! @Name 			RAC_Common_MassUpdateOperations
'! @Details 		Action word to perform operations on Mass Update dialog
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam2		sInvokeOption		: Mass Update dialog invoke option
'! @InputParam3 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			18 Jul 2017
'! @Version 		1.0
'! @Example 		dictMassUpdateInfo("Operation")="Replace Part"
'! @Example 		dictMassUpdateInfo("ReplacementFlag")=True
'! @Example 		dictMassUpdateInfo("ReplacementID")="123456"
'! @Example 		dictMassUpdateInfo("ReplacementName")="Asm Cockpit"
'! @Example 		dictMassUpdateInfo("ReplacementRevision")="AA"
'! @Example 		dictMassUpdateInfo("ImpactedPartsToUpdate")="23456/AA-Name"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_MassUpdateOperations","RAC_Common_MassUpdateOperations",OneIteration,"execute","menu",""

Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sInvokeOption,sButton
Dim sPerspective,sObjectName
Dim objMassUpdate
Dim iCount
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sButton = Parameter("sButton")

'Invoking [ Mass Update ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditMassUpdate"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating object of [ Mass Update ] dialog
Select Case lcase(sPerspective)
	Case "","myteamcenter"
		Set objMassUpdate = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_MassUpdate","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_MassUpdateOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of Mass Update dialog
If Fn_UI_Object_Operations("RAC_Common_MassUpdateOperations","Exist", objMassUpdate, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Mass Update ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capture execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_MassUpdateOperations",sAction,"","")
		
Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create execute Mass Update
	Case "execute"
		
		If dictMassUpdateInfo("Operation")<>"" Then
			'Select value for Mass Update Operation 
			If Fn_UI_JavaList_Operations("RAC_Common_MassUpdateOperations", "Select", objMassUpdate, "jlst_Operation", dictMassUpdateInfo("Operation"), "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ MassUpdate ] dialog as failed to set value of MassUpdate as [ " & dictMassUpdateInfo("Operation") & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		
		'Selecting replacement 
		If Cbool(dictMassUpdateInfo("ReplacementFlag"))=True Then
			Call Fn_UI_Object_Operations("RAC_Common_MassUpdateOperations","settoproperty",objMassUpdate.JavaStaticText("jstx_MassUpdateLabel"),"","label","Replacement")
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateMassUpdate", "Click", objMassUpdate, "jbtn_OpenSearchDialog")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Open Search Dialog ] of replacement section on [ MassUpdate ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Checking existance of Search dialog
			If Fn_UI_Object_Operations("RAC_Common_MassUpdateOperations","Exist", objMassUpdate.JavaWindow("jwnd_Search"), "","","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Mass Update Object Search ] dialog as dialog does not exist","","","","","")
				Call Fn_ExitTest()
			End If
			'Setting ID
			If Fn_UI_JavaEdit_Operations("RAC_Common_MassUpdateOperations", "Set", objMassUpdate.JavaWindow("jwnd_Search"), "jedt_ID",dictMassUpdateInfo("ReplacementID") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set replacement object id in [ ID ] field of [ Mass Update Object Search ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			
			If dictMassUpdateInfo("ReplacementName")<>"" Then
				If Fn_UI_JavaEdit_Operations("RAC_Common_MassUpdateOperations", "Set", objMassUpdate.JavaWindow("jwnd_Search"), "jedt_Name",dictMassUpdateInfo("ReplacementName") )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set replacement object name in [ Name ] field of [ Mass Update Object Search ] dialog","","","","","")
					Call Fn_ExitTest()
				End If
			End IF
			
			If dictMassUpdateInfo("ReplacementRevision")<>"" Then
				If Fn_UI_JavaEdit_Operations("RAC_Common_MassUpdateOperations", "Set", objMassUpdate.JavaWindow("jwnd_Search"), "jedt_Revision",dictMassUpdateInfo("ReplacementRevision") )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set replacement object revision in [ Revision ] field of [ Mass Update Object Search ] dialog","","","","","")
					Call Fn_ExitTest()
				End If
			End IF
			
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateMassUpdate", "Click", objMassUpdate.JavaWindow("jwnd_Search"), "jbtn_Search")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Search ] of replacement section on [ Mass Update Object Search ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			bFlag = False
			For iCount = 0 to objMassUpdate.JavaWindow("jwnd_Search").JavaTable("jtbl_SearchResults").GetROProperty("rows") -1
				bFlag = False
				sObjectName=objMassUpdate.JavaWindow("jwnd_Search").JavaTable("jtbl_SearchResults").Object.getItem(iCount).getData().getPropertyValue("object_string").toString()
				If cStr(dictMassUpdateInfo("ReplacementID")) <> "" AND dictMassUpdateInfo("ReplacementName") <> "" AND dictMassUpdateInfo("ReplacementRevision") <> "" Then
					If Trim(sObjectName) = dictMassUpdateInfo("ReplacementID") & "/" & dictMassUpdateInfo("ReplacementRevision") & "-" & dictMassUpdateInfo("ReplacementName") Then
						bFlag = True
					End IF
				End If
				If bFlag = True Then
					objMassUpdate.JavaWindow("jwnd_Search").JavaTable("jtbl_SearchResults").DoubleClickCell iCount,"Object","LEFT"
					Call Fn_RAC_ReadyStatusSync(5)
					Exit for
				End If
			Next
			If bFlag = True Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully open object based on search criteria : ID = [ " & Cstr(dictMassUpdateInfo("ReplacementID")) & " ] from [ Mass Update Object Search ] dialog","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to find and open object from [ Mass Update Object Search ] dialog as there are no objects found based on search criteria : ID = [ " & Cstr(dictMassUpdateInfo("ReplacementID")) & " ]","","","","","")
				Call Fn_ExitTest()
			End IF		
		End If
		
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateMassUpdate", "Click", objMassUpdate, "jbtn_Next")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button on [ Mass Update ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		If dictMassUpdateInfo("ImpactedPartsToUpdate")<>"" Then
			bFlag = False
			For iCount = 0 to objMassUpdate.JavaTable("jtbl_TargetUsed").GetROProperty("rows") -1
				bFlag = False
				sObjectName=objMassUpdate.JavaTable("jtbl_TargetUsed").Object.getItem(iCount).getData().getPropertyValue("object_string").toString()
				If Trim(sObjectName) = dictMassUpdateInfo("ImpactedPartsToUpdate") Then
					bFlag = True
				End IF
				
				If bFlag = True Then
					objMassUpdate.JavaTable("jtbl_TargetUsed").SelectCell iCount,0
					Call Fn_RAC_ReadyStatusSync(1)
					Exit for
				End If
			Next
			If bFlag = True Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected Impacted Parts [ " & Cstr(dictMassUpdateInfo("ImpactedPartsToUpdate")) & " ] from [ Impacted Parts to update ] table","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Impacted Parts [ " & Cstr(dictMassUpdateInfo("ImpactedPartsToUpdate")) & " ] from [ Impacted Parts to update ] table","","","","","")
				Call Fn_ExitTest()
			End IF
		End If
		
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateMassUpdate", "Click", objMassUpdate, "jbtn_Next")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button on [ Mass Update ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateMassUpdate", "Click", objMassUpdate, "jbtn_Execute")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Execute ] button on [ Mass Update ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Checking existance of Results dialog
		If Fn_UI_Object_Operations("RAC_Common_MassUpdateOperations","Exist", JavaWindow("jwnd_DefaultWindow").JavaWindow("jwnd_Results"), GBL_DEFAULT_TIMEOUT,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Mass Update ] dialog as [ Results ] dialog does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		bFlag = False
		For iCount = 0 to JavaWindow("jwnd_DefaultWindow").JavaWindow("jwnd_Results").JavaTable("jtbl_Results").GetROProperty("rows") -1
			bFlag = False
			sObjectName=JavaWindow("jwnd_DefaultWindow").JavaWindow("jwnd_Results").JavaTable("jtbl_Results").Object.getItem(iCount).getData().getResultStatus()
			If Trim(Cstr(sObjectName)) = "0" Then
				bFlag = True
			Else
				Exit For
			End IF			
		Next
		
		If bFlag = True Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully execute mass properties","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to execute mass properties","","","","","")
			Call Fn_ExitTest()
		End IF
		
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateMassUpdate", "Click", JavaWindow("jwnd_DefaultWindow").JavaWindow("jwnd_Results"), "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Close ] button on [ Mass Update Results ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)			
End Select

'Releasing object of MassUpdate dialog
Set objMassUpdate =Nothing

Function Fn_ExitTest()
	'Releasing object of MassUpdate dialog
	Set objMassUpdate =Nothing
	ExitTest
End Function

