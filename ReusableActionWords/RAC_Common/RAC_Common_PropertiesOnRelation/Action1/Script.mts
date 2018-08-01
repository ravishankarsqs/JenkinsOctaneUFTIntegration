'! @Name 			RAC_Common_PropertiesOnRelation
'! @Details 		This action word is used to perform operations on Properties On Relation
'! @InputParam1 	sAction 		: Action to be performed
'! @InputParam2 	sInvokeOption 	: Invoke Option (menu or nooption)
'! @InputParam3 	sPropertyName 	: Property Name
'! @InputParam4 	sPropertyValue 	: Property Value
'! @InputParam5 	sButtonName 	: button name to be clicked after verifying the properties
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			07 Jul 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_PropertiesOnRelation","RAC_Common_PropertiesOnRelation",OneIteration,"ModifyProperty","summarytabrelatedpartrevisionpaste","Raw Material Quantity","10","Finish"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_PropertiesOnRelation","RAC_Common_PropertiesOnRelation",OneIteration,"verifymandatoryfields","Menu","Raw Material Quantity","","Cancel"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sPropertyName,sPropertyValue,sButtonName
Dim objPropertiesOnRelation
Dim aPropertyName,aPropertyValue
Dim iCounter,sPerspective

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sPropertyName =	Parameter("sPropertyName")
sPropertyValue =Parameter("sPropertyValue")
sButtonName = Parameter("sButtonName")

If sPerspective="" Then
	sPerspective=Fn_RAC_GetActivePerspectiveName("")
End If

'Creating Object of [ Edit Properties ] dialog
Select Case LCase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "","myteamcenter"
		Set objPropertiesOnRelation=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_PropertiesOnRelation","")
End Select


'Invoking properties dialog
Select Case Lcase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this option when user want to open dialog out of this function
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditPaste"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "summarytabrelatedpartrevisionpaste"	
		LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Summary","Related Part Revisions"
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", JavaWindow("jwnd_DefaultWindow").JavaStaticText("jstx_SummaryTabTableHeader"),"","label","Related Raw Material Revs")
		'Click on Add New button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", JavaWindow("jwnd_DefaultWindow"),"jbtn_Paste") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on Relation Properties dialog as fail to click on [ Paste ] button from [ Related Part Revisions ] tab of [ Summary ] tab","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_PropertiesOnRelation"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_PropertiesOnRelation",sAction,"","")


Select Case sAction		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'case to select value static text drop down
	Case "ModifyProperty"
		aPropertyName=Split(sPropertyName,"~")
		aPropertyValue=Split(sPropertyValue,"~")
		For iCounter = 0 to Ubound(aPropertyName)
			If Fn_UI_Object_Operations("RAC_Common_PropertiesOnRelation","settoproperty",objPropertiesOnRelation.JavaStaticText("jstx_PropertyLabel"),"","label",aPropertyName(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aPropertyName(iCounter)) & " ] as property does not exist\available on Properties on relation dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_PropertiesOnRelation", "Exist",objPropertiesOnRelation.JavaEdit("jedt_PropertyEdit"),"","","") Then
				If Fn_UI_JavaEdit_Operations("RAC_Common_PropertiesOnRelation", "Set", objPropertiesOnRelation, "jedt_PropertyEdit",aPropertyValue(iCounter) )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aPropertyName(iCounter)) & " ] as fail to set value in edit box on Properties on relation dialog","","","","","")
					Call Fn_ExitTest()
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited property [ " & Cstr(aPropertyName(iCounter)) & " ] value to [ " & Cstr(aPropertyValue(iCounter)) & " ] on Properties on relation dialog","","","","DONOTSYNC","")
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aPropertyName(iCounter)) & " ] as property does not exist\available on Properties on relation dialog","","","","","")
				Call Fn_ExitTest()
			End If			
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create validate specific fields are mandetory
	Case "verifymandatoryfields"		
		aPropertyName=Split(sPropertyName,"~")
		For iCounter=0 to Ubound(aPropertyName)	
			Call Fn_UI_Object_Operations("RAC_Common_PropertiesOnRelation","SetTOProperty", objPropertiesOnRelation.JavaStaticText("jstx_PropertyLabel"),"","label",aPropertyName(iCounter) & ":")
			If Fn_UI_Object_Operations("RAC_Common_PropertiesOnRelation","Exist", objPropertiesOnRelation.JavaStaticText("jstx_Asterix"), GBL_DEFAULT_TIMEOUT,"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(aPropertyName(iCounter)) & " ] is not mandatory field on Properties on relation dialog","","","","","")
				Call Fn_ExitTest()
			Else
			    Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : Successfully verified fields [ " & Cstr(aPropertyName(iCounter)) & " ] are mandatory field on Properties on relation dialog","","","","DONOTSYNC","")
			End If
		Next	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to click on specific button		
	Case "ClickButton"	
End Select

'Clicking on button
If sButtonName<>"" Then
	IF Fn_UI_JavaButton_Operations("RAC_Common_PropertiesOnRelation", "Click", objPropertiesOnRelation,"jbtn_" & sButtonName)=False then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & " ] operation as fail to click on [ " & Cstr(sButtonName) & " ] button of Properties on relation dialog","","","","","")
		Call Fn_ExitTest()
	End IF
	'Capture business functionality end time
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_PropertiesOnRelation",sAction,"","")
Else
	'Capture business functionality end time
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_PropertiesOnRelation",sAction,"","")
End If

'Releasing all objects
Set objPropertiesOnRelation=Nothing

Function Fn_ExitTest()
	'Releasing all objects
	Set objPropertiesOnRelation=Nothing
	ExitTest
End Function
