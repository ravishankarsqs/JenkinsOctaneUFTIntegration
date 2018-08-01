'! @Name 			RAC_Common_EditProperties
'! @Details 		This action word is used to edit objects properties
'! @InputParam1 	sAction 		: Action to be performed
'! @InputParam2 	sLink 			: Name of the link on properties dialog which is to be clicked
'! @InputParam3 	sInvokeOption 	: Invoke Option (menu or nooption)
'! @InputParam4 	sButtonName 	: button name to be clicked after verifying the properties
'! @InputParam5 	dictEditProperties 	: external dictionary parameter to pass properties details
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			24 Jun 2016.
'! @Version 		1.0
'! @Example 		dictEditProperties("PropertyName") = "Real Description"
'! @Example 		dictEditProperties("PropertyValue") = "Updated Description" 
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_EditProperties","RAC_Common_EditProperties",oneIteration,"ListBox","All","menu","Cancel"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sLink,sInvokeOption,sButtonName
Dim objEditProperties,objCheckOut,objCheckIn
Dim objDescription,objStaticText,objChildObjectCollection
Dim aProperty,aValues,sTempPropertyName
Dim sPerspective
Dim iCounter,iCount
Dim bFlag,sTempValue

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sLink =	Parameter("sLink")
sInvokeOption = Parameter("sInvokeOption")
sButtonName = Parameter("sButtonName")

If sPerspective="" Then
	sPerspective=Fn_RAC_GetActivePerspectiveName("")
End If

'Creating Object of [ Edit Properties ] dialog
Select Case LCase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "","myteamcenter"
		Set objEditProperties=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_EditProperties","")
		Set objCheckOut=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_CheckingOut","")
		Set objCheckIn=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_CheckingIn","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "structuremanager"
		Set objEditProperties=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_EditProperties@2","")
		Set objCheckOut=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_CheckingOut","")
		Set objCheckIn=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_CheckingIn","")
End Select

'Setting property dialog title
If dictEditProperties("PropertyDialogTitle")<>"" Then
	Call Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties,"","title",dictEditProperties("PropertyDialogTitle"))
End If

'Invoking properties dialog
Select Case Lcase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this option when user want to open dialog out of this function
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditProperties"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "keypress"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"KeyPress","EditProperties"
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_EditProperties"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

If sAction<>"VerifyErrorDialogAndCheckIn" And sAction<>"ClickButton" Then
	'Checking existance of Check Out dialog
	If Fn_UI_Object_Operations("RAC_Common_EditProperties","Exist",objCheckOut,GBL_MIN_TIMEOUT,"","") Then
		LoadAndRunAction "RAC_Common\RAC_Common_ObjectCheckOut","RAC_Common_ObjectCheckOut",OneIteration,"CheckOut","nooption",sPerspective,"","","","",""
	End If
	
GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_EditProperties"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction	

	'Checking existance of [ Properties ] dialog
	If Fn_UI_Object_Operations("RAC_Common_EditProperties","Exist",objEditProperties,GBL_MIN_TIMEOUT,"","")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit object properties as [ Edit Properties ] dialog does not exist","","","","","")
		Call Fn_ExitTest()
	End If

	'Capture business functionality start time
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_EditProperties",sAction,"","")

	'Clicking on page link
	If sLink="<<SKIP>>" Then
		'Do nothing
	ElseIf sLink="" Then
		If Fn_UI_Object_Operations("RAC_Common_EditProperties", "settoproperty", objEditProperties.JavaStaticText("jstx_BottomLink"),"","label","General")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ General ] link of properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist", objEditProperties.JavaStaticText("jstx_BottomLink"),GBL_ZERO_TIMEOUT,"","") Then
			If Fn_UI_JavaStaticText_Operations("RAC_Common_EditProperties","Click",objEditProperties,"jstx_BottomLink",1,1,"LEFT")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ General ] link of properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ General ] link of properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		'to click on [ Show empty properties... ] link
		If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_BottomLink"),"","label","Show empty properties...")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ Show empty properties... ] link of properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist", objEditProperties.JavaStaticText("jstx_BottomLink"),"2","","") Then
			If Fn_UI_JavaStaticText_Operations("RAC_Common_EditProperties","Click",objEditProperties,"jstx_BottomLink",1,1,"LEFT")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ Show empty properties... ] link of properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If			
	Else	
		If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_BottomLink"),"","label",sLink)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ " & Cstr(sLink) & " ] link of properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		
		If Fn_UI_Object_Operations("RAC_Common_EditProperties","Exist",objEditProperties.JavaStaticText("jstx_BottomLink"),GBL_ZERO_TIMEOUT,"","") Then
			If Fn_UI_JavaStaticText_Operations("RAC_Common_EditProperties","Click",objEditProperties,"jstx_BottomLink",1,1,"LEFT")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ " & Cstr(sLink) & " ] link of properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ " & Cstr(sLink) & " ] link of properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sLink="All" Then
			'to click on [ Show empty properties... ] link
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_BottomLink"),"","label","Show empty properties...")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ Show empty properties... ] link of properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist", objEditProperties.JavaStaticText("jstx_BottomLink"),GBL_ZERO_TIMEOUT,"","") Then
				If Fn_UI_JavaStaticText_Operations("RAC_Common_EditProperties","Click",objEditProperties,"jstx_BottomLink",1,1,"LEFT")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ Show empty properties... ] link of properties dialog","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			End If
		End If
	End If
End If

Select Case sAction		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'case to select value static text drop down
	Case "DropDown_StaticText"
		If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",dictEditProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as property does not exist\available on edit properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		
		If Fn_UI_JavaButton_Operations("RAC_Common_EditProperties","Click", objEditProperties,"jbtn_StaticTextDropDown") =false Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as fail to click on dropdown button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	
		bFlag = False
		Set objDescription=Description.Create()
		objDescription("Class Name").value = "JavaStaticText"		
		Set objStaticText = Fn_UI_Object_GetChildObjects("RAC_Common_EditProperties", objEditProperties, "Class Name", "JavaStaticText")
		For iCounter = 0 to objStaticText.count-1
			If  Fn_UI_Object_Operations("RAC_Common_EditProperties","getroproperty",objStaticText(iCounter),"","label","")= dictEditProperties("PropertyValue") Then
				objStaticText(iCounter).Click 1,1
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				bFlag=True
				Exit for
			End If
		Next
		Set objStaticText =Nothing
		Set objDescription=Nothing
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] value","","","","","")'pass log
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited property [ " & Cstr(dictEditProperties("PropertyName")) & " ] value to [ " & Cstr(dictEditProperties("PropertyValue")) & " ]","","","","","")
		End If				 
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to modify property value from edit boxes
	Case "EditBox"
		aProperty=Split(dictEditProperties("PropertyName"),"~")
		aValues=Split(dictEditProperties("PropertyValue"),"~")
		
		For iCounter=0 to ubound(aProperty)
		
		    sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
		
			If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist",objEditProperties.JavaEdit("jedt_PropertyEdit"),"","","") Then
				If Fn_UI_JavaEdit_Operations("RAC_Common_EditProperties", "Set", objEditProperties, "jedt_PropertyEdit",aValues(iCounter) )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as fail to set value in edit box","","","","","")
					Call Fn_ExitTest()
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited property [ " & Cstr(aProperty(iCounter)) & " ] value to [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next	
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to modify property value of Radio Button
	Case "RadioButton"
		aProperty=Split(dictEditProperties("PropertyName"),"~")
		aValues=Split(dictEditProperties("PropertyValue"),"~")
		
		For iCounter=0 to ubound(aProperty)
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF

			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaRadioButton("jrdb_PropertyRadioButton"),"","attached text",aValues(iCounter) & ":")=False Then	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
						
			If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist", objEditProperties.JavaRadioButton("jrdb_PropertyRadioButton"),"","","") Then
				If Fn_UI_JavaRadioButton_Operations("RAC_Common_EditProperties", "Set", objEditProperties, "jrdb_PropertyRadioButton", aValues(iCounter)) =False then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as fail to select radio button option [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited property [ " & Cstr(aProperty(iCounter)) & " ] value to [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
				End IF
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()			
			End If
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
	Case "ListOfValues"
		aProperty=Split(dictEditProperties("PropertyName"),"~")
		aValues=Split(dictEditProperties("PropertyValue"),"~")
		
		For iCounter=0 to ubound(aProperty)			
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_JavaCheckBox_Operations("RAC_Common_EditProperties", "Set", objEditProperties.JavaCheckBox("jckb_ListOfValues"), "", "ON") =False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as fail to check the option","","","","","")
				Call Fn_ExitTest()
			End if
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			'Verify existence of list of values dropdown
			If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist", objEditProperties.JavaList("jlst_ListOfValues"),"","","") Then			
				'Verify existence of value
				If Fn_UI_JavaList_Operations("RAC_Common_EditProperties","Exist",objEditProperties,"jlst_ListOfValues",aValues(iCounter), "", "")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as value [ " & Cstr(aValues(iCounter)) & " ] does not exist in listbox","","","","","")
					Call Fn_ExitTest()
				End If
				
				'Select the value
				If Fn_UI_JavaList_Operations("RAC_Common_EditProperties","Activate",objEditProperties,"jlst_ListOfValues",aValues(iCounter), "", "") = False Then					
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as fail to select value [ " & Cstr(aValues(iCounter)) & " ] from listbox","","","","","")
					Call Fn_ExitTest()
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited property [ " & Cstr(aProperty(iCounter)) & " ] value to [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "PasteLink"
		If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",dictEditProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as property does not exist\available on edit properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		objEditProperties.JavaObject("jobj_LinkOptionDropDown").Click 1,1
		Select Case sAction
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "PasteLink"
					objEditProperties.JavaMenu("index:=0","label:=Paste").Select
		End Select
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as fail to select [ Paste ] option","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited property [ " & Cstr(aProperty(iCounter)) & " ] by performing [ Paste link ] option","","","","","")
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify edit boxes does not allow alphabet value
	Case "EditBox_VerifyAlphabetValueNotAllowed"		
		aProperty=Split(dictEditProperties("PropertyName"),"~")
		
		For iCounter=0 to UBound(aProperty)
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist", objEditProperties.JavaEdit("jedt_PropertyEdit"),"","","") Then
				Call Fn_UI_JavaEdit_Operations("RAC_Common_EditProperties", "Set", objEditProperties, "jedt_PropertyEdit", "ABcdEfG")
				If Fn_UI_JavaEdit_Operations("RAC_Common_EditProperties", "GetText", objEditProperties, "jedt_PropertyEdit", "")="ABcdEfG" Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property allow aphabet value","","","","","")
					Call Fn_ExitTest()
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property does not allow aphabet value","","","","DONOTSYNC","")
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictEditProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify edit boxes does not allow alphanumeric value
	Case "EditBox_VerifyAlphaNumericValueNotAllowed"		
		aProperty=Split(dictEditProperties("PropertyName"),"~")
		
		For iCounter=0 to UBound(aProperty)
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist", objEditProperties.JavaEdit("jedt_PropertyEdit"),"","","") Then
				Call Fn_UI_JavaEdit_Operations("RAC_Common_EditProperties", "Set", objEditProperties, "jedt_PropertyEdit", "123ABcdEfG45")
				If Fn_UI_JavaEdit_Operations("RAC_Common_EditProperties", "GetText", objEditProperties, "jedt_PropertyEdit", "")="123ABcdEfG45" Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property allow alpha numeric value","","","","","")
					Call Fn_ExitTest()
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property does not allow alpha numeric value","","","","DONOTSYNC","")
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictEditProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
	Case "LOVTreeTable"
		aProperty=Split(dictEditProperties("PropertyName"),"~")
		aValues=Split(dictEditProperties("PropertyValue"),"~")
		
		For iCounter=0 to ubound(aProperty)
		
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			If iCounter<>0 Then
				wait 1
			End If
			If Fn_UI_JavaButton_Operations("RAC_Common_EditProperties","Click", objEditProperties,"jbtn_StaticTextDropDown") =false Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as fail to click on dropdown button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			wait 2
			
			Set objDescription = Description.Create()
			objDescription("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
			objDescription("path").Value = "LOVTreeTable;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JWindow;EditPropertiesDialog;WEmbeddedFrame;Composite;Shell;"
			objDescription("displayed").Value = "1"
			Set objChildObjectCollection = objEditProperties.ChildObjects(objDescription)
			If objChildObjectCollection.Count = 0 Then
				wait 10
				Set objChildObjectCollection =Nothing
				Set objDescription =Nothing
				Set objDescription = Description.Create()
				objDescription("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
				objDescription("path").Value = "LOVTreeTable;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JWindow;EditPropertiesDialog;WEmbeddedFrame;Composite;Shell;"
				objDescription("displayed").Value = "1"
				Set objChildObjectCollection = objEditProperties.ChildObjects(objDescription)
			End If
			
			If objChildObjectCollection.Count > 0 Then
				bFlag=False
				For iCount=0 to objChildObjectCollection(0).GetROProperty("rows")-1				
					If Trim(aValues(iCounter))=trim(objChildObjectCollection(0).Object.getValueAt(iCount,0).getDisplayableValue()) Then
						objChildObjectCollection(0).DoubleClickCell iCount,0
						Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
						bFlag=True
						Exit for
					End If
				Next	
			End If
			Set objChildObjectCollection = Nothing
			Set objDescription = Nothing
			
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(aProperty(iCounter)) & " ] as fail to select value [ " & Cstr(aValues(iCounter)) & " ] from listbox","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited property [ " & Cstr(aProperty(iCounter)) & " ] value to [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
			End If			
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
	'Verify LOV table exist
	Case "LOVTreeTableEXT"
		aValues=Split(dictEditProperties("PropertyValue"),"~")
		sTempValue=aValues(ubound(aValues))
		
		If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",dictEditProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as property does not exist\available on edit properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		
		If Fn_UI_JavaButton_Operations("RAC_Common_EditProperties","Click", objEditProperties,"jbtn_StaticTextDropDown") =false Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as fail to click on dropdown button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Wait 2
		
		Set objDescription = Description.Create()
		objDescription("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
		objDescription("path").Value = "LOVTreeTable;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JWindow;EditPropertiesDialog;WEmbeddedFrame;Composite;Shell;"
		objDescription("displayed").Value = "1"
		Set objChildObjectCollection = objEditProperties.ChildObjects(objDescription)
		If objChildObjectCollection.Count = 0 Then
			Wait 10
			Set objChildObjectCollection =Nothing
			Set objDescription =Nothing
			Set objDescription = Description.Create()
			objDescription("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
			objDescription("path").Value = "LOVTreeTable;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JWindow;EditPropertiesDialog;WEmbeddedFrame;Composite;Shell;"
			objDescription("displayed").Value = "1"
			Set objChildObjectCollection = objEditProperties.ChildObjects(objDescription)
		End If
		
		If objChildObjectCollection.Count > 0 Then
			bFlag=False
			For iCount=0 to objChildObjectCollection(0).GetROProperty("rows")-1				
				If Trim(sTempValue)=trim(objChildObjectCollection(0).Object.getValueAt(iCount,0).getDisplayableValue()) Then
					objChildObjectCollection(0).DoubleClickCell iCount,0
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
					bFlag=True
					Exit for
				End If
			Next	
		End If
		Set objChildObjectCollection = Nothing
		Set objDescription = Nothing
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as fail to select value [ " & Cstr(sTempValue) & " ] from listbox","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited property [ " & Cstr(dictEditProperties("PropertyName")) & " ] value to [ " & Cstr(sTempValue) & " ]","","","","DONOTSYNC","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify Error dialog and check in
	Case "VerifyErrorDialogAndCheckIn"
		dictErrorMessageInfo("Action")="Basic"
		dictErrorMessageInfo("Perspective")=LCase(sPerspective)
		dictErrorMessageInfo("DialogTitle")="Properties..."
		If dictEditProperties("PropertyName")<>"" Then			
			dictEditProperties("ErrorMessage")=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ErrorMessage_ERM",dictEditProperties("ErrorMessage"),"")
			dictEditProperties("ErrorMessage")= dictEditProperties("ErrorMessage") & " " & dictEditProperties("PropertyName")
			dictErrorMessageInfo("ErrorMessageXMLName")=""
		Else
			dictErrorMessageInfo("ErrorMessageXMLName")="RAC_ErrorMessage_ERM"
		End IF	
		dictErrorMessageInfo("ErrorMessage")=dictEditProperties("ErrorMessage")		
		dictErrorMessageInfo("Button")="OK"
		LoadAndRunAction "RAC_ErrorMessage\RAC_ErrM_ErrorMessageOperations","RAC_ErrM_ErrorMessageOperations",OneIteration,"ComonErrorDialog"
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_EditProperties"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
		LoadAndRunAction "RAC_Common\RAC_Common_ObjectCheckIn","RAC_Common_ObjectCheckIn",OneIteration,"CheckIn", "nooption", sPerspective
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_EditProperties"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		'Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to check existance of Property lable
	Case "VerifyPropertyLabels"
		aProperty=Split(dictEditProperties("PropertyName"),"~")		
		For iCounter=0 to ubound(aProperty)
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist", objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","","") Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property available on edit properties dialog","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "VerifyPropertyLabelsNonExist"
		aProperty=Split(dictEditProperties("PropertyName"),"~")		
		For iCounter=0 to ubound(aProperty)
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on edit properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_EditProperties", "Exist", objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property Not available on edit properties dialog","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property exists on edit properties dialog","","","","DONOTSYNC","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify specific Edit box is editable
	Case "EditBox_VerifyEditableState"
		aProperty=Split(dictEditProperties("PropertyName"),"~")	
		For iCounter=0 to ubound(aProperty)
		      If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify [ " & Cstr(aProperty(iCounter)) & " ] property is  editable as [ " & Cstr(aProperty(iCounter)) & " ] property field does not exist\available on properties dialog","","","","","")
					Call Fn_ExitTest()
			  Else
					If Fn_UI_Object_Operations("RAC_Common_EditProperties","getroproperty",objEditProperties.JavaEdit("jedt_PropertyEdit"),"","editable","")<>0 Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property field is in editable state","","","","DONOTSYNC","")
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property field is in non editable state","","","","","")
						Call Fn_ExitTest()
					End If
			  End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify specific Edit box is not editable
	Case "EditBox_VerifyNonEditableState"
		aProperty=Split(dictEditProperties("PropertyName"),"~")	
		For iCounter=0 to ubound(aProperty)
			If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify [ " & Cstr(aProperty(iCounter)) & " ] property is non editable as [ " & Cstr(aProperty(iCounter)) & " ] property field does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()	
			Else
				If Fn_UI_Object_Operations("RAC_Common_EditProperties","getroproperty",objEditProperties.JavaEdit("jedt_PropertyEdit"),"","editable","")=0 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property field is in non editable state","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property field is in  editable state","","","","","")
					Call Fn_ExitTest()
				End If
		   End If
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
	'Verify LOV table values
	Case "VerifyLOVTreeTableValues"
		aValues=Split(dictEditProperties("PropertyValue"),"~")
		
		If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",dictEditProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictEditProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		
		If Fn_UI_JavaButton_Operations("RAC_Common_EditProperties","Click", objEditProperties,"jbtn_StaticTextDropDown") =false Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as fail to click on dropdown button of property [ " & Cstr(dictEditProperties("PropertyName")) & " ] on properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		wait 2
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		For iCounter=0 to ubound(aValues)		
			Set objDescription = Description.Create()
			objDescription("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
			objDescription("path").Value = "LOVTreeTable;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JWindow;EditPropertiesDialog;WEmbeddedFrame;Composite;Shell;"
			objDescription("displayed").Value = "1"
			Set objChildObjectCollection = objEditProperties.ChildObjects(objDescription)
			'Added additional code
			If objChildObjectCollection.Count = 0 Then
				wait 10
				Set objChildObjectCollection = Nothing
				Set objDescription = Nothing
				Set objDescription = Description.Create()
				objDescription("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
				objDescription("path").Value = "LOVTreeTable;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JWindow;EditPropertiesDialog;WEmbeddedFrame;Composite;Shell;"
				objDescription("displayed").Value = "1"
				Set objChildObjectCollection = objEditProperties.ChildObjects(objDescription)
			End If
			
			If objChildObjectCollection.Count > 0 Then
				bFlag=False
				For iCount=0 to objChildObjectCollection(0).GetROProperty("rows")-1				
					If Trim(aValues(iCounter))=trim(objChildObjectCollection(0).Object.getValueAt(iCount,0).getDisplayableValue()) Then
						bFlag=True
						Exit for
					End If
				Next	
			End If
			Set objChildObjectCollection = Nothing
			Set objDescription = Nothing
			
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictEditProperties("PropertyName")) & " ] property does not contain value [ " & Cstr(aValues(iCounter)) & " ] on properties dialog","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictEditProperties("PropertyName")) & " ] property contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
			End If			
		Next
		If Fn_UI_JavaButton_Operations("RAC_Common_EditProperties","Click", objEditProperties,"jbtn_StaticTextDropDown") =false Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as fail to click on dropdown button of property [ " & Cstr(dictEditProperties("PropertyName")) & " ] on properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
	'Verify LOV table exist
	Case "DropDownListLOVTreeTable"
		aValues=Split(dictEditProperties("PropertyValue"),"~")
		
		sTempPropertyName=dictEditProperties("PropertyName")
		dictEditProperties("PropertyName")=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",dictEditProperties("PropertyName"),"")
		If Cstr(dictEditProperties("PropertyName"))="False" Then
			dictEditProperties("PropertyName")=sTempPropertyName
		End If
			
		If Fn_UI_Object_Operations("RAC_Common_EditProperties","settoproperty",objEditProperties.JavaStaticText("jstx_PropertyLabel"),"","label",dictEditProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as property does not exist\available on edit properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		
		If Fn_UI_JavaCheckBox_Operations("RAC_Common_EditProperties", "Set", objEditProperties.JavaCheckBox("jckb_Edit16"), "", "ON") =False then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as fail to check the option","","","","","")
			Call Fn_ExitTest()
		End if
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
										
		For iCounter=0 to ubound(aValues)
			bFlag=False
			If iCounter<>0 Then
				wait 1
			End If
			
			If Fn_UI_JavaButton_Operations("RAC_Common_EditProperties","Click", objEditProperties,"jbtn_ListLOVDropDown") =false Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as fail to click on dropdown button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			wait 2
			
			Set objDescription = Description.Create()
			objDescription("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
			objDescription("path").Value = "LOVTreeTable;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JWindow;EditPropertiesDialog;WEmbeddedFrame;Composite;Shell;"
			objDescription("displayed").Value = "1"
			Set objChildObjectCollection = objEditProperties.ChildObjects(objDescription)
			If objChildObjectCollection.Count = 0 Then
				wait 10
				Set objChildObjectCollection =Nothing
				Set objDescription =Nothing
				Set objDescription = Description.Create()
				objDescription("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
				objDescription("path").Value = "LOVTreeTable;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JWindow;EditPropertiesDialog;WEmbeddedFrame;Composite;Shell;"
				objDescription("displayed").Value = "1"
				Set objChildObjectCollection = objEditProperties.ChildObjects(objDescription)
			End If
			
			If objChildObjectCollection.Count > 0 Then
				For iCount=0 to objChildObjectCollection(0).GetROProperty("rows")-1				
					If Trim(aValues(iCounter))=trim(objChildObjectCollection(0).Object.getValueAt(iCount,0).getDisplayableValue()) Then
						objChildObjectCollection(0).DoubleClickCell iCount,0
						Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
						bFlag=True
						Exit for
					End If
				Next	
			End If
			Set objChildObjectCollection = Nothing
			Set objDescription = Nothing
			If bFlag=False Then
				Exit For
			End If
			If Fn_UI_JavaButton_Operations("RAC_Common_EditProperties","Click", objEditProperties,"jbtn_Add16") =false Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as fail to click on Add button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Next	
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as fail to select value [ " & Cstr(aValues(iCounter)) & " ] from listbox","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully edited property [ " & Cstr(dictEditProperties("PropertyName")) & " ] value to [ " & Cstr(dictEditProperties("PropertyValue")) & " ]","","","","","")
		End If
		
		If Fn_UI_JavaCheckBox_Operations("RAC_Common_EditProperties", "Set", objEditProperties.JavaCheckBox("jckb_Edit16"), "", "OFF") =False then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to edit property [ " & Cstr(dictEditProperties("PropertyName")) & " ] as fail to check the option","","","","","")
			Call Fn_ExitTest()
		End if
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to click on specific button		
	Case "ClickButton"	
End Select

'Clicking on button
If sButtonName<>"" Then
	If sButtonName="SaveAndCheckIn_WithoutCheckIn" or sButtonName="jbtn_SaveAndCheckIn_WithoutCheckIn" Then
		sButtonName="SaveAndCheckIn"
		IF Fn_UI_JavaButton_Operations("RAC_Common_EditProperties", "Click", objEditProperties,"jbtn_" & sButtonName)=False then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ Edit properties ] operation as fail to click on [ " & Cstr(sButtonName) & " ] button of properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		'Capture business functionality end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_EditProperties",sAction,"","")
	Else	
		IF Fn_UI_JavaButton_Operations("RAC_Common_EditProperties", "Click", objEditProperties,"jbtn_" & sButtonName)=False then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ Edit properties ] operation as fail to click on [ " & Cstr(sButtonName) & " ] button of properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		'Capture business functionality end time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_EditProperties",sAction,"","")
		
		If sButtonName="SaveAndCheckIn" or sButtonName="jbtn_SaveAndCheckIn" Then
			LoadAndRunAction "RAC_Common\RAC_Common_ObjectCheckIn","RAC_Common_ObjectCheckIn",OneIteration,"CheckIn", "nooption", sPerspective
		End If
	End If
Else
	'Capture business functionality end time
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_EditProperties",sAction,"","")
End If

'Releasing all objects
Set objEditProperties=Nothing
Set objCheckOut=Nothing
Set objCheckIn=Nothing

Function Fn_ExitTest()
	'Releasing all objects
	Set objEditProperties=Nothing
	Set objCheckOut=Nothing
	Set objCheckIn=Nothing
	ExitTest
End Function

