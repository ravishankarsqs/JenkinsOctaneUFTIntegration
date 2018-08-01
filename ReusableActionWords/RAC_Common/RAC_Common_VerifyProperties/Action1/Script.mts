'! @Name 			RAC_Common_VerifyProperties
'! @Details 		This action word is used to verify objects properties
'! @InputParam1 	sAction 		: Action to be performed
'! @InputParam2 	sLink 			: Name of the link on properties dialog which is to be clicked
'! @InputParam3 	sInvokeOption 	: Invoke Option (menu or nooption)
'! @InputParam4 	sButtonName 	: button name to be clicked after verifying the properties
'! @InputParam5 	dictProperties 	: external dictionary parameter to pass properties details
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			24 Jun 2016
'! @Version 		1.0
'! @Example 		dictProperties("PropertyName") = "Released Status"
'! @Example 		dictProperties("PropertyValue") = "Released" 
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_VerifyProperties","RAC_Common_VerifyProperties",oneIteration,"ListBox","All","menu","Cancel"
'! @Example 		
'! @Example 		dictProperties("PropertyName") = "Has Migration Forms"
'! @Example 		dictProperties("PropertyValue") = 0
'! @Example 		dictProperties("AdditionalProperties_FieldType") = "EditBox"
'! @Example 		dictProperties("AdditionalProperties_FieldNames") = Fn_FSOUtil_XMLFileOperations("getallnodevalues","RAC_DataMigration_Migration_Properties_APL","","")
'! @Example 		dictProperties("AdditionalProperties_CloseDialog") = True
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_VerifyProperties","RAC_Common_VerifyProperties",oneIteration,"OpenListBoxItemAndGetProperties","All","menu","Cancel"

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sLink,sInvokeOption,sButtonName
Dim iCounter,iItemsCount
Dim aValues,aProperty
Dim sPerspective, sDialogTitle,sTempPropertyName
Dim objPropeties, objReadOnly, objAdditionalPropertiesDialog
Dim sDate, sTempValue

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

'Creating Object of [ Checking Out ] dialog
Select Case LCase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "","myteamcenter"
		Set objPropeties=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_Properties","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "structuremanager"
		Set objPropeties=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_Properties@2","")
End Select

If Cbool(dictProperties("PopupPropertyDialog"))=True Then
	objPropeties.SetTOProperty "index",1
End If

'Setting property dialog title
If dictProperties("PropertyDialogTitle")<>"" Then
	Call Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "settoproperty", objPropeties,"","title",dictProperties("PropertyDialogTitle"))
End If


'Invoking properties dialog
Select Case Lcase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this option when user want to open dialog out of this function
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
	Case "menu" , ""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewProperties"
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "keypress"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"KeyPress","ViewProperties"
End Select		

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_VerifyProperties"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of [ Properties ] dialog
If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objPropeties,GBL_MIN_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify object properties as [ Properties ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Verify Properties",sAction,dictProperties("PropertyName"),dictProperties("PropertyValue"))
'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Clicking on page link
If sLink="<<SKIP>>" Then
	'Do Nothing
ElseIf sLink="" Then
	If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "settoproperty", objPropeties.JavaStaticText("jstx_BottomLink"),"","label","General")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ General ] link of properties dialog","","","","","")
		Call Fn_ExitTest()
	End IF
	If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaStaticText("jstx_BottomLink"),GBL_DEFAULT_MIN_TIMEOUT,"","") Then
		If Fn_UI_JavaStaticText_Operations("RAC_Common_VerifyProperties","Click",objPropeties,"jstx_BottomLink",1,1,"LEFT")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ General ] link of properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ General ] link of properties dialog","","","","","")
		Call Fn_ExitTest()
	End If
	'to click on [ Show empty properties... ] link
	If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_BottomLink"),"","label","Show empty properties...")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ Show empty properties... ] link of properties dialog","","","","","")
		Call Fn_ExitTest()
	End IF
	If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaStaticText("jstx_BottomLink"),GBL_DEFAULT_MIN_TIMEOUT,"","") Then
		If Fn_UI_JavaStaticText_Operations("RAC_Common_VerifyProperties","Click",objPropeties,"jstx_BottomLink",1,1,"LEFT")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ Show empty properties... ] link of properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	End If
Else	
	If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_BottomLink"),"","label",sLink)=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ " & Cstr(sLink) & " ] link of properties dialog","","","","","")
		Call Fn_ExitTest()
	End IF
	
	If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objPropeties.JavaStaticText("jstx_BottomLink"),GBL_DEFAULT_MIN_TIMEOUT,"","") Then
		If Fn_UI_JavaStaticText_Operations("RAC_Common_VerifyProperties","Click",objPropeties,"jstx_BottomLink",1,1,"LEFT")=False Then
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
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_BottomLink"),"","label","Show empty properties...")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ Show empty properties... ] link of properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaStaticText("jstx_BottomLink"),GBL_DEFAULT_MIN_TIMEOUT,"","") Then
			If Fn_UI_JavaStaticText_Operations("RAC_Common_VerifyProperties","Click",objPropeties,"jstx_BottomLink",1,1,"LEFT")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify properties as fail to click on [ Show empty properties... ] link of properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	End If
End If

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify values from list Box
	'to verify values of List box , list box should be enabled so if want to verify values from List box object should be check out
	Case "ListBox"
		sTempPropertyName=dictProperties("PropertyName")
		dictProperties("PropertyName")=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",dictProperties("PropertyName"),"")
		
		If Cstr(dictProperties("PropertyName"))="False" Then
			dictProperties("PropertyName")=sTempPropertyName
		End If
		
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",dictProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objPropeties.JavaList("jlst_PropertyField"),"","","") Then
			aValues=Split(dictProperties("PropertyValue"),"~")
			For iCounter=0 to uBound(aValues)
				If Fn_UI_JavaList_Operations("RAC_Common_VerifyProperties","Exist", objPropeties, "jlst_PropertyField", aValues(iCounter), "", "")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictProperties("PropertyName")) & " ] property contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
				End IF
			Next
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify property value from edit boxes
	Case "EditBoxInStr"
		aProperty=Split(dictProperties("PropertyName"),"~")
		aValues=Split(dictProperties("PropertyValue"),"~")
		
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
	
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
				If InStr(1,Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "GetText", objPropeties, "jedt_PropertyField", ""),aValues(iCounter)) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify property value from edit boxes
	Case "EditBox"
		aProperty=Split(dictProperties("PropertyName"),"~")
		aValues=Split(dictProperties("PropertyValue"),"~")
				
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If aProperty(iCounter)="CM Flag" Then
				'As part of 17.04 rediployement CM Flag field is removed from Properties dialog
			Else
				If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
					Call Fn_ExitTest()
				End IF
				
				If aValues(iCounter)="<<EMPTY>>" Then
					aValues(iCounter)=""
				End If
				
				If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
					If Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "GetText", objPropeties, "jedt_PropertyField", "")=aValues(iCounter) Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
						Call Fn_ExitTest()
					End If	
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
					Call Fn_ExitTest()
				End If
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify Radio Buttons
	Case "RadioButton"
		aProperty=Split(dictProperties("PropertyName"),"~")
		aValues=Split(dictProperties("PropertyValue"),"~")
		
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaRadioButton("jrdb_PropertyField"),"","attached text",aValues(iCounter))=False Then	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF

			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaRadioButton("jrdb_PropertyField"),"","","")  Then
				If Cstr(Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaRadioButton("jrdb_PropertyField"),"","value",""))="1" Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to get specific property state of specific Edit box : e.g { current value, editable state, enabled state }
	Case "EditBox_GetPropertyState"
		sTempPropertyName=dictProperties("PropertyName")
		dictProperties("PropertyName")=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",dictProperties("PropertyName"),"")
		
		If Cstr(dictProperties("PropertyName"))="False" Then
			dictProperties("PropertyName")=sTempPropertyName
		End If
		
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",dictProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1		
			DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_VerifyProperties"
			DataTable.Value("ReusableActionWordReturnValue","Global")= Cstr(Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaEdit("jedt_PropertyField"),"",dictProperties("PropertyState"),""))
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End if
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "ListBox_CheckProperty"
		sTempPropertyName=dictProperties("PropertyName")
		dictProperties("PropertyName")=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",dictProperties("PropertyName"),"")
		
		If Cstr(dictProperties("PropertyName"))="False" Then
			dictProperties("PropertyName")=sTempPropertyName
		End If
		
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",dictProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaList("jlst_PropertyField"),"","","") Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1		
			DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_VerifyProperties"
			DataTable.Value("ReusableActionWordReturnValue","Global") = objPropeties.JavaList("Property_field").CheckProperty(dictProperties("CheckPropertyName"),dictProperties("PropertyValue"))
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to get specific property state of specific Check Box : e.g { current value, attached text }
	Case "CheckBox_GetPropertyState"
		sTempPropertyName=dictProperties("PropertyName")
		dictProperties("PropertyName")=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",dictProperties("PropertyName"),"")
		
		If Cstr(dictProperties("PropertyName"))="False" Then
			dictProperties("PropertyName")=sTempPropertyName
		End If

		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",dictProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaCheckBox("jckb_PropertyField"), "") Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1		
			DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_VerifyProperties"
			DataTable.Value("ReusableActionWordReturnValue","Global") =Fn_UI_Object_GetROProperty("RAC_Common_VerifyProperties",objPropeties.JavaCheckBox("Property_field"),dictProperties("PropertyState"))
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End if
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify property links
	Case "Link"
		aProperty=Split(dictProperties("PropertyName"),"~")
		aValues=Split(dictProperties("PropertyValue"),"~")
		
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objPropeties.JavaStaticText("jstx_PropertyLabelValue"),"","","") Then
				If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaStaticText("jstx_PropertyLabelValue"),"","label","")=aValues(iCounter) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
	'Case to verify property links using lcase comparison
	Case "LinkExt"
		aProperty=Split(dictProperties("PropertyName"),"~")
		aValues=Split(dictProperties("PropertyValue"),"~")
		
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objPropeties.JavaStaticText("jstx_PropertyLabelValue"),"","","") Then
				If lcase(Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaStaticText("jstx_PropertyLabelValue"),"","label",""))=lcase(aValues(iCounter)) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "VerifyPropertyLabels"
		aProperty=Split(dictProperties("PropertyName"),"~")		
		For iCounter=0 to ubound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If

			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaStaticText("jstx_PropertyLabel"),"","","") Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property available on properties page","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties page","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "VerifyPropertyLabelsNotExist"
		aProperty=Split(dictProperties("PropertyName"),"~")		
		For iCounter=0 to ubound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If

			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaStaticText("jstx_PropertyLabel"),"","","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property not available on properties page","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property is exist\available on properties page","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "DateCheckbox_GetDate"	,"DateEditBox_GetDate"
		aProperty=Split(dictProperties("PropertyName"),"~")
		sDate=""
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_VerifyProperties"
		DataTable.Value("ReusableActionWordReturnValue","Global") ="False"
		
		For iCounter=0 to ubound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If

			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If sAction="DateCheckbox_GetDate" Then
	 		    If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaCheckBox("jckb_PropertyDateCheckbox"),"","","") Then
					If sDate="" Then
						sDate=Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaCheckBox("jckb_PropertyDateCheckbox"),"","label","")
					Else
						sDate=sDate & "~" & Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaCheckBox("jckb_PropertyDateCheckbox"),"","label","")
					End If
				End If	
			ElseIf sAction="DateEditBox_GetDate" Then
			      If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
					If sDate="" Then
						sDate=Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaEdit("jedt_PropertyField"),"","value","")
					Else
						sDate=sDate & "~" & Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaEdit("jedt_PropertyField"),"","value","")
					End If
				End If
			Else
				 Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				 Call Fn_ExitTest()
			End If
		Next
		
		If sDate<>"" Then
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
			Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
			DataTable.SetCurrentRow 1		
			DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_VerifyProperties"
			DataTable.Value("ReusableActionWordReturnValue","Global") =sDate
		End If
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify specific Edit box is not editable
	Case "EditBox_VerifyNonEditableState"
		sTempPropertyName=dictProperties("PropertyName")
		dictProperties("PropertyName")=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",dictProperties("PropertyName"),"")
		
		If Cstr(dictProperties("PropertyName"))="False" Then
			dictProperties("PropertyName")=sTempPropertyName
		End If

		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",dictProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify [ " & Cstr(dictProperties("PropertyName")) & " ] property is non editable as [ " & Cstr(dictProperties("PropertyName")) & " ] property field does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaEdit("jedt_PropertyField"),"","editable","")=0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictProperties("PropertyName")) & " ] property field is in non editable state","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property field is in editable state","","","","","")
				Call Fn_ExitTest()
			End If			
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify [ " & Cstr(dictProperties("PropertyName")) & " ] property is non editable as [ " & Cstr(dictProperties("PropertyName")) & " ] property field does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()					
		End if
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "VerifyCurrentSelectedStyleSheet"
		Call Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "settoproperty", objPropeties.JavaStaticText("jstx_BottomLink"),"","label",dictProperties("PropertyStyleSheetName"))		
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaStaticText("jstx_BottomLink"), "","","") Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictProperties("PropertyStyleSheetName")) & " ] style sheet is currently selected on properties dialog","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyStyleSheetName")) & " ] style sheet does not available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "VerifyLastModifiedTimeGreaterThanCreatedTime"
		If DateDiff("s",dictProperties("InitialModifiedDate"),dictProperties("LastModifiedDate"))>0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified last modified time is greater ( current time ) than created time","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as last modified time is not greater ( current time ) than created time","","","","","")
			Call Fn_ExitTest()
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "OpenListBoxItemAndGetProperties", "OpenListBoxItemAndVerifyProperties"
		'Set label property of java static text
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",dictProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		
		'Verify existence of java list
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objPropeties.JavaList("jlst_PropertyField"),"","","") Then
			'Activate the list item
			If Cint(objPropeties.JavaList("jlst_PropertyField").GetRoProperty("items count")) > 0 Then
				If isNumeric(dictProperties("PropertyValue")) Then
					sDialogTitle = objPropeties.JavaList("jlst_PropertyField").GetItem(dictProperties("PropertyValue"))
					objPropeties.JavaList("jlst_PropertyField").Activate dictProperties("PropertyValue")
				Else
					'Add code here if user specifies value rather than the node index to be selected
				End If
			Else
				aProperty = Split(dictProperties("AdditionalProperties_FieldNames"),"~")
				For iCounter = 0 To Ubound(aProperty) Step 1
					If iCounter = 0 Then
						sTempValue = "False"
					Else
						sTempValue = sTempValue & "~False"
					End If
				Next
				
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				
				GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER = Datatable.GetCurrentRow
				Datatable.SetCurrentRow 1
				DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_VerifyProperties"
				DataTable.Value("ReusableActionWordReturnValue","Global")= sTempValue
				DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
				
				'Clicking on button
				If sButtonName<>"" Then
					IF Fn_UI_JavaButton_Operations("Fn_VerifyProperties", "Click", objPropeties, "jbtn_" & sButtonName)=False then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as fail to click on [ " & Cstr(sButtonName) & " ] button of properties dialog","","","","","")
						Call Fn_ExitTest()
					End IF
				End If
				
				ExitAction
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Verify existence of read only dialog
		Set objReadOnly = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ReadOnly","")
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objReadOnly,"","","") Then
			IF Fn_UI_JavaButton_Operations("Fn_VerifyProperties", "Click", objReadOnly, "jbtn_Yes")=False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on Yes button on read only dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		Set objReadOnly = Nothing
		
		'Set title property of dialog to the item activated
		Set objAdditionalPropertiesDialog = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_AddtionalPropertiesDialog","")
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objAdditionalPropertiesDialog,"","title",sDialogTitle)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set TO property of additonal properties dialog as [" & sDialogTitle & "]","","","","","")
			Call Fn_ExitTest()
		End IF
		
		'Verify existence of additional properties dialog
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objAdditionalPropertiesDialog,"","","") =False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence of additional properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Based on the type of field to be validated select the case
		Select Case dictProperties("AdditionalProperties_FieldType")
		
			Case "EditBox"
				'Loop to iterate through each property name provided
				aProperty = Split(dictProperties("AdditionalProperties_FieldNames"),"~")
				For iCounter = 0 To Ubound(aProperty) Step 1
					'If action if the get the values and not to verify the properties
					If sAction = "OpenListBoxItemAndGetProperties" Then
						'Set the editbox attached text property 
						If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objAdditionalPropertiesDialog.JavaEdit("jedt_AdditionalPropertiesEdit"),"","attached text",aProperty(iCounter) & ":")=False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set TO property of editbox on additonal properties dialog as [" & aProperty(iCounter) & "]","","","","","")
							Call Fn_ExitTest()
						End IF
						'Get values of editboxes adn store in temporary variable
						If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objPropeties,2,"","") Then
							If iCounter = 0 Then
								sTempValue = Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "gettext", objAdditionalPropertiesDialog.JavaEdit("jedt_AdditionalPropertiesEdit"),"", "" )
							Else
								sTempValue = sTempValue & "~" & Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "gettext", objAdditionalPropertiesDialog.JavaEdit("jedt_AdditionalPropertiesEdit"),"", "" )
							End If
						Else
							If iCounter = 0 Then
								sTempValue = "False"
							Else
								sTempValue = sTempValue & "~False"
							End If
						End IF
					End If
				Next
				
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
				
				GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER = Datatable.GetCurrentRow
				Datatable.SetCurrentRow 1
				DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_VerifyProperties"
				DataTable.Value("ReusableActionWordReturnValue","Global")= sTempValue
				DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
				
		End Select
		
		'Close the additional propertie dialog
		If dictProperties("AdditionalProperties_CloseDialog") = True Then
			IF Fn_UI_JavaButton_Operations("Fn_VerifyProperties", "Click", objAdditionalPropertiesDialog, "jbtn_Close")=False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as fail to click on [ Close ] button of properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "verifyeditboxvaluenotempty"   
		aProperty=Split(dictProperties("PropertyName"),"~")
		
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
	
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
			    aValues=Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "GetText", objPropeties, "jedt_PropertyField", "")
				If  aValues <> "" Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property is not empty and contains value [ " & Cstr(aValues) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property value is empty","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "verifyeditboxvalueisempty"   
		aProperty=Split(dictProperties("PropertyName"),"~")
		
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
	
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
			    aValues=Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "GetText", objPropeties, "jedt_PropertyField", "")
				If  Trim(aValues) = "" Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property is empty","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property value is not empty and contains value [ " & Cstr(aValues) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "verifylistboxisempty"   
		aProperty=Split(dictProperties("PropertyName"),"~")
		
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
	
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaList("jlst_PropertyField"),"","","") Then
			    aValues=CInt(objPropeties.JavaList("jlst_PropertyField").GetROProperty("items count"))
				If  aValues = 0 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property is empty\does not contains any value","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property value is not empty\contains some values","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "GetPropertiesValue"
		aProperty=Split(dictProperties("PropertyName"),"~")
		sTempValue=""		
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to get value of Property [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If objPropeties.JavaEdit("jedt_PropertyField").Exist(1) Then
				If sTempValue="" Then
					sTempValue=Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "GetText", objPropeties, "jedt_PropertyField", "")
				Else
					sTempValue=sTempValue & "^" & Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "GetText", objPropeties, "jedt_PropertyField", "")
				End If
			ElseIf objPropeties.JavaList("jlst_PropertyField").Exist(0) Then
				If sTempValue="" Then
					'sTempValue=objPropeties.JavaList("jlst_PropertyField").GetROProperty("Value")
					sTempValue=objPropeties.JavaList("jlst_PropertyField").GetItem(0)
				Else
					'sTempValue=sTempValue & "^" & objPropeties.JavaList("jlst_PropertyField").GetROProperty("Value")
					sTempValue=sTempValue & "^" & objPropeties.JavaList("jlst_PropertyField").GetItem(0)
				End If
			Elseif objPropeties.JavaStaticText("jstx_PropertyLabelValue").Exist(1) Then
				If sTempValue="" Then
					sTempValue=objPropeties.JavaStaticText("jstx_PropertyLabelValue").GetROProperty("label")
				Else
					sTempValue=sTempValue & "^" & objPropeties.JavaStaticText("jstx_PropertyLabelValue").GetROProperty("label")
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to get value of Property [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_VerifyProperties"
		DataTable.Value("ReusableActionWordReturnValue","Global")=sTempValue
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify property links
	Case "ClickLink"		
		sTempPropertyName=dictProperties("PropertyName")
		dictProperties("PropertyName")=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",dictProperties("PropertyName"),"")
		
		If dictProperties("PropertyName")="False" Then
			dictProperties("PropertyName")=sTempPropertyName
		End If
		
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",dictProperties("PropertyName") & ":")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click link of property [ " & Cstr(dictProperties("PropertyName")) & " ] as property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		
		If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","Exist",objPropeties.JavaStaticText("jstx_PropertyLabelValue"),"","","") Then
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","getroproperty",objPropeties.JavaStaticText("jstx_PropertyLabelValue"),"","label","")=dictProperties("PropertyValue") Then
				objPropeties.JavaStaticText("jstx_PropertyLabelValue").Click 1,1
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click link of property [ " & Cstr(dictProperties("PropertyName")) & " ] as property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click link of property [ " & Cstr(dictProperties("PropertyName")) & " ] as property does not exist\available on properties dialog","","","","","")
			Call Fn_ExitTest()
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	' Case to click on specific button of properties dialog
	Case "ClickButton"		
	'Nothing do anything
	
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on link of property [ " & Cstr(dictProperties("PropertyName")) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on link of property [ " & Cstr(dictProperties("PropertyName")) & " ]","","","","","")
		End If
	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "VerifySecondCreatedObjectLastModifiedTimeGreaterThanFirstCreatedObjectLastModifiedTime"
		If DateDiff("s",dictProperties("FirstCreatedObjectLastModifiedDate"),dictProperties("SecondCreatedObjectLastModifiedDate"))>0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified last modified time of second created object is greater than first created object last modified time","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as last modified time of second created object is not greater than first created object last modified time","","","","","")
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "EditBox_VerifyDate"
		aProperty=Split(dictProperties("PropertyName"),"~")
		aValues=Split(dictProperties("PropertyValue"),"~")
				
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If aValues(iCounter)="<<EMPTY>>" Then
				aValues(iCounter)=""
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
				If Cdate(Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "GetText", objPropeties, "jedt_PropertyField", ""))=Cdate(aValues(iCounter)) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If	
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "EditBox_VerifyDateFormat"
		aProperty=Split(dictProperties("PropertyName"),"~")
		aValues=Split(dictProperties("DateFormat"),"~")
				
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF
			
			If aValues(iCounter)="<<EMPTY>>" Then
				aValues(iCounter)=""
			End If						
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
				sTempValue=Fn_UI_JavaEdit_Operations("RAC_Common_VerifyProperties", "GetText", objPropeties, "jedt_PropertyField", "")
				Select Case aValues(iCounter)
					Case "DD-MMM-YYYY"
						sTempValue=Split(sTempValue,"-")
						If Len(sTempValue(0))=2 And Len(sTempValue(1))=3 And Len(sTempValue(2))=4 Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] is in [ " & Cstr(aValues(iCounter)) & " ] format.","","","","DONOTSYNC","")
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] is not in [ " & Cstr(aValues(iCounter)) & " ] format","","","","","")
							Call Fn_ExitTest()
						End If
				End Select
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
			If Err.Number <> 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] is not in [ " & Cstr(aValues(iCounter)) & " ] format","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "ListBox_VerifyTimeFormat"
		aProperty=Split(dictProperties("PropertyName"),"~")
		aValues=Split(dictProperties("TimeFormat"),"~")
				
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties","settoproperty",objPropeties.JavaStaticText("jstx_PropertyLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End IF				
			
			If Fn_UI_Object_Operations("RAC_Common_VerifyProperties", "Exist", objPropeties.JavaEdit("jedt_PropertyField"),"","","") Then
				sTempValue=objPropeties.JavaList("jlst_PropertyField").GetROProperty("value")
				Select Case aValues(iCounter)
					Case "24HOUR"
						sDate=Split(sTempValue,":")
						If Len(sTempValue)=5 and Cint(sDate(0))>=0 and Cint(sDate(0))<24 and Cint(sDate(1))<60 Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] time is in [ " & Cstr(aValues(iCounter)) & " ] format.","","","","DONOTSYNC","")
						Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] time is not in [ " & Cstr(aValues(iCounter)) & " ] format","","","","","")
							Call Fn_ExitTest()
						End If
				End Select
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on properties dialog","","","","","")
				Call Fn_ExitTest()
			End If
			If Err.Number <> 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] time is not in [ " & Cstr(aValues(iCounter)) & " ] format","","","","","")
				Call Fn_ExitTest()
			End If
		Next
End	Select

'Clicking on button
If sButtonName<>"" Then
	IF Fn_UI_JavaButton_Operations("Fn_VerifyProperties", "Click", objPropeties, "jbtn_" & sButtonName)=False then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & "] as fail to click on [ " & Cstr(sButtonName) & " ] button of properties dialog","","","","","")
		Call Fn_ExitTest()
	End IF
End If

'Capture execution end time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Verify Properties",sAction,dictProperties("PropertyName"),dictProperties("PropertyValue"))

'Releasing object of [ Properties ] dialog
Set objPropeties=Nothing

Function Fn_ExitTest()
	'Releasing object of [ Properties ] dialog
	Set objPropeties=Nothing
	ExitTest
End Function

