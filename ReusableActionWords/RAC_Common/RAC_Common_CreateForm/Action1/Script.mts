'! @Name 			RAC_Common_CreateForm
'! @Details 		Action word to perform operations on New Form creation dialog. eg. Form basic create , item detail create
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam3 	sFormType			: Form Type
'! @InputParam4 	sInvokeOption		: New Form creation dialog invoke option
'! @InputParam6 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			12 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateForm","RAC_Common_CreateForm",OneIteration,"autocreatecustomerpartformwithallfields","FormType_CustomerPartForm","menu",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateForm","RAC_Common_CreateForm",OneIteration,"autobasiccreate","FormType_ItemMaster","menu",""

Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sFormType,sInvokeOption,sButton
Dim sCustomerPartFormNum,sCurrentFormName,sTempFormType,sFormName
Dim iFormCount,iCounter,iFormNodeCount,iCount
Dim objNewForm
Dim bFlag,sPerspective
Dim aValue

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sFormType = Parameter("sFormType")
sInvokeOption = Parameter("sInvokeOption")
sButton = Parameter("sButton")

sTempFormType=sFormType

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Invoking [ New Form ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileNewForm"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating object of [ New Form ] dialog
Select Case Lcase(sPerspective)
	Case "myteamcenter","my teamcenter","structuremanager","structure manager"
		'Creating object of [ New Form ]
		Set objNewForm=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_NewForm","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreateForm"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of new Form dialog
If Fn_UI_Object_Operations("RAC_Common_CreateForm","Exist", objNewForm, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ New Form ] creation dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting Form count
If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="create" or Lcase(sAction)="createwithoutclose" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACFormCount","","")
	iFormCount=Fn_CommonUtil_DataTableOperations("GetValue","RACFormCount","","")
	If iFormCount="" Then
		iFormCount=1
	Else
		iFormCount=iFormCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACFormCount",iFormCount,"")
End If

'Get actual item type name
If sFormType<>"" Then
	sFormType=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewFormValues_APL",sFormType,""))
	sTempFormType=sFormType
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_CreateForm",sAction,"","")

If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="create" or Lcase(sAction)="createwithoutclose" or Lcase(sAction)="verifylovvalues" or Lcase(sAction)="verifydefaultvalue" or Lcase(sAction)="modifyeditboxvalue" or Lcase(sAction)="seteditboxvalueandverifylength" Then
	DataTable.SetCurrentRow iFormCount
	If sFormType<>"" Then
		'Select Form type
		iFormNodeCount=Fn_UI_Object_Operations("RAC_Common_CreateForm","GetROProperty",objNewForm.JavaTree("jtree_FormType"),"","items count","")
		For iCounter=0 To iFormNodeCount-1
			sCurrentFormName = objNewForm.JavaTree("jtree_FormType").GetItem(iCounter)
			If Trim(sCurrentFormName)="Most Recently Used~" & Trim(sFormType) Then
				sFormType = "Most Recently Used~" & Trim(sFormType)
				bFlag=True
				Exit For
			ElseIf Trim(sCurrentFormName)="Complete List~" & Trim(sFormType) Then
				sFormType = "Complete List~" & Trim(sFormType)
				bFlag=True
				Exit For
			End If
		Next
		
		If bFlag = True Then
			If Fn_UI_JavaTree_Operations("RAC_Common_CreateForm","Select",objNewForm,"jtree_FormType",sFormType,"","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Form type [ " & Cstr(sFormType) & " ] from new Form creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Form type [ " & Cstr(sFormType) & " ] from new Form creation dialog as specified Form type does not exist in list","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on next button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm,"jbtn_Next") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button from new Form creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
	sFormType=sTempFormType
	
	DataTable.SetCurrentRow iFormCount
End If

Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic Form with standard values
	Case "autobasiccreate"		
		'Set Form Name
		sFormName = Fn_Setup_GenerateObjectInformation("getname",sFormType)
		Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label","Name:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "Set",  objNewForm, "jedt_FormEdit", sFormName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to set Form name vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
					
		'Set Form description		
		Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label","Description:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "Set",  objNewForm, "jedt_FormEdit", sFormName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to set Form description vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
				
		'Click on Finish button
		objNewForm.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateForm",sAction,"","")
				
		'Store Form Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACFormName","","")
		DataTable.Value("RACFormName","Global") = sFormName						
				
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Form of type [ " & Cstr(sFormType) & " ] with Form name [ " & Cstr(Datatable.Value("RACFormName", "Global")) & " ]","","","","","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create customer part form with all values
	Case "create","createwithoutclose"
		'Set Form Name
		If dictFormInfo("FormName")<>"" Then
			sFormName=dictFormInfo("FormName")
			Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label","Customer Part Name:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "Set",  objNewForm, "jedt_FormEdit", sFormName )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to set Form Customer Part Name vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
		End If
		
		'Set Form description
		If dictFormInfo("Description")<>"" Then	
			Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label","Description:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "Set",  objNewForm, "jedt_FormEdit", dictFormInfo("Description") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to set Form description vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
		End If
	
		'Selecting Customer name
		If dictFormInfo("Customer")<>"" Then
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)			
			Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label","Customer:")
			'If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_LOVDropDown")=False Then
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "Set",  objNewForm, "jedt_FormEdit", dictFormInfo("Customer") ) =False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to select customer vlaue","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			wait 2
			
			If Lcase(dictFormInfo("Customer"))="abc" Then
				dictFormInfo("Customer")="BMW"
			End If
			
			'If Fn_UI_JavaTree_Operations("RAC_Common_CreateForm","Doubleclick",objNewForm.JavaWindow("jwnd_LOVTreeShell"),"jtee_LOVTree",dictFormInfo("Customer"),"","")=False Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_LOVDropDown")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to select [ Customer ] value [ " & Cstr(dictFormInfo("Customer")) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If			
		
		'Selecting Is Serviceable
		If dictFormInfo("IsServiceable")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label","Is Serviceable:")
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_LOVDropDown")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to select Is Serviceable vlaue","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTree_Operations("RAC_Common_CreateForm","Doubleclick",objNewForm.JavaWindow("jwnd_LOVTreeShell"),"jtee_LOVTree",dictFormInfo("IsServiceable"),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to select [ Is Serviceable ] value [ " & Cstr(dictFormInfo("IsServiceable")) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		'Set form ng5_customer_part_form_num
		If dictFormInfo("CustomerPartNumber")<>"" Then			
			sCustomerPartFormNum=dictFormInfo("CustomerPartNumber")
			Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label","Customer Part Number:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "Set",  objNewForm, "jedt_FormEdit",sCustomerPartFormNum )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to set [ ng5_customer_part_form_num ] value [ " & Cstr(sCustomerPartFormNum) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		
		'Click on Finish button
		objNewForm.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		If LCase(sAction)<>"createwithoutclose" Then
			'Click on Close button
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_Close")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] as fail to click on [ Close ] button","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateForm",sAction,"","")

		'Store Form Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACFormName","","")
		DataTable.Value("RACFormName","Global") = sFormName						
				
		'Store Form Description
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACFormDescription","","")
		DataTable.Value("RACFormDescription","Global") = dictFormInfo("Description")	
		
		'Store Form Revision
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACFormPartRevision","","")
		DataTable.Value("RACFormPartRevision","Global") = "-"	
		
		'Store Form Part Number
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACFormPartNumber","","")
		DataTable.Value("RACFormPartNumber","Global") = sCustomerPartFormNum
		
		'Store Form Customer
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACFormCustomer","","")
		DataTable.Value("RACFormCustomer","Global") = dictFormInfo("Customer")
						
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Form of type [ " & Cstr(sFormType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Form of type [ " & Cstr(sFormType) & " ] with Form name [ " & Cstr(Datatable.Value("RACFormName", "Global")) & " ]","","","","","")
		End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "clickbutton"
		objNewForm.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_" & sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of [ Form ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)				
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify LOV value
	Case "verifylovvalues"
		Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label",dictFormInfo("PropertyName") & ":")
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_LOVDropDown")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as fail to click on LOV drop down button of [ " & Cstr(dictFormInfo("PropertyName")) & " ] on [ New Form ] creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		aValue=Split(dictFormInfo("PropertyValue"),"~")
		For iCounter=0 to Ubound(aValue)
			bFlag=False
			For iCount=0 to objNewForm.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtee_LOVTree").GetROProperty("items count")-1
				If aValue(iCounter)=objNewForm.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtee_LOVTree").GetItem(iCount) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictFormInfo("PropertyName")) & " ] property contains [ " & Cstr(aValue(iCounter)) & " ] value on [ New Form ] creation dialog","","","","DONOTSYNC","")
					bFlag=True
					Exit For	
				End If
			Next
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictFormInfo("PropertyName")) & " ] property does not contain value [ " & Cstr(aValue(iCounter)) & " ] on [ New Form ] creation dialog","","","","","")
				Call Fn_ExitTest()
				Exit For
			End If
		Next
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_LOVDropDown")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as fail to click on LOV drop down button of [ " & Cstr(dictFormInfo("PropertyName")) & " ] on [ New Form ] creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		If sButton<>"" Then
			objNewForm.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of [ Form ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify default value
	Case "verifydefaultvalue"
		Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label",dictFormInfo("PropertyName") & ":")
		If Cstr(Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "gettext",  objNewForm, "jedt_FormEdit","" ))=Cstr(dictFormInfo("PropertyValue")) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictFormInfo("PropertyName")) & " ] property by default contains [ " & Cstr(dictFormInfo("PropertyValue")) & " ] value on [ New Form ] creation dialog","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictFormInfo("PropertyName")) & " ] property by default does not contain value [ " & Cstr(dictFormInfo("PropertyValue")) & " ] on [ New Form ] creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sButton<>"" Then
			objNewForm.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of [ Form ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to modify value
	Case "modifyeditboxvalue"
		Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label",dictFormInfo("PropertyName") & ":")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "Set",  objNewForm, "jedt_FormEdit",dictFormInfo("PropertyValue"))=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to modify [ " & Cstr(dictFormInfo("PropertyName")) & " ] field value to [ " & Cstr(dictFormInfo("PropertyValue")) & " ] on [ New Form ] dialog","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully modified [ " & Cstr(dictFormInfo("PropertyName")) & " ] field value to [ " & Cstr(dictFormInfo("PropertyValue")) & " ] on [ New Form ] dialog","","","","","")
		End If
		
		If sButton<>"" Then
			objNewForm.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of [ Form ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
   	Case "seteditboxvalueandverifylength"		
		Call Fn_UI_Object_Operations("RAC_Common_CreateForm","SetTOProperty", objNewForm.JavaStaticText("jstx_FormLabel"),"","label",dictFormInfo("PropertyName") & ":")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "Set",  objNewForm, "jedt_FormEdit",dictFormInfo("PropertyValue"))=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set value of editbox [ " & Cstr(dictFormInfo("PropertyName")) & " ] field value to [ " & Cstr(dictFormInfo("PropertyValue")) & " ] on [ New Form ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		If Cint(Len(Fn_UI_JavaEdit_Operations("RAC_Common_CreateForm", "gettext",  objNewForm, "jedt_FormEdit", "")))=Cint(dictFormInfo("Length")) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as editbox [ " & Cstr(dictFormInfo("PropertyName")) & " ] has string value of length equal to [ " & dictFormInfo("Length") & " ]","","","","DONOTSYNC","")	
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as editbox [ " & Cstr(dictFormInfo("PropertyName")) & " ] has string value of length not equal to [ " & dictFormInfo("Length") & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sButton<>"" Then
			objNewForm.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateForm", "Click", objNewForm, "jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of [ Form ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If	
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing object of New Form dialog
Set objNewForm =Nothing

Function Fn_ExitTest()
	Set objNewForm =Nothing
	ExitTest
End Function
