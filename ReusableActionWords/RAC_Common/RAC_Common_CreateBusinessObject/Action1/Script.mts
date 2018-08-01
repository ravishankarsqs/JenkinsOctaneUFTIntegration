'! @Name 			RAC_Common_CreateBusinessObject
'! @Details 		Action word to perBusinessObject operations on New Business Object creation dialog
'! @InputParam1 	sAction 				: String to indicate what action is to be perBusinessObjected
'! @InputParam3 	sBusinessObjectType		: BusinessObject Type
'! @InputParam4 	sInvokeOption			: New BusinessObject creation dialog invoke option
'! @InputParam6 	sButton 				: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			12 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateBusinessObject","RAC_Common_CreateBusinessObject",OneIteration,"autobasiccreate","","menu",""

Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sBusinessObjectType,sInvokeOption,sButton,sExternalId,sExternalName
Dim sCustomerPartBusinessObjectNum,sCurrentBusinessObjectName,sTempBusinessObjectType,sBusinessObjectName
Dim iBusinessObjectCount,iCounter,iBusinessObjectNodeCount
Dim objNewBusinessObject,objDefaultWindow
Dim sPerspective,sSupportDesignNode,sEngineeredDrawingNode,sMirroredORNode
dim objSupportdesigntable,objEngineereddrawingtable,objmirroredortable
Dim bFlag
Dim objNewItem

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sBusinessObjectType = Parameter("sBusinessObjectType")
sInvokeOption = Parameter("sInvokeOption")
sButton = Parameter("sButton")

sTempBusinessObjectType=sBusinessObjectType

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Creating object of Teamcenter Default Window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_DefaultWindow","")

'creating object of support design table in Linked Data tab
Set objSupportdesigntable=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jtbl_LinkedDataSupportDesigns","")

'creating object of engineered drawing table in Linked Data tab
Set objEngineereddrawingtable=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jtbl_LinkedDataEnginneredDrawing","")	

'creating object of mirrored or handed parts table in Linked Data tab
Set objmirroredortable=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jtbl_LinkedDataMirroredORHandedParts","")	


'Invoking [ New BusinessObject ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileNewOther"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating object of [ New BusinessObject ] dialog
Select Case LCase(sPerspective)
	Case "myteamcenter"
		'Creating object of [ New BusinessObject ]
		Set objNewBusinessObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_NewBusinessObject","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreateBusinessObject"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
'Checking existance of new BusinessObject dialog
If sBusinessObjectType<> "" Then
   If Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","Exist", objNewBusinessObject, GBL_DEFAULT_TIMEOUT,"","")=False Then
	   Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perBusinessObject operation [ " & Cstr(sAction) & " ] on [ New BusinessObject ] creation dialog as dialog does not exist","","","","","")
	   Call Fn_ExitTest()
    End If
End If
'Setting BusinessObject count
If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="create" or Lcase(sAction)="createwithoutclose" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBusinessObjectCount","","")
	iBusinessObjectCount=Fn_CommonUtil_DataTableOperations("GetValue","RACBusinessObjectCount","","")
	If iBusinessObjectCount="" Then
		iBusinessObjectCount=1
	Else
		iBusinessObjectCount=iBusinessObjectCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACBusinessObjectCount",iBusinessObjectCount,"")
End If

'Get actual item type name
If sBusinessObjectType<>"" Then
	sBusinessObjectType=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewBusinessObjectValues_APL",sBusinessObjectType,""))
	sTempBusinessObjectType=sBusinessObjectType
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_CreateBusinessObject",sAction,"","")

If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="create" or Lcase(sAction)="createwithoutclose" Then
	DataTable.SetCurrentRow iBusinessObjectCount
	If sBusinessObjectType<>"" Then
		'Select BusinessObject type
		iBusinessObjectNodeCount=Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","GetROProperty",objNewBusinessObject.JavaTree("jtree_BusinessObjectType"),"","items count","")
		For iCounter=0 To iBusinessObjectNodeCount-1
			sCurrentBusinessObjectName = objNewBusinessObject.JavaTree("jtree_BusinessObjectType").GetItem(iCounter)
			If Trim(sCurrentBusinessObjectName)="Most Recently Used~" & Trim(sBusinessObjectType) Then
				sBusinessObjectType = "Most Recently Used~" & Trim(sBusinessObjectType)
				bFlag=True
				Exit For
			ElseIf Trim(sCurrentBusinessObjectName)="Complete List~" & Trim(sBusinessObjectType) Then
				sBusinessObjectType = "Complete List~" & Trim(sBusinessObjectType)
				bFlag=True
				Exit For
			End If
		Next
		
		If bFlag = True Then
			If Fn_UI_JavaTree_Operations("RAC_Common_CreateBusinessObject","Select",objNewBusinessObject,"jtree_BusinessObjectType",sBusinessObjectType,"","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select BusinessObject type [ " & Cstr(sBusinessObjectType) & " ] from new BusinessObject creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select BusinessObject type [ " & Cstr(sBusinessObjectType) & " ] from new BusinessObject creation dialog as specified BusinessObject type does not exist in list","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on next button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject,"jbtn_Next") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button from new BusinessObject creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
	sBusinessObjectType=sTempBusinessObjectType
	
	DataTable.SetCurrentRow iBusinessObjectCount
End If

Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic BusinessObject with standard values
	Case "autobasiccreate"		
		'Set BusinessObject Name
		sBusinessObjectName = Fn_Setup_GenerateObjectInBusinessObjectation("getname",sBusinessObjectType)
		Call Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","SetTOProperty", objNewBusinessObject.JavaStaticText("jstx_BusinessObjectLabel"),"","label","Name:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateBusinessObject", "Set",  objNewBusinessObject, "jedt_BusinessObjectEdit", sBusinessObjectName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to set BusinessObject name vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
					
		'Set BusinessObject description		
		Call Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","SetTOProperty", objNewBusinessObject.JavaStaticText("jstx_BusinessObjectLabel"),"","label","Description:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateBusinessObject", "Set",  objNewBusinessObject, "jedt_BusinessObjectEdit", sBusinessObjectName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to set BusinessObject description vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
				
		'Click on Finish button
		objNewBusinessObject.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateBusinessObject",sAction,"","")
				
		'Store BusinessObject Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBusinessObjectName","","")
		DataTable.Value("RACBusinessObjectName","Global") = sBusinessObjectName						
				
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] with BusinessObject name [ " & Cstr(Datatable.Value("RACBusinessObjectName", "Global")) & " ]","","","","","")
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create customer part BusinessObject with all values
	Case "create","createwithoutclose"	
		'Set BusinessObject Name
		If dictBusinessObjectInfo("CustomerPartName")<>"" Then
			sBusinessObjectName = dictBusinessObjectInfo("CustomerPartName")
			Call Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","SetTOProperty", objNewBusinessObject.JavaStaticText("jstx_BusinessObjectLabel"),"","label","Customer Part Name:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateBusinessObject", "Set",  objNewBusinessObject, "jedt_BusinessObjectEdit", sBusinessObjectName )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to set BusinessObject Customer Part Name vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
		End IF
		
		If dictBusinessObjectInfo("Description")<>"" Then
			'Set BusinessObject description		
			Call Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","SetTOProperty", objNewBusinessObject.JavaStaticText("jstx_BusinessObjectLabel"),"","label","Description:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateBusinessObject", "Set",  objNewBusinessObject, "jedt_BusinessObjectEdit", dictBusinessObjectInfo("Description") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to set BusinessObject description vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
		End If
		
		If dictBusinessObjectInfo("Customer")<>"" Then
			'Selecting Customer name
			Call Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","SetTOProperty", objNewBusinessObject.JavaStaticText("jstx_BusinessObjectLabel"),"","label","Customer:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateBusinessObject", "Set",  objNewBusinessObject, "jedt_BusinessObjectEdit", dictBusinessObjectInfo("Customer"))=False Then
			'If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject, "jbtn_LOVDropDown")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to select customer vlaue","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			wait 2
			
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject, "jbtn_LOVDropDown")=False Then
			'If Fn_UI_JavaTree_Operations("RAC_Common_CreateBusinessObject","Doubleclick",objNewBusinessObject.JavaWindow("jwnd_LOVTreeShell"),"jtee_LOVTree",dictBusinessObjectInfo("Customer"),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to select [ Customer ] value [ " & Cstr(dictBusinessObjectInfo("Customer")) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End IF
		
		If dictBusinessObjectInfo("IsServiceable")<>"" Then
			'Selecting Is Serviceable
			Call Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","SetTOProperty", objNewBusinessObject.JavaStaticText("jstx_BusinessObjectLabel"),"","label","Is Serviceable:")
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject, "jbtn_LOVDropDown")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to select Is Serviceable vlaue","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If Fn_UI_JavaTree_Operations("RAC_Common_CreateBusinessObject","Doubleclick",objNewBusinessObject.JavaWindow("jwnd_LOVTreeShell"),"jtee_LOVTree",dictBusinessObjectInfo("IsServiceable"),"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to select [ Is Serviceable ] value [ " & Cstr(dictBusinessObjectInfo("IsServiceable")) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End IF
		
		'Set BusinessObject ng5_customer_part_BusinessObject_num			
		If dictBusinessObjectInfo("CustomerPartNumber")<>"" Then
			sCustomerPartBusinessObjectNum=dictBusinessObjectInfo("CustomerPartNumber")
			Call Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","SetTOProperty", objNewBusinessObject.JavaStaticText("jstx_BusinessObjectLabel"),"","label","Customer Part Number:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateBusinessObject", "Set",  objNewBusinessObject, "jedt_BusinessObjectEdit",sCustomerPartBusinessObjectNum )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to set [ ng5_customer_part_BusinessObject_num ] value [ " & Cstr(sCustomerPartBusinessObjectNum) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		End IF
		
		'Click on Finish button
		objNewBusinessObject.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		If LCase(sAction)<>"createwithoutclose" Then
			'Click on Close button
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject, "jbtn_Close")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] as fail to click on [ Close ] button","","","","","")
				Call Fn_ExitTest()
			End If
		End IF
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateBusinessObject",sAction,"","")
				
		'Store BusinessObject Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBusinessObjectName","","")
		DataTable.Value("RACBusinessObjectName","Global") = sBusinessObjectName						
				
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created BusinessObject of type [ " & Cstr(sBusinessObjectType) & " ] with BusinessObject name [ " & Cstr(Datatable.Value("RACBusinessObjectName", "Global")) & " ]","","","","","")
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "clickbutton"
		objNewBusinessObject.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject, "jbtn_" & sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of [ Business Object ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to verify default value
	Case "verifydefaultvalue"
		Call Fn_UI_Object_Operations("RAC_Common_CreateBusinessObject","SetTOProperty", objNewBusinessObject.JavaStaticText("jstx_BusinessObjectLabel"),"","label",dictItemInfo("PropertyName") & ":")
		If Cstr(Fn_UI_JavaEdit_Operations("RAC_Common_CreateBusinessObject", "gettext",  objNewBusinessObject, "jedt_BusinessObjectEdit","" ))=Cstr(dictItemInfo("PropertyValue")) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(dictItemInfo("PropertyName")) & " ] property by default contains [ " & Cstr(dictItemInfo("PropertyValue")) & " ] value on [ New BuisnessObject ] creation dialog","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictItemInfo("PropertyName")) & " ] property by default does not contain value [ " & Cstr(dictItemInfo("PropertyValue")) & " ] on [ New BuisnessObject ] creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sButton<>"" Then
			objNewBusinessObject.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objNewBusinessObject, "jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of [ BuisnessObject ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing object of New BusinessObject dialog
Set objNewBusinessObject =Nothing

Function Fn_ExitTest()
	Set objNewBusinessObject =Nothing
	ExitTest
End Function

