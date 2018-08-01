'! @Name 			RAC_Common_CreatePart
'! @Details 		Action word to perform operations on New Part creation dialog. eg. Part basic create , Part detail create
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam2 	sPartType 			: Part Type
'! @InputParam3 	sInvokeOption		: New Part creation dialog invoke option
'! @InputParam4 	sPerspective	 	: Perspective name in which user wants to perform operations on New Part dialog
'! @InputParam5 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			07 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreatePart","RAC_Common_CreatePart",OneIteration,"autobasiccreate","PartType_Part","menu","myteamcenter",""

Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sPartType,sInvokeOption,sPerspective,sButton
Dim sName,sPartRevisionID,sPartID
Dim objNewPart,objChildObjects
Dim iPartCount,iCounter
Dim DictItems,DictKeys

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sPartType = Parameter("sItemType")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sButton = Parameter("sButton")

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Invoking [ New Part ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileNewPart"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

If sPerspective="" Then	
	'Get active perspective name
	sPerspective=Fn_RAC_GetActivePerspectiveName("")
End If
'Creating object of [ New Part ] dialog
Select Case sPerspective
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter","","structuremanager"
		Set objNewPart=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jdlg_NewPart","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreatePart"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of new Part dialog
If Fn_UI_Object_Operations("RAC_Common_CreatePart","Exist", objNewPart, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ New Part ] creation dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting part count
If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="basiccreate" or Lcase(sAction)="detailcreate" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartCount","","")
	iPartCount=Fn_CommonUtil_DataTableOperations("GetValue","RACPartCount","","")
	If iPartCount="" Then
		iPartCount=1
	Else
		iPartCount=iPartCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACPartCount",iPartCount,"")	
End If

'Get actual Part type name
If sPartType<>"" Then
	sPartType=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewPartValues_APL",sPartType,""))
End If

'Capture execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Part Details Create",sAction,"","")

If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="basiccreate" or Lcase(sAction)="detailcreate" Then
	DataTable.SetCurrentRow iPartCount
	If sPartType<>"" Then
		'Selecting Part Type from list
		If Fn_UI_JavaList_Operations("RAC_Common_CreatePart", "Select", objNewPart,"jlst_PartType",sPartType, "", "")=False Then		
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select part type [ " & CStr(sPartType) & " ] from part type list while creating new part","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		'Clicking On Next button
		objNewPart.JavaButton("jbtn_Next").WaitProperty "enabled", 1, 60000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart,"jbtn_Next")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button from new part creation dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
End If
		
Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic part with standard values
	Case "autobasiccreate"							
		'Click on Assign button		
		If Fn_UI_Object_Operations("RAC_Common_CreatePart","getroproperty", objNewPart.JavaButton("jbtn_Assign"),"","enabled","")=1 Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart,"jbtn_Assign")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to click on [ Assign ] button to assign part id and revision","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to click on [ Assign ] button to assign part id and revision","","","","","")
			Call Fn_ExitTest()
		End If		
		'Fetch Part ID
		sPartID = Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart","GetText",objNewPart,"jedt_PartID", "" )
		'Fetch Revision ID	
		sPartRevisionID = Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart","GetText",objNewPart,"jedt_RevisionID","" )
		
		sName = Fn_Setup_GenerateObjectInformation("getname",sPartType)
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart", "Set",  objNewPart, "jedt_Name", sName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to set part name vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart", "Set",  objNewPart, "jedt_Description", sName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to set part description vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Store Part ID and Part Revision 
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartNode","","")
		DataTable.Value("RACPartNode","Global") = sPartID & "-" & sName
		
		'Store nav tree revision node details in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartRevisionNode","","")
		DataTable.Value("RACPartRevisionNode","Global") = sPartID & "/" & sPartRevisionID & ";1-" & sName
		
		'Store Part ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartID","","")
		DataTable.Value("RACPartID","Global") = sPartID
		
		'Store Part Revision ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartRevisionID","","")
		DataTable.Value("RACPartRevisionID","Global") = sPartRevisionID
		
		'Store Part Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartName","","")
		DataTable.Value("RACPartName","Global") = sName
		
		'Click on Finish button
		objNewPart.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Part of type [ " & Cstr(sPartType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Part of type [ " & Cstr(sPartType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreatePart",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Part of type [ " & Cstr(sPartType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Part of type [ " & Cstr(sPartType) & " ] with Part Id [ " & Cstr(Datatable.Value("RACPartID", "Global")) & " ] , Part Revision Id [ " & Cstr(Datatable.Value("RACPartRevisionID", "Global")) & " ] and Part name [ " & Cstr(Datatable.Value("RACPartName", "Global")) & " ]","","","","","")
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic part with custom values
	Case "basiccreate"		
		'Assign item ID
		If dictPartInfo("ID")="" or dictPartInfo("ID")="Assign" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart,"jbtn_Assign")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to click on [ Assign ] button to assign part id and revision","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Else
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart", "Set",  objNewPart, "jedt_PartID", dictPartInfo("ID") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sItemType) & " ] as fail to set part id vlaue","","","","","")
				Call Fn_ExitTest()
			End If
			If dictPartInfo("Revision")<>"" Then
				If Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart", "Set",  objNewPart, "jedt_RevisionID", dictPartInfo("Revision") )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sItemType) & " ] as fail to set part Revision id vlaue","","","","","")
					Call Fn_ExitTest()
				End If
			End If
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		'Fetch Part ID
		sPartID = Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart","GetText",objNewPart,"jedt_PartID", "" )
		'Fetch Revision ID	
		sPartRevisionID = Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart","GetText",objNewPart,"jedt_RevisionID","" )
		
		'Setting part name
		If dictPartInfo("Name")<>"" Then
			IF dictPartInfo("Name")="Assign" Then
				sName = Fn_Setup_GenerateObjectInformation("getname",sPartType)
			Else
				sName=dictPartInfo("Name")
			End If
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart", "Set",  objNewPart, "jedt_Name", sName )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to set part name vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		'Setting part description	
		If dictPartInfo("Description")<>"" Then
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart", "Set",  objNewPart, "jedt_Description", dictPartInfo("Description") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to set part description vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End IF
		
		'Setting Unit of measure
		If dictPartInfo("Unit of Measure")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_CreatePart","SetTOProperty", objNewPart.JavaStaticText("jstx_PartLabel"),"","label","Unit of Measure:")
			Call Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart, "jbtn_LOVDropDown")
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Set objChildObjects = Fn_UI_Object_GetChildObjects("RAC_Common_CreatePart", objNewPart, "Class Name~label", "JavaStaticText~" & dictPartInfo("Unit of Measure"))
			If isObject(objChildObjects) Then
				objChildObjects(0).Click 1,1
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to set part Unit of Measure vlaue","","","","","")
				Call Fn_ExitTest()
			End If
			Set objChildObjects = Nothing
		End If
		
		'Store Part ID and Part Revision 
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartNode","","")
		DataTable.Value("RACPartNode","Global") = sPartID & "-" & sName
		
		'Store nav tree revision node details in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartRevisionNode","","")
		DataTable.Value("RACPartRevisionNode","Global") = sPartID & "/" & sPartRevisionID & ";1-" & sName
		
		'Store Part ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartID","","")
		DataTable.Value("RACPartID","Global") = sPartID
		
		'Store Part Revision ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartRevisionID","","")
		DataTable.Value("RACPartRevisionID","Global") = sPartRevisionID
		
		'Store Part Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartName","","")
		DataTable.Value("RACPartName","Global") = sName
		
		'Click on Finish button
		objNewPart.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Part of type [ " & Cstr(sPartType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Part of type [ " & Cstr(sPartType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreatePart",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Part of type [ " & Cstr(sPartType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Part of type [ " & Cstr(sPartType) & " ] with Part Id [ " & Cstr(Datatable.Value("RACPartID", "Global")) & " ] , Part Revision Id [ " & Cstr(Datatable.Value("RACPartRevisionID", "Global")) & " ] and Part name [ " & Cstr(Datatable.Value("RACPartName", "Global")) & " ]","","","","","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic part with custom values
	Case "detailcreate"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_CreatePart"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		
		DictItems = dictPartInfo.Items
		DictKeys = dictPartInfo.Keys
		For iCounter=0 to dictPartInfo.count-1
			Select Case DictKeys(iCounter)
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Assign or Enter ID & Revision
				Case "ID","Revision"
					If LCase(DictItems(iCounter))="assign" Then
						'Click on Assign button		
						If Fn_UI_Object_Operations("RAC_Common_CreatePart","getroproperty", objNewPart.JavaButton("jbtn_Assign"),"","enabled","")=1 Then
							If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart,"jbtn_Assign")=False Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to click on [ Assign ] button to assign part id and revision","","","","","")
								Call Fn_ExitTest()
							End If
							Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
						End If
					Else
						If DictKeys(iCounter)="ID" Then
							If Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart", "Set",  objNewPart, "jedt_PartID", dictPartInfo("ID") )=False Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sItemType) & " ] as fail to set part id value","","","","","")
								Call Fn_ExitTest()
							End If
						Else
							If Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart", "Set",  objNewPart, "jedt_RevisionID", dictPartInfo("Revision") )=False Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sItemType) & " ] as fail to set part Revision id value","","","","","")
								Call Fn_ExitTest()
							End If
						End If
						Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					End If
					
					If DictKeys(iCounter)="ID" Then
						sPartID = Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart","GetText",objNewPart,"jedt_PartID", "" )
					Else
						sPartRevisionID = Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart","GetText",objNewPart,"jedt_RevisionID","" )
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name"
					If DictKeys(iCounter)="Name" Then
						If DictItems(iCounter)="Assign" Then
							DictItems(iCounter)=Fn_Setup_GenerateObjectInformation("getname",sPartType)
						End If
						sName=DictItems(iCounter)
					End If
					If Fn_UI_JavaEdit_Operations("RAC_Common_CreatePart", "Set",  objNewPart, DictKeys(iCounter), DictItems(iCounter) )=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to set part [ " & Cstr(DictKeys(iCounter)) & " ] value","","","","","")
						Call Fn_ExitTest()
					End If									
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
				Case "Description"
					Call Fn_UI_Object_Operations("RAC_Common_CreatePart","SetTOProperty", objNewPart.JavaStaticText("jstx_PartLabel"),"","label",DictKeys(iCounter) & ":")
					If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "Set",  objNewPart, "jedt_PartEdit", DictItems(iCounter) )=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to set part [ " & Cstr(DictKeys(iCounter)) & " ] value","","","","","")
						Call Fn_ExitTest()
					End If									
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -				
				Case "Unit of Measure"
					Call Fn_UI_Object_Operations("RAC_Common_CreatePart","SetTOProperty", objNewPart.JavaStaticText("jstx_PartLabel"),"","label",DictKeys(iCounter) & ":")
					Call Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart, "jbtn_LOVDropDown")
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Set objChildObjects = Fn_UI_Object_GetChildObjects("RAC_Common_CreatePart", objNewPart, "Class Name~label", "JavaStaticText~" & DictItems(iCounter))
					If isObject(objChildObjects) Then
						objChildObjects(0).Click 1,1
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new part of type [ " & Cstr(sPartType) & " ] as fail to set part [ " & Cstr(DictKeys(iCounter)) & " ] value","","","","","")
						Call Fn_ExitTest()
					End If
					Set objChildObjects = Nothing
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -		
				'Click on Next Button
				Case "Next","Next@1","Next@2","Next@3"
					objNewPart.JavaButton("jbtn_Next").WaitProperty "enabled", 1, 60000
					If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart,"jbtn_Next")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button from new part creation dialog","","","","","")
						Call Fn_ExitTest()
					End IF
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
			End Select
		Next
		
		'Click on button
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Part of type [ " & Cstr(sPartType) & " ] as fail to click on [ " & Cstr(sButton) & " ] button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		If sPartID<>"" and sPartRevisionID<>"" and sName<>""Then
			'Store Part ID and Part Revision 
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartNode","","")
			DataTable.Value("RACPartNode","Global") = sPartID & "-" & sName
			
			'Store nav tree revision node details in datatable
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartRevisionNode","","")
			DataTable.Value("RACPartRevisionNode","Global") = sPartID & "/" & sPartRevisionID & ";1-" & sName
			
			'Store Part ID
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartID","","")
			DataTable.Value("RACPartID","Global") = sPartID
			
			'Store Part Revision ID
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartRevisionID","","")
			DataTable.Value("RACPartRevisionID","Global") = sPartRevisionID
			
			'Store Part Name
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACPartName","","")
			DataTable.Value("RACPartName","Global") = sName
		Else
			DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
		End If
			
		If sButton="Finish" Then
			If Fn_UI_Object_Operations("RAC_Common_CreatePart","Exist", objNewPart,"","","")=True Then
				'Click on Close button
				If Fn_UI_JavaButton_Operations("RAC_Common_CreatePart", "Click", objNewPart, "jbtn_Close")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Part of type [ " & Cstr(sPartType) & " ] as fail to click on [ Close ] button","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			End If	
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreatePart",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Part of type [ " & Cstr(sPartType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			If sPartID<>"" and sPartRevisionID<>"" and sName<>""Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Part of type [ " & Cstr(sPartType) & " ] with Part Id [ " & Cstr(Datatable.Value("RACPartID", "Global")) & " ] , Part Revision Id [ " & Cstr(Datatable.Value("RACPartRevisionID", "Global")) & " ] and Part name [ " & Cstr(Datatable.Value("RACPartName", "Global")) & " ]","","","","","")
			End If
		End If
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing object of New Part dialog
Set objNewPart =Nothing

Function Fn_ExitTest()
	'Releasing object of New Part dialog
	Set objNewPart =Nothing
	ExitTest
End Function


