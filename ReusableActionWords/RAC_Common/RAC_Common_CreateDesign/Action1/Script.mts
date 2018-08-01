'! @Name 			RAC_Common_CreateDesign
'! @Details 		Action word to perform operations on New Design creation dialog. eg. Design basic create , Design detail create
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam2 	sDesignType 		: Design Type
'! @InputParam3 	sInvokeOption		: New Design creation dialog invoke option
'! @InputParam4 	sPerspective	 	: Perspective name in which user wants to perform operations on New Design dialog
'! @InputParam5 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			07 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateDesign","RAC_Common_CreateDesign",OneIteration,"autobasiccreate","DesignType_Design","menu","myteamcenter",""

Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sDesignType,sInvokeOption,sPerspective,sButton
Dim sName,sDesignRevisionID,sDesignID
Dim objNewDesign,objChildObjects
Dim iDesignCount,iCounter
Dim DictItems,DictKeys

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sDesignType = Parameter("sDesignType")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sButton = Parameter("sButton")

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Invoking [ New Design ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","NewDesign"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

If sPerspective="" Then
	sPerspective=Fn_RAC_GetActivePerspectiveName("")
End If

'Creating object of [ New Design ] dialog
Select Case sPerspective
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter","","structuremanager"
		Set objNewDesign=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jdlg_NewDesign","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreateDesign"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of new Design dialog
If Fn_UI_Object_Operations("RAC_Common_CreateDesign","Exist", objNewDesign, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ New Design ] creation dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting Design count
If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="basiccreate" or Lcase(sAction)="detailcreate" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignCount","","")
	iDesignCount=Fn_CommonUtil_DataTableOperations("GetValue","RACDesignCount","","")
	If iDesignCount="" Then
		iDesignCount=1
	Else
		iDesignCount=iDesignCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACDesignCount",iDesignCount,"")	
End If

'Get actual Design type name
If sDesignType<>"" Then
	sDesignType=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewDesignValues_APL",sDesignType,""))
End If

'Capture execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Design Details Create",sAction,"","")

If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="basiccreate" or Lcase(sAction)="detailcreate" Then
	DataTable.SetCurrentRow iDesignCount
	If sDesignType<>"" Then
		'Selecting Design Type from list
		If Fn_UI_JavaList_Operations("RAC_Common_CreateDesign", "Select", objNewDesign,"jlst_DesignType",sDesignType, "", "")=False Then		
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Design type [ " & CStr(sDesignType) & " ] from Design type list while creating new Design","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		'Clicking On Next button
		objNewDesign.JavaButton("jbtn_Next").WaitProperty "enabled", 1, 60000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign,"jbtn_Next")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button from new Design creation dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
End If
		
Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic Design with standard values
	Case "autobasiccreate"							
		'Click on Assign button		
		If Fn_UI_Object_Operations("RAC_Common_CreateDesign","getroproperty", objNewDesign.JavaButton("jbtn_Assign"),"","enabled","")=1 Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign,"jbtn_Assign")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ Assign ] button to assign Design id and revision","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ Assign ] button to assign Design id and revision","","","","","")
			Call Fn_ExitTest()
		End If		
		'Fetch Design ID
		sDesignID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign","GetText",objNewDesign,"jedt_DesignID", "" )
		'Fetch Revision ID	
		sDesignRevisionID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign","GetText",objNewDesign,"jedt_RevisionID","" )
		
		sName = Fn_Setup_GenerateObjectInformation("getname",sDesignType)
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign", "Set",  objNewDesign, "jedt_Name", sName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to set Design name vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign", "Set",  objNewDesign, "jedt_Description", sName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to set Design description vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Store Design ID and Design Revision 
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignNode","","")
		DataTable.Value("RACDesignNode","Global") = sDesignID & "-" & sName
		
		'Store nav tree revision node details in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignRevisionNode","","")
		DataTable.Value("RACDesignRevisionNode","Global") = sDesignID & "/" & sDesignRevisionID & ";1-" & sName
		
		'Store Design ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignID","","")
		DataTable.Value("RACDesignID","Global") = sDesignID
		
		'Store Design Revision ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignRevisionID","","")
		DataTable.Value("RACDesignRevisionID","Global") = sDesignRevisionID
		
		'Store Design Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignName","","")
		DataTable.Value("RACDesignName","Global") = sName
		
		'Click on Finish button
		objNewDesign.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateDesign",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Design of type [ " & Cstr(sDesignType) & " ] with Design Id [ " & Cstr(Datatable.Value("RACDesignID", "Global")) & " ] , Design Revision Id [ " & Cstr(Datatable.Value("RACDesignRevisionID", "Global")) & " ] and Design name [ " & Cstr(Datatable.Value("RACDesignName", "Global")) & " ]","","","","","")
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic Design with custom values
	Case "basiccreate"		
		'Assign item ID
		If dictDesignInfo("ID")="" or dictDesignInfo("ID")="Assign" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign,"jbtn_Assign")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ Assign ] button to assign Design id and revision","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Else
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign", "Set",  objNewDesign, "jedt_DesignID", dictDesignInfo("ID") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sItemType) & " ] as fail to set Design id vlaue","","","","","")
				Call Fn_ExitTest()
			End If
			If dictDesignInfo("Revision")<>"" Then
				If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign", "Set",  objNewDesign, "jedt_RevisionID", dictDesignInfo("Revision") )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sItemType) & " ] as fail to set Design Revision id vlaue","","","","","")
					Call Fn_ExitTest()
				End If
			End If
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		'Fetch Design ID
		sDesignID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign","GetText",objNewDesign,"jedt_DesignID", "" )
		'Fetch Revision ID	
		sDesignRevisionID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign","GetText",objNewDesign,"jedt_RevisionID","" )
		
		'Setting Design name
		If dictDesignInfo("Name")<>"" Then
			IF dictDesignInfo("Name")="Assign" Then
				sName = Fn_Setup_GenerateObjectInformation("getname",sDesignType)
			Else
				sName=dictDesignInfo("Name")
			End If
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign", "Set",  objNewDesign, "jedt_Name", sName )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to set Design name vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		'Setting Design description	
		If dictDesignInfo("Description")<>"" Then
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign", "Set",  objNewDesign, "jedt_Description", dictDesignInfo("Description") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to set Design description vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End IF
		
		'Setting Unit of measure
		If dictDesignInfo("Unit of Measure")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_CreateDesign","SetTOProperty", objNewDesign.JavaStaticText("jstx_DesignLabel"),"","label","Unit of Measure:")
			Call Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign, "jbtn_LOVDropDown")
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			Set objChildObjects = Fn_UI_Object_GetChildObjects("RAC_Common_CreateDesign", objNewDesign, "Class Name~label", "JavaStaticText~" & dictDesignInfo("Unit of Measure"))
			If isObject(objChildObjects) Then
				objChildObjects(0).Click 1,1
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to set Design Unit of Measure vlaue","","","","","")
				Call Fn_ExitTest()
			End If
			Set objChildObjects = Nothing
		End If
		
		'Store Design ID and Design Revision 
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignNode","","")
		DataTable.Value("RACDesignNode","Global") = sDesignID & "-" & sName
		
		'Store nav tree revision node details in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignRevisionNode","","")
		DataTable.Value("RACDesignRevisionNode","Global") = sDesignID & "/" & sDesignRevisionID & ";1-" & sName
		
		'Store Design ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignID","","")
		DataTable.Value("RACDesignID","Global") = sDesignID
		
		'Store Design Revision ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignRevisionID","","")
		DataTable.Value("RACDesignRevisionID","Global") = sDesignRevisionID
		
		'Store Design Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignName","","")
		DataTable.Value("RACDesignName","Global") = sName
		
		'Click on Finish button
		objNewDesign.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateDesign",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Design of type [ " & Cstr(sDesignType) & " ] with Design Id [ " & Cstr(Datatable.Value("RACDesignID", "Global")) & " ] , Design Revision Id [ " & Cstr(Datatable.Value("RACDesignRevisionID", "Global")) & " ] and Design name [ " & Cstr(Datatable.Value("RACDesignName", "Global")) & " ]","","","","","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic Design with custom values
	Case "detailcreate"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_CreateDesign"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		
		DictItems = dictDesignInfo.Items
		DictKeys = dictDesignInfo.Keys
		For iCounter=0 to dictDesignInfo.count-1
			Select Case DictKeys(iCounter)
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				'Assign or Enter ID & Revision
				Case "ID","Revision"
					If LCase(DictItems(iCounter))="assign" Then
						'Click on Assign button		
						If Fn_UI_Object_Operations("RAC_Common_CreateDesign","getroproperty", objNewDesign.JavaButton("jbtn_Assign"),"","enabled","")=1 Then
							If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign,"jbtn_Assign")=False Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ Assign ] button to assign Design id and revision","","","","","")
								Call Fn_ExitTest()
							End If
							Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
						End If
					Else
						If DictKeys(iCounter)="ID" Then
							If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign", "Set",  objNewDesign, "jedt_DesignID", dictDesignInfo("ID") )=False Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sItemType) & " ] as fail to set Design id value","","","","","")
								Call Fn_ExitTest()
							End If
						Else
							If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign", "Set",  objNewDesign, "jedt_RevisionID", dictDesignInfo("Revision") )=False Then
								Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sItemType) & " ] as fail to set Design Revision id value","","","","","")
								Call Fn_ExitTest()
							End If
						End If
						Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					End If
					
					If DictKeys(iCounter)="ID" Then
						sDesignID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign","GetText",objNewDesign,"jedt_DesignID", "" )
					Else
						sDesignRevisionID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign","GetText",objNewDesign,"jedt_RevisionID","" )
					End If
				'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Name"
					If DictKeys(iCounter)="Name" Then
						If DictItems(iCounter)="Assign" Then
							DictItems(iCounter)=Fn_Setup_GenerateObjectInformation("getname",sDesignType)
						End If
						sName=DictItems(iCounter)
					End If
					If Fn_UI_JavaEdit_Operations("RAC_Common_CreateDesign", "Set",  objNewDesign, DictKeys(iCounter), DictItems(iCounter) )=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to set Design [ " & Cstr(DictKeys(iCounter)) & " ] value","","","","","")
						Call Fn_ExitTest()
					End If									
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -						
				Case "Description"
					Call Fn_UI_Object_Operations("RAC_Common_CreateDesign","SetTOProperty", objNewDesign.JavaStaticText("jstx_DesignLabel"),"","label",DictKeys(iCounter) & ":")
					If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "Set",  objNewDesign, "jedt_DesignEdit", DictItems(iCounter) )=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to set Design [ " & Cstr(DictKeys(iCounter)) & " ] value","","","","","")
						Call Fn_ExitTest()
					End If									
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -				
				Case "Unit of Measure"
					Call Fn_UI_Object_Operations("RAC_Common_CreateDesign","SetTOProperty", objNewDesign.JavaStaticText("jstx_DesignLabel"),"","label",DictKeys(iCounter) & ":")
					Call Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign, "jbtn_LOVDropDown")
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Set objChildObjects = Fn_UI_Object_GetChildObjects("RAC_Common_CreateDesign", objNewDesign, "Class Name~label", "JavaStaticText~" & DictItems(iCounter))
					If isObject(objChildObjects) Then
						objChildObjects(0).Click 1,1
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to set Design [ " & Cstr(DictKeys(iCounter)) & " ] value","","","","","")
						Call Fn_ExitTest()
					End If
					Set objChildObjects = Nothing
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -		
				'Click on Next Button
				Case "Next","Next@1","Next@2","Next@3"
					objNewDesign.JavaButton("jbtn_Next").WaitProperty "enabled", 1, 60000
					If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign,"jbtn_Next")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button from new Design creation dialog","","","","","")
						Call Fn_ExitTest()
					End IF
					Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
			End Select
		Next
		
		'Click on button
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ " & Cstr(sButton) & " ] button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		If sDesignID<>"" and sDesignRevisionID<>"" and sName<>""Then
			'Store Design ID and Design Revision 
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignNode","","")
			DataTable.Value("RACDesignNode","Global") = sDesignID & "-" & sName
			
			'Store nav tree revision node details in datatable
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignRevisionNode","","")
			DataTable.Value("RACDesignRevisionNode","Global") = sDesignID & "/" & sDesignRevisionID & ";1-" & sName
			
			'Store Design ID
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignID","","")
			DataTable.Value("RACDesignID","Global") = sDesignID
			
			'Store Design Revision ID
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignRevisionID","","")
			DataTable.Value("RACDesignRevisionID","Global") = sDesignRevisionID
			
			'Store Design Name
			Call Fn_CommonUtil_DataTableOperations("AddColumn","RACDesignName","","")
			DataTable.Value("RACDesignName","Global") = sName
		Else
			DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
		End If
			
		If sButton="Finish" Then
			If Fn_UI_Object_Operations("RAC_Common_CreateDesign","Exist", objNewDesign,"","","")=True Then
				'Click on Close button
				If Fn_UI_JavaButton_Operations("RAC_Common_CreateDesign", "Click", objNewDesign, "jbtn_Close")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] as fail to click on [ Close ] button","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			End If	
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateDesign",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Design of type [ " & Cstr(sDesignType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			If sDesignID<>"" and sDesignRevisionID<>"" and sName<>""Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Design of type [ " & Cstr(sDesignType) & " ] with Design Id [ " & Cstr(Datatable.Value("RACDesignID", "Global")) & " ] , Design Revision Id [ " & Cstr(Datatable.Value("RACDesignRevisionID", "Global")) & " ] and Design name [ " & Cstr(Datatable.Value("RACDesignName", "Global")) & " ]","","","","","")
			End If
		End If
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing object of New Design dialog
Set objNewDesign =Nothing

Function Fn_ExitTest()
	'Releasing object of New Design dialog
	Set objNewDesign =Nothing
	ExitTest
End Function

