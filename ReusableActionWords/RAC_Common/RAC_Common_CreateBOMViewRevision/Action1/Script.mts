'! @Name 			RAC_Common_CreateBOMViewRevision
'! @Details 		Action word to perform operations on new BOMViewRevision creation dialog.
'! @InputParam1 	sAction					: Action Name
'! @InputParam2		sInvokeOption			: New BOMViewRevision creation dialog invoke option
'! @InputParam3		sBOMViewRevisionType	: BOMViewRevision type
'! @InputParam4		sID						: ID
'! @InputParam5		sRevisionID				: Revision ID
'! @InputParam6		sName					: Name
'! @InputParam7		sCreateAs				: Create As option
'! @InputParam8		bOpenOnCreate			: Open on create option
'! @InputParam9		sButton					: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			19 Jul 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateBOMViewRevision","RAC_Common_CreateBOMViewRevision",OneIteration, "autobasiccreate","Menu","","","","","","",""

Option Explicit
Err.Clear

'Declaring varaibles

'Variable Declaration
Dim sAction,sInvokeOption,sBOMViewRevisionType,sID,sRevisionID,sName,sCreateAs,bOpenOnCreate,sButton
Dim objNewBOMViewRevision,objChildObjects
Dim sPerspective
Dim iBOMViewRevisionCount,iCounter

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sBOMViewRevisionType = Parameter("sBOMViewRevisionType")
sID = Parameter("sID")
sRevisionID = Parameter("sRevisionID")
sName = Parameter("sName")
sCreateAs = Parameter("sCreateAs")
bOpenOnCreate = Parameter("bOpenOnCreate")
sButton = Parameter("sButton")

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Invoking [ New BOMViewRevision ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",oneIteration,"Select","FileNewBOMViewRevision"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating Object of [ New BOMViewRevision ] Dialog
Select Case Lcase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "myteamcenter",""
		Set objNewBOMViewRevision=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_NewBOMViewRevision","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreateBOMViewRevision"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			
'Checking existance of new BOMViewRevision dialog
If Fn_UI_Object_Operations("RAC_Common_CreateBOMViewRevision","Exist", objNewBOMViewRevision,GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ New BOMViewRevision ] creation dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting BOMViewRevision count
If Lcase(sAction)="autobasiccreate" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBOMViewRevisionCount","","")
	iBOMViewRevisionCount=Fn_CommonUtil_DataTableOperations("GetValue","RACBOMViewRevisionCount","","")
	If iBOMViewRevisionCount="" Then
		iBOMViewRevisionCount=1
	Else
		iBOMViewRevisionCount=iBOMViewRevisionCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACBOMViewRevisionCount",iBOMViewRevisionCount,"")
	DataTable.SetCurrentRow iBOMViewRevisionCount
End If

If Lcase(sAction)= "create" or Lcase(sAction)="autobasiccreate" Then
	'Clickng on [ More... ] Option to Choose BOMViewRevision Type
	If Fn_UI_JavaCheckBox_Operations("RAC_Common_CreateBOMViewRevision", "Set", objNewBOMViewRevision, "jckb_More", "ON")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BOMViewRevision of type [ " & CStr(sBOMViewRevisionType) & " ] as fail to click on [ More ] option from new BOMViewRevision creation dialog","","","","","")
		Call Fn_ExitTest()
	End If
	Call Fn_RAC_ReadyStatusSync(2)
	
	'Set BOMViewRevision Type
	Set objChildObjects = Fn_UI_Object_GetChildObjects("RAC_Common_CreateBOMViewRevision",objNewBOMViewRevision,"Class Name~label","JavaStaticText~" & sBOMViewRevisionType)
	objChildObjects(0).Click 1,1
	Set objChildObjects =Nothing
	
	If Err.Number <> 0 Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BOMViewRevision of type [ " & CStr(sBOMViewRevisionType) & " ] as specified BOMViewRevision type does not exist on new BOMViewRevision creation dialog","","","","","")
		Call Fn_ExitTest()
	End If
	Call Fn_RAC_ReadyStatusSync(1)
End If

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to Create New BOMViewRevision	
	Case "autobasiccreate"				
		sID = Cstr(Fn_UI_JavaEdit_Operations("RAC_Common_CreateBOMViewRevision","GetText",objNewBOMViewRevision,"jedt_ItemID","" ))
		sRevisionID = Cstr(Fn_UI_JavaEdit_Operations("RAC_Common_CreateBOMViewRevision","GetText",objNewBOMViewRevision,"jedt_RevisionID","" ))
		sName = Cstr(Fn_UI_JavaEdit_Operations("RAC_Common_CreateBOMViewRevision","GetText",objNewBOMViewRevision,"jedt_Name","" ))
		
		If sCreateAs<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_CreateBOMViewRevision","SetTOProperty",objNewBOMViewRevision.JavaRadioButton("jrdb_CreateAs"),"","attached text",sCreateAs)
			If Fn_UI_JavaRadioButton_Operations("RAC_Common_CreateBOMViewRevision","Set",objNewBOMViewRevision,"jrdb_CreateAs","ON")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BOMViewRevision of type [ " & CStr(sBOMViewRevisionType) & " ] as fail to set Create As option [ " & Cstr(sCreateAs) & " ] on new BOMViewRevision creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End IF
		
		'Clicking On OK Button to create BOMViewRevision
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBOMViewRevision", "Click",objNewBOMViewRevision,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new BOMViewRevision of type [ " & CStr(sBOMViewRevisionType) & " ] as fail to click on [ OK ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created BOMViewRevision of type [ " & CStr(sBOMViewRevisionType) & " ] and ID as [ " & Cstr(sID) & " ]","","","","","")
		
		'Set value of new BOMViewRevision name in BOMViewRevisionName column
		DataTable.SetCurrentRow iBOMViewRevisionCount
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBOMViewRevisionID","","")
		DataTable.Value("RACBOMViewRevisionID","Global") = sID
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBOMViewRevisionRevisionID","","")
		DataTable.Value("RACBOMViewRevisionRevisionID","Global") = sRevisionID
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBOMViewRevisionName","","")
		DataTable.Value("RACBOMViewRevisionName","Global") = sName
		
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACBOMViewRevisionNode","","")
		DataTable.Value("RACBOMViewRevisionNode","Global") = sID & "/" & sRevisionID & "-View"
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Relasing Object of [ New BOMViewRevision ] Dialog
Set objNewBOMViewRevision=Nothing

Public Function Fn_ExitTest()
	'Relasing Object of [ New BOMViewRevision ] Dialog
	Set objNewBOMViewRevision=Nothing
	Set objImportFile=Nothing
	Set objUploadFile=Nothing
	ExitTest
End Function
