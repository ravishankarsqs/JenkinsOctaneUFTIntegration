'! @Name 			RAC_Common_CreateChange
'! @Details 		Action word to perform operations on New Change creation dialog. eg. Change basic create , item detail create
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam2 	sParentNode			: Parent node under which user wants to create change
'! @InputParam3 	sChangeType			: Change Type
'! @InputParam4 	sInvokeOption		: New Change creation dialog invoke option
'! @InputParam5 	sPerspective	 	: Perspective name in which user wants to perform operations on New Change dialog
'! @InputParam6 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			12 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateChange","RAC_Common_CreateChange",OneIteration,"autobasiccreate","ChangeType_ProblemReport","menu","myteamcenter",""


Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sParentNode,sChangeType,sInvokeOption,sPerspective,sButton
Dim sChangeRevisionID,sChangeID,sCurrentChangeName,sTempChangeType,sChangeName
Dim iChangeCount,iCounter,iChangeNodeCount
Dim objNewChange
Dim aChangeInfo
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sParentNode = Parameter("sParentNode")
sChangeType = Parameter("sChangeType")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sButton = Parameter("sButton")

sTempChangeType=sChangeType

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Invoking [ New Change ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileNewChange"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

If sPerspective="" Then
	sPerspective=Fn_RAC_GetActivePerspectiveName("")
End If

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreateChange"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Creating object of [ New Change ] dialog
Select Case sPerspective
	Case "myteamcenter"
		'Creating object of [ New Change ]
		Set objNewChange=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_NewChange","")
		Set objLOVTree =objNewChange.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtee_LOVTree")
End Select

'Checking existance of new Change dialog
If Fn_UI_Object_Operations("RAC_Common_CreateChange","Exist", objNewChange, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ New Change ] creation dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting Change count
If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="basiccreate" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeCount","","")
	iChangeCount=Fn_CommonUtil_DataTableOperations("GetValue","RACChangeCount","","")
	If iChangeCount="" Then
		iChangeCount=1
	Else
		iChangeCount=iChangeCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACChangeCount",iChangeCount,"")
End If

'Get actual item type name
If sChangeType<>"" Then
	sChangeType=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",sChangeType,""))
	sTempChangeType=sChangeType
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_CreateChange",sAction,"","")

If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="basiccreate" Then
	DataTable.SetCurrentRow iChangeCount
	If sChangeType<>"" Then
		'Select Change type
		iChangeNodeCount=Fn_UI_Object_Operations("RAC_Common_CreateChange","GetROProperty", objNewChange.JavaTree("jtree_ChangeType"),"","Items count","")
		For iCounter=0 To iChangeNodeCount-1
			sCurrentChangeName = objNewChange.JavaTree("jtree_ChangeType").GetChange(iCounter)
			If Trim(sCurrentChangeName)="Most Recently Used~" & Trim(sChangeType) Then
				sChangeType = "Most Recently Used~" & Trim(sChangeType)
				bFlag=True
				Exit For
			ElseIf Trim(sCurrentChangeName)="Complete List~" & Trim(sChangeType) Then
				sChangeType = "Complete List~" & Trim(sChangeType)
				bFlag=True
				Exit For
			End If
		Next
		
		If bFlag = True Then
			If Fn_UI_JavaTree_Operations("RAC_Common_CreateChange","Select",objNewChange,"jtree_ChangeType",sChangeType,"","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Change type [ " & Cstr(sChangeType) & " ] from new Change creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Change type [ " & Cstr(sChangeType) & " ] from new Change creation dialog as specified Change type does not exist in list","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on next button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateChange", "Click", objNewChange,"jbtn_Next") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button from new Change creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End If
	sChangeType=sTempChangeType
	
	DataTable.SetCurrentRow iChangeCount
End If

Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic Change with standard values
	Case "autobasiccreate"		
		'Set Change Name
		sChangeName = Fn_Setup_GenerateObjectInformation("getname",sChangeType)
		Call Fn_UI_Object_Operations("RAC_Common_CreateChange","SetTOProperty", objNewChange.JavaStaticText("jstx_ChangeLabel"),"","label","Synopsis:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateChange", "Set",  objNewChange, "jedt_ChangeEdit", sChangeName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] as fail to set Change name\synopsis vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
		'Set Change description		
		Call Fn_UI_Object_Operations("RAC_Common_CreateChange","SetTOProperty", objNewChange.JavaStaticText("jstx_ChangeLabel"),"","label","Description:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateChange", "Set",  objNewChange, "jedt_ChangeEdit", sChangeName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] as fail to set Change description vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"GetChildrenByName",sParentNode & "^" & sChangeName,""
		DataTable.SetCurrentRow 1		
		
		If DataTable.Value("ReusableActionWordReturnValue","Global")="False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new change of type [ " & Cstr(sChangeType) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		
		aChangeInfo=Split(DataTable.Value("ReusableActionWordReturnValue","Global"),"-")
		DataTable.SetCurrentRow iChangeCount
		sChangeID=aChangeInfo(0) & "-" & aChangeInfo(1)
			
		'Form the Change node name in nav tree and store in Datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeNode","","")
		DataTable.Value("RACChangeNode","Global") = sChangeID & "-" & sChangeName
		
		'Store nav tree revision node details in datatable
		sChangeRevisionID="A"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeRevisionNode","","")
		DataTable.Value("RACChangeRevisionNode","Global") = sChangeID & "/" & sChangeRevisionID & ";1-" & sChangeName
		
		'Store Change ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeID","","")
		DataTable.Value("RACChangeID","Global") = sChangeID
		
		'Store Change Revision ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeRevisionID","","")
		DataTable.Value("RACChangeRevisionID","Global") = sChangeRevisionID
		
		'Store Change Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeName","","")
		DataTable.Value("RACChangeName","Global") = sChangeName
				
		'Click on Finish button
		objNewChange.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateChange", "Click", objNewChange, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateChange", "Click", objNewChange, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateChange",sAction,"","")
		
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Change of type [ " & Cstr(sChangeType) & " ] with Change Id [ " & Cstr(Datatable.Value("RACChangeID", "Global")) & " ] , Change Revision Id [ " & Cstr(Datatable.Value("RACChangeRevisionID", "Global")) & " ] and Change name [ " & Cstr(Datatable.Value("RACChangeName", "Global")) & " ]","","","","","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic Change with standard values
	Case "basiccreate"		
		'Set Change Name
		If dictItemInfo("Synopsis")="Assign" Then
			sChangeName = Fn_Setup_GenerateObjectInformation("getname",sChangeType)
		Else
			sChangeName = dictItemInfo("Synopsis")
		End If
		Call Fn_UI_Object_Operations("RAC_Common_CreateChange","SetTOProperty", objNewChange.JavaStaticText("jstx_ChangeLabel"),"","label","Synopsis:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateChange", "Set",  objNewChange, "jedt_ChangeEdit", sChangeName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] as fail to set Change name\synopsis vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
		'Set Change description
		If dictItemInfo("Description")<>"" Then	
			Call Fn_UI_Object_Operations("RAC_Common_CreateChange","SetTOProperty", objNewChange.JavaStaticText("jstx_ChangeLabel"),"","label","Description:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateChange", "Set",  objNewChange, "jedt_ChangeEdit", dictItemInfo("Description") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] as fail to set Change description vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations","RAC_MyTc_NavigationTreeOperations",OneIteration,"GetChildrenByName",sParentNode & "^" & sChangeName,""
		DataTable.SetCurrentRow 1		
		
		If DataTable.Value("ReusableActionWordReturnValue","Global")="False" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new change of type [ " & Cstr(sChangeType) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		
		aChangeInfo=Split(DataTable.Value("ReusableActionWordReturnValue","Global"),"-")
		DataTable.SetCurrentRow iChangeCount
		sChangeID=aChangeInfo(0) & "-" & aChangeInfo(1)
			
		'Form the Change node name in nav tree and store in Datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeNode","","")
		DataTable.Value("RACChangeNode","Global") = sChangeID & "-" & sChangeName
		
		'Store nav tree revision node details in datatable
		sChangeRevisionID="A"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeRevisionNode","","")
		DataTable.Value("RACChangeRevisionNode","Global") = sChangeID & "/" & sChangeRevisionID & ";1-" & sChangeName
		
		'Store Change ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeID","","")
		DataTable.Value("RACChangeID","Global") = sChangeID
		
		'Store Change Revision ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeRevisionID","","")
		DataTable.Value("RACChangeRevisionID","Global") = sChangeRevisionID
		
		'Store Change Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACChangeName","","")
		DataTable.Value("RACChangeName","Global") = sChangeName
				
		'Click on Finish button
		objNewChange.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateChange", "Click", objNewChange, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateChange", "Click", objNewChange, "jbtn_Close")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateChange",sAction,"","")
		
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Change of type [ " & Cstr(sChangeType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created Change of type [ " & Cstr(sChangeType) & " ] with Change Id [ " & Cstr(Datatable.Value("RACChangeID", "Global")) & " ] , Change Revision Id [ " & Cstr(Datatable.Value("RACChangeRevisionID", "Global")) & " ] and Change name [ " & Cstr(Datatable.Value("RACChangeName", "Global")) & " ]","","","","","")
		End If		
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing object of New Change dialog
Set objNewChange =Nothing

Function Fn_ExitTest()
	Set objNewChange =Nothing
	ExitTest
End Function

