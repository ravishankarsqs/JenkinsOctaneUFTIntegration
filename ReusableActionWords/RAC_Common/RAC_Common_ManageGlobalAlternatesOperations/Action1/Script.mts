'! @Name 			RAC_Common_ManageGlobalAlternatesOperations
'! @Details 		Action word to perform operations on Manage Global Alternates dialog.
'! @InputParam1 	sAction			: Action Name
'! @InputParam2		sInvokeOption	: Open By Name dialog invoke option
'! @InputParam3		sName			: Object Name
'! @InputParam4		sID				: Object ID
'! @InputParam5		sObject			: Object Details
'! @InputParam6		sColumnName		: Column name
'! @InputParam7		sValue			: Column Value
'! @InputParam8		sButton			: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			12 Jul 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ManageGlobalAlternatesOperations","RAC_Common_ManageGlobalAlternatesOperations",OneIteration,"Add","Menu","","123456","","","",""

Option Explicit
Err.Clear

'Declaring varaibles

'Variable Declaration
Dim sAction,sInvokeOption,sName,sID,sObject,sColumnName,sValue,sButton
Dim objManageGlobalAlternates,sPerspective
Dim iCounter
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sInvokeOption= Parameter("sInvokeOption")
sName = Parameter("sName")
sID = Parameter("sID")
sObject = Parameter("sObject")
sColumnName = Parameter("sColumnName")
sValue = Parameter("sValue")
sButton = Parameter("sButton")

'Invoking [ Manage Global Alternates ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",oneIteration,"Select","ToolsManageGlobalAlternates"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating Object of [ Manage Global Alternates ] Dialog
Select Case Lcase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "myteamcenter",""
		Set objManageGlobalAlternates=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ManageGlobalAlternates","")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "structuremanager"
		Set objManageGlobalAlternates=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_ManageGlobalAlternates@2","")
End Select


	GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ManageGlobalAlternatesOperations"
	GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
	
'Checking existance of Manage Global Alternates dialog
If Fn_UI_Object_Operations("RAC_Common_ManageGlobalAlternatesOperations","Exist", objManageGlobalAlternates,GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Manage Global Alternates ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to add new Global Alternates
	Case "add"						
		If Fn_UI_JavaCheckBox_Operations("RAC_Common_ManageGlobalAlternatesOperations", "Set", objManageGlobalAlternates, "jchk_OpenByName", "ON")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to add item as global alternate as fail to click [ Open By Name ] option on [ Manage Global Alternates ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(1)
		
		If sName <> "" and sID <> "" Then
			LoadAndRunAction "RAC_Common\RAC_Common_OpenByNameOperations","RAC_Common_OpenByNameOperations",OneIteration,"findbyid&nameandopen","nooption",sName,sID,"","","",""
		ElseIf sName = "" and sID <> "" Then
			LoadAndRunAction "RAC_Common\RAC_Common_OpenByNameOperations","RAC_Common_OpenByNameOperations",OneIteration,"findbyidandopen","nooption","",sID,"","","",""
		End If
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ManageGlobalAlternatesOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added item [ " & Cstr(sID) & " ] as global alternate of currently selected item","","","","","")		
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "verifyexist"
		bFlag=False
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_ManageGlobalAlternatesOperations","GetRowCount",objManageGlobalAlternates.JavaTable("jtbl_Object"),"","","","","",""))-1	
			bFlag=False
			If Trim(objManageGlobalAlternates.JavaTable("jtbl_Object").GetCellData(iCounter,"Object"))=trim(sObject) Then				
				If sColumnName<>"" Then
					bFlag=False
					If trim(sValue)=trim(objManageGlobalAlternates.JavaTable("jtbl_Object").GetCellData(iCounter,sColumnName)) Then
						bFlag=True
					End If
				Else
					bFlag=True
				End If
				Exit For
			End If
		Next
		If bFlag=True Then
			IF sValue<>"" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sValue) & " ] exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ] on [ Manage Global Alternates ] dialog","","","","DONOTSYNC","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sObject) & " ] exist under column [ Object ] on [ Manage Global Alternates ] dialog","","","","DONOTSYNC","") 
			End If	
		Else
			IF sValue<>"" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sValue) & " ] does not exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ] on [ Manage Global Alternates ] dialog","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sObject) & " ] does not exist under column [ Object ] on [ Manage Global Alternates ] dialog","","","","","") 
			End IF	
			Call Fn_ExitTest()
		End If
End Select

If sButton<>"" Then
	objManageGlobalAlternates.JavaButton("jbtn_" & sButton).WaitProperty "enabled", 1, 20000
	If Fn_UI_JavaButton_Operations("RAC_Common_ManageGlobalAlternatesOperations", "Click", objManageGlobalAlternates, "jbtn_" & sButton)=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to complete operation [ " & Cstr(sAction) & " ] on [ Manage Global Alternates ] dialog as fail to click on [ " & Cstr(sButton) & " ]","","","","","")
		Call Fn_ExitTest()
	End If
	Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
End If

'Relasing Object of [ Manage Global Alternates ] Dialog
Set objManageGlobalAlternates=Nothing

Public Function Fn_ExitTest()
	'Relasing Object of [ Manage Global Alternates ] Dialog
	Set objManageGlobalAlternates=Nothing
	ExitTest
End Function


