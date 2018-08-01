'! @Name 			RAC_Common_OpenByNameOperations
'! @Details 		Action word to perform operations on Open By Name dialog.
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
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_OpenByNameOperations","RAC_Common_OpenByNameOperations",OneIteration,"findbyidandopen","nooption", "", "123456","","","",""

Option Explicit
Err.Clear

'Declaring varaibles

'Variable Declaration
Dim sAction,sInvokeOption,sName,sID,sObject,sColumnName,sValue,sButton
Dim objOpenByName, sPerspective
Dim iCount
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

'Invoking [ Open By Name ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating Object of [ Open By Name ] Dialog
Select Case Lcase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "myteamcenter",""
		Set objOpenByName=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_OpenByName","")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "structuremanager"
		Set objOpenByName=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_OpenByName@2","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_OpenByNameOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
			
'Checking existance of Open By Name dialog
If Fn_UI_Object_Operations("RAC_Common_OpenByNameOperations","Exist", objOpenByName,GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Open By Name ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to find object by ID and open
	Case "findbyidandopen","findbyid&nameandopen"
		If sColumnName="" Then
			sColumnName="Object"
		End If
		
		If Lcase(sAction)="findbyid&nameandopen" Then		
			'Setting Object name
			If Fn_UI_JavaEdit_Operations("RAC_Common_OpenByNameOperations","Set",objOpenByName,"jedt_Name",sName) = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to find object as failed to set [ Name ] value on [ Open By Name ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		If Fn_UI_JavaEdit_Operations("RAC_Common_OpenByNameOperations","Set",objOpenByName,"jedt_ID",sID) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to find object as failed to set [ ID ] value on [ Open By Name ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Clicking On Find Button to create Dataset
		If Fn_UI_JavaButton_Operations("RAC_Common_OpenByNameOperations", "Click",objOpenByName,"jbtn_Find")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to find object as failed to click on [ Find ] button on [ Open By Name ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(5)
		
		For iCount = 0 to objOpenByName.JavaTable("jtbl_Object").GetROProperty("rows") -1
			bFlag = False
			If cStr(sID) <> "" AND sName <> "" Then
				If instr(objOpenByName.JavaTable("jtbl_Object").GetCellData(iCount,sColumnName), sID & "-" & sName) > 0 Then bFlag = True
			ElseIf sName = "" AND cStr(sID) <> "" Then
				If instr(objOpenByName.JavaTable("jtbl_Object").GetCellData(iCount,sColumnName), cStr(sID)) > 0 Then bFlag = True
			End If
			If bFlag = True Then
				objOpenByName.JavaTable("jtbl_Object").DoubleClickCell iCount,sColumnName,"LEFT"
				Call Fn_RAC_ReadyStatusSync(5)
				Exit for
			End If
		Next
		If bFlag = True Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully open object based on search criteria : ID = [ " & Cstr(sID) & " ] from [ Open By Name ] dialog","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to find and open object from [ Open By Name ] dialog as there are no objects found based on search criteria : ID = [ " & Cstr(sID) & " ]","","","","","")
			Call Fn_ExitTest()
		End IF	
End Select

'Relasing Object of [ Open By Name ] Dialog
Set objOpenByName=Nothing

Public Function Fn_ExitTest()
	'Relasing Object of [ Open By Name ] Dialog
	Set objOpenByName=Nothing
	ExitTest
End Function
