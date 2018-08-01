'! @Name 			RAC_PSE_ObjectAddOperations
'! @Details 		Action word to perform operations on Add dialog.
'! @InputParam1 	sAction			: Action Name
'! @InputParam2		sInvokeOption	: Add dialog invoke option
'! @InputParam3		sItemID			: Item Id to add
'! @InputParam4		sRevision		: Item revision
'! @InputParam5		sItemName		: Item Name
'! @InputParam6		sName			: Name
'! @InputParam7		sViewType		: View Type
'! @InputParam8		sButton			: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			14 Jul 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_ObjectAddOperations","RAC_PSE_ObjectAddOperations",OneIteration,"Add","nooption","123456", "","","","",""

Option Explicit
Err.Clear

'Declaring varaibles

'Variable Declaration
Dim sAction,sInvokeOption,sItemID,sRevision,sItemName,sName,sViewType,sButton
Dim objAdd

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sInvokeOption= Parameter("sInvokeOption")
sItemID = Parameter("sItemID")
sRevision = Parameter("sRevision")
sItemName = Parameter("sItemName")
sName = Parameter("sName")
sViewType = Parameter("sViewType")
sButton = Parameter("sButton")

'Invoking [ Add ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

'Getting object of Add dialog
Set objAdd=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_Add","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ObjectAddOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of Add dialog
If Fn_UI_Object_Operations("RAC_PSE_ObjectAddOperations","Exist", objAdd,GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Add ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Case to find object by ID and open
	Case "add"
		'Find object by ID	
		Call Fn_UI_Object_Operations("RAC_PSE_ObjectAddOperations","settoproperty",objAdd.JavaStaticText("jstx_AddLabel"),"","label","Item ID:")
		If Fn_UI_JavaEdit_Operations("RAC_PSE_ObjectAddOperations","Set",objAdd,"jedt_AddEdit",sItemID) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to find & add object as failed to set [ Item ID ] value on [ Add ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		Call Fn_UI_JavaEdit_Operations("RAC_PSE_ObjectAddOperations","Activate",objAdd,"jedt_AddEdit","")
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on [ OK ] button
		If Fn_UI_JavaButton_Operations("RAC_PSE_ObjectAddOperations", "Click", objAdd,"jbtn_OK")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Add object as failed to click on [ OK ] button on [ Add ] dialog","","","","","")
			Call Fn_ExitTest()
		End If			
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to add object of ID [ " & Cstr(sItemID) & " ] on selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully added object of ID [ " & Cstr(sItemID) & " ] on selected object in assembly","","","","","")
		End If		
End Select

'Relasing Object of [ Add ] Dialog
Set objAdd=Nothing

Public Function Fn_ExitTest()
	'Relasing Object of [ Add ] Dialog
	Set objAdd=Nothing
	ExitTest
End Function
