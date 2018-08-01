'! @Name 			RAC_PSE_ListSubstitutesOperations
'! @Details 		Action word to perform operations on ListSubstitutes dialog
'! @InputParam1 	sAction			: Action Name
'! @InputParam2		sInvokeOption	: Add dialog invoke option
'! @InputParam3		sItemID			: Item Id to add
'! @InputParam4		sRevision		: Item revision
'! @InputParam5		sItemName		: Item Name
'! @InputParam6		sName			: Name
'! @InputParam7		sViewType		: View Type
'! @InputParam8		sBOMLine		: BOM Line node
'! @InputParam8		sButton			: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			14 Jul 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_ListSubstitutesOperations","RAC_PSE_ListSubstitutesOperations",OneIteration,"Add","nooption","123456", "","","","","",""

Option Explicit
Err.Clear

'Declaring varaibles

'Variable Declaration
Dim sAction,sInvokeOption,sItemID,sRevision,sItemName,sName,sViewType,sBOMLine,sButton
Dim objListSubstitutes

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
sBOMLine = Parameter("sBOMLine")
sButton = Parameter("sButton")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ListSubstitutesOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Invoking [ ListSubstitutes ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "toolbar"
		'Case to invoke List Substitutes panel from bottom toolbar
		JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaCheckBox("jckb_AddSubstitute").Set "ON"
		JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaStaticText("jstx_PanelHeader").SetTOProperty "label","List Substitutes"
		If Fn_UI_Object_Operations("RAC_PSE_ListSubstitutesOperations","Exist", JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaStaticText("jstx_PanelHeader"),GBL_DEFAULT_TIMEOUT,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ ListSubstitutes ] dialog as dialog does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaStaticText("jstx_PanelHeader").DblClick 5,5
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End Select

'Getting object of ListSubstitutes dialog
Set objListSubstitutes=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_ListSubstitutes","")
			
'Checking existance of ListSubstitutes dialog
If Fn_UI_Object_Operations("RAC_PSE_ListSubstitutesOperations","Exist", objListSubstitutes,GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ ListSubstitutes ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "add"	
		'Click on [ Add ] button
		If Fn_UI_JavaButton_Operations("RAC_PSE_ListSubstitutesOperations", "Click", objListSubstitutes,"jbtn_Add")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to ListSubstitutes object as failed to click on [ Add ] button on [ ListSubstitutes ] dialog","","","","","")
			Call Fn_ExitTest()
		End If			
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		LoadAndRunAction "RAC_StructureManager\RAC_PSE_ObjectAddOperations","RAC_PSE_ObjectAddOperations",OneIteration,"Add","nooption",sItemID,"","","","",""
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ListSubstitutesOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Add list substitute object of ID [ " & Cstr(sItemID) & " ] on selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Added list substitute object of ID [ " & Cstr(sItemID) & " ] on selected object in assembly","","","","","")
		End If		
End Select

'Relasing Object of [ ListSubstitutes ] Dialog
Set objListSubstitutes=Nothing

Public Function Fn_ExitTest()
	'Relasing Object of [ ListSubstitutes ] Dialog
	Set objListSubstitutes=Nothing
	ExitTest
End Function

