'! @Name 			RAC_PSE_FindInDisplayOperations
'! @Details 		Action word to perform operations on Find In Display dialog
'! @InputParam1 	sAction			: Action Name
'! @InputParam2		sInvokeOption	: Add dialog invoke option
'! @InputParam3		sCondition		: Search Condition
'! @InputParam4		sPropertyName	: Property Name
'! @InputParam5		sOperator		: Operator
'! @InputParam6		sSearchingValue	: Searching vlaue
'! @InputParam7		sButton			: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			14 Jul 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_StructureManager\RAC_PSE_FindInDisplayOperations","RAC_PSE_FindInDisplayOperations",OneIteration,"find","toolbar","","Find No.","=","10","Close"

Option Explicit
Err.Clear

'Declaring varaibles

'Variable Declaration
Dim sAction,sInvokeOption,sCondition,sPropertyName,sOperator,sSearchingValue,sButton
Dim aPropertyName,aOperator,aSearchingValue
Dim objFindInDisplay
Dim iCounter,iCount
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sInvokeOption= Parameter("sInvokeOption")
sCondition = Parameter("sCondition")
sPropertyName = Parameter("sPropertyName")
sOperator = Parameter("sOperator")
sSearchingValue = Parameter("sSearchingValue")
sButton = Parameter("sButton")

'Invoking [ FindInDisplay ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "toolbar"
		'Case to invoke Find In Display panel from bottom toolbar
		JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaCheckBox("jckb_FindWithProperties").Set "ON"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		If JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaStaticText("jstx_FindInDisplay").Exist(20) Then
			JavaWindow("jwnd_StructureManager").JavaApplet("japt_PSEApplet").JavaStaticText("jstx_FindInDisplay").DblClick 5,5,"LEFT"
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
End Select

'Getting object of FindInDisplay dialog
Set objFindInDisplay=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_FindInDisplay","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_FindInDisplayOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of FindInDisplay dialog
If Fn_UI_Object_Operations("RAC_PSE_FindInDisplayOperations","Exist", objFindInDisplay,GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ FindInDisplay ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case "find"	
		'Click on [ jbtn_Clear ] button
		If Fn_UI_JavaButton_Operations("RAC_PSE_FindInDisplayOperations", "Click", objFindInDisplay,"jbtn_Clear")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to FindInDisplay object as failed to click on [ Clear ] button on [ FindInDisplay ] dialog","","","","","")
			Call Fn_ExitTest()
		End If			
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Click on [ Add ] button
		If Fn_UI_JavaButton_Operations("RAC_PSE_FindInDisplayOperations", "Click", objFindInDisplay,"jbtn_Add")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to FindInDisplay object as failed to click on [ Add ] button on [ FindInDisplay ] dialog","","","","","")
			Call Fn_ExitTest()
		End If			
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		aPropertyName=Split(sPropertyName,"~")
		aOperator=Split(sOperator,"~")
		aSearchingValue=Split(sSearchingValue,"~")
		
		bFlag=False
		For iCounter = 0 to Ubound(aPropertyName)
			bFlag=False
			For iCount = 0 to objFindInDisplay.JavaTable("jtbl_Criteria").GetROProperty("rows")-1
				If aPropertyName(iCounter)<>"" Then
					objFindInDisplay.JavaTable("jtbl_Criteria").SetCellData iCount,"Property Name",aPropertyName(iCounter)
				End If
				If aOperator(iCounter)<>"" Then
					objFindInDisplay.JavaTable("jtbl_Criteria").SetCellData iCount,"Operator",aOperator(iCounter)
				End If
				If aSearchingValue(iCounter)<>"" Then
					objFindInDisplay.JavaTable("jtbl_Criteria").SetCellData iCount,"Searching Value",aSearchingValue(iCounter)
				End If				
				If Err.Number<>0 then
					bFlag=False
				Else
					bFlag=True
				End If	
			Next
			If bFlag=False Then
				Exit For
			End If
		Next
				
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Add search criteria on [ Find In Display ] dialog due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		'Click on [ Find ] button
		If Fn_UI_JavaButton_Operations("RAC_PSE_FindInDisplayOperations", "Click", objFindInDisplay,"jbtn_Find")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to FindInDisplay object as failed to click on [ Find ] button on [ FindInDisplay ] dialog","","","","","")
			Call Fn_ExitTest()
		End If			
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		If sButton="Close" Then
			objFindInDisplay.Close
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform Find In Display for criteria [ Property Name = " & Cstr(sPropertyName) & " ],[ Searching Value = " & Cstr(sSearchingValue) & " ],[ Operator = " & Cstr(sOperator) & " ] ","","","","","")
End Select

'Relasing Object of [ FindInDisplay ] Dialog
Set objFindInDisplay=Nothing

Public Function Fn_ExitTest()
	'Relasing Object of [ FindInDisplay ] Dialog
	Set objFindInDisplay=Nothing
	ExitTest
End Function
