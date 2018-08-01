'! @Name 			RAC_Search_QuickSearchOperations
'! @Details 		This action word is used to perform quick search operation
'! @InputParam1 	sSrchType 	: Quick search type
'! @InputParam2 	sSrchText 	: Search text
'! @InputParam3 	sObject 	: Object name
'! @InputParam4 	sColumnName	: column name
'! @InputParam5 	sValue 		: value
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			01 Jul 2015
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Search\RAC_Search_QuickSearchOperations","RAC_Search_QuickSearchOperations",OneIteration,"Item ID","001236","","",""

Option Explicit
Err.Clear

'Declaring variables
Dim sSearchType, sSearchText,sObject,sColumnName,sValue
Dim objDefaultWindow,objNoAccessibleObjects,objQuickOpenResults,objQuickOpenResults2
Dim iCounter,iRows
Dim bFlag,bReturn
Dim sCellData, sTempValue
Dim aCellData

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

sSearchType = Parameter("sSearchType")
sSearchText = Parameter("sSearchText")
sObject = Parameter("sObject")
sColumnName = Parameter("sColumnName")
sValue = Parameter("sValue")

'Creating object of Teamcenter Default Window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","jwnd_SearchDefaultWindow","")
'Creating object quick open results dialog
Set objQuickOpenResults=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","jdlg_QuickOpenResults","")
Set objQuickOpenResults2=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","jdlg_QuickOpenResults@2","")
'Creating object of No Accessible Object Window
Set objNoAccessibleObjects=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","jdlg_NoAccessibleObjects","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Search_QuickSearchOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Capture business functionality start time	
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Quick Search","","Search Type",sSearchType)

'selecting quick search type
sTempValue = ""
If sSearchType = "Item ID_Ext" Then
	sSearchType = "Item ID"
	sTempValue = "Item ID_Ext"
End If
If Fn_UI_JavaToolbar_Operations("RAC_Search_QuickSearchOperations", "DropdownMenuSelect", objDefaultWindow, "jtlbr_QuickSearchToolbar","Perform Search", "", sSearchType, "")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform quick search operation as failed to select quick search type [ " & Cstr(sSearchType) & " ]","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

If sTempValue <> "" Then
	sSearchType = "Item ID_Ext"
End If

'Enter quick search criteria
If Fn_UI_JavaEdit_Operations("RAC_Search_QuickSearchOperations","Set",objDefaultWindow,"jedt_QuickSearch",sSearchText)=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set value [ " & Cstr(sSearchText) & " ] in quick search edit field","","","","","")
	Call Fn_ExitTest()
End If			
'Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

'Click [ Perform Search ] button
If Fn_UI_JavaToolbar_Operations("RAC_Search_QuickSearchOperations", "Click", objDefaultWindow,"jtlbr_QuickSearchToolbar","Perform Search", "", "", "")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform quick search operation as fail to click on [ Perform Search ] button","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

'Checking Existance of [ No Accessible Objects ] dialog
If Fn_UI_Object_Operations("RAC_Search_QuickSearchOperations","Exist",objNoAccessibleObjects,GBL_MIN_TIMEOUT,"","") Then
	objNoAccessibleObjects.JavaButton("jbtn_OK").Click
	If sValue="VerifyNoResultFound" Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified quick search results not found for search type [ " & Cstr(sSearchType) & " ] & search value [ " & Cstr(sSearchText) & " ]","","","","","")
		ExitAction
	ELse
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform quick search operation for type [ " & Cstr(sSearchType) & " ] and value [ " & Cstr(sSearchText) & " ]","","","","","")
		Call Fn_ExitTest()
	End If
Else
	If sValue="VerifyNoResultFound" Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as search results found for search type [ " & Cstr(sSearchType) & " ] & search value [ " & Cstr(sSearchText) & " ]","","","","","")
		Call Fn_ExitTest()
	End If
End If

'Creating Object [ Quick Open Results ] dialog
If Fn_UI_Object_Operations("RAC_Search_QuickSearchOperations","Exist",objQuickOpenResults,GBL_MIN_TIMEOUT,"","") Then
	'do nothing
ElseIf Fn_UI_Object_Operations("RAC_Search_QuickSearchOperations","Exist",objQuickOpenResults2,GBL_MIN_TIMEOUT,"","") Then
	Set  objQuickOpenResults = objQuickOpenResults2
Else 
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Quick Search","","Search Type",sSearchType)
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform quick search for type [ " & Cstr(sSearchType) & " ] and value [ " & Cstr(sSearchText) & " ]","","","","","")
	ExitAction
End if

'Checking existance of [ Load All ] button to load all search results
If Fn_UI_Object_Operations("RAC_Search_QuickSearchOperations","Exist",objQuickOpenResults.JavaButton("jbtn_LoadAll"),"","","") then
	If Cint(Fn_UI_Object_Operations("RAC_Search_QuickSearchOperations","GetROProperty",objQuickOpenResults.JavaButton("jbtn_LoadAll"),"","enabled","")) = 1 Then
		If Fn_UI_JavaButton_Operations("RAC_Search_QuickSearchOperations","Click",objQuickOpenResults, "jbtn_LoadAll")=False Then
	   		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform quick search operation as fail to click on [ Load All ] button of [ Open Quick Search ] dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	End If
End If

'Set object name
If sObject="" Then
	sObject=sSearchText	
End If

'Get Row count
iRows = Fn_UI_Object_Operations("RAC_Search_QuickSearchOperations","GetROProperty",objQuickOpenResults.JavaTable("jtbl_QuickSearchResultsTable"),"","rows","")
For iCounter = 0 to iRows - 1
	'Get cell data
	sCellData =  Fn_UI_JavaTable_Operations("RAC_Search_QuickSearchOperations","GetCellData",objQuickOpenResults,"jtbl_QuickSearchResultsTable",iCounter,0,"","","")
	
	Select Case sSearchType
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'This will work when name will be unique
		Case "Item Name"
			If instr(1,CStr(sCellData),CStr(sObject)) Then
				If sColumnName<>"" Then
					If trim(Fn_UI_JavaTable_Operations("RAC_Search_QuickSearchOperations","GetCellData",objQuickOpenResults,"jtbl_QuickSearchResultsTable",iCounter,sColumnName,"","",""))=Trim(sValue) Then
						Call Fn_UI_JavaTable_Operations("RAC_Search_QuickSearchOperations","SelectRow",objQuickOpenResults,"jtbl_QuickSearchResultsTable",iCounter,"","","","")
						bReturn =Fn_UI_JavaButton_Operations("RAC_Search_QuickSearchOperations","Click",objQuickOpenResults, "jbtn_Open")
						Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
						Exit For
					End If
				Else
					Call Fn_UI_JavaTable_Operations("RAC_Search_QuickSearchOperations","SelectRow",objQuickOpenResults,"jtbl_QuickSearchResultsTable",iCounter,"","","","")
					bReturn =Fn_UI_JavaButton_Operations("RAC_Search_QuickSearchOperations","Click",objQuickOpenResults, "jbtn_Open")
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Exit For
				 End If					 
			Else 
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform quick search operation for type [ " & Cstr(sSearchType) & " ] and value [ " & Cstr(sSearchText) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Item ID"
			bFlag = False
			aCellData = Split(sCellData ,"-",-1)
			If isNumeric(aCellData(0)) Then
				If cdbl(sObject) = Cdbl(aCellData(0)) Then bFlag = True
			Else
				If (sObject) = (aCellData(0)) Then bFlag = True
			End If
			If bFlag Then
				Call Fn_UI_JavaTable_Operations("RAC_Search_QuickSearchOperations","SelectRow",objQuickOpenResults,"jtbl_QuickSearchResultsTable",iCounter,"","","","")
				bReturn = Fn_UI_JavaButton_Operations("RAC_Search_QuickSearchOperations","Click",objQuickOpenResults, "jbtn_Open")
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				Exit For
			Else 
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform quick search operation for type [ " & Cstr(sSearchType) & " ] and value [ " & Cstr(sSearchText) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'This will work when name will be unique
		Case "Item ID_Ext"
			If instr(1,CStr(sCellData),CStr(sObject)) Then
				If sColumnName<>"" Then
					If Trim(Fn_UI_JavaTable_Operations("RAC_Search_QuickSearchOperations","GetCellData",objQuickOpenResults,"jtbl_QuickSearchResultsTable",iCounter,sColumnName,"","",""))=Trim(sValue) Then
						Call Fn_UI_JavaTable_Operations("RAC_Search_QuickSearchOperations","SelectRow",objQuickOpenResults,"jtbl_QuickSearchResultsTable",iCounter,"","","","")
						bReturn =Fn_UI_JavaButton_Operations("RAC_Search_QuickSearchOperations","Click",objQuickOpenResults, "jbtn_Open")
						Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
						Exit For
					End If
				Else
					Call Fn_UI_JavaTable_Operations("RAC_Search_QuickSearchOperations","SelectRow",objQuickOpenResults,"jtbl_QuickSearchResultsTable",iCounter,"","","","")
					bReturn =Fn_UI_JavaButton_Operations("RAC_Search_QuickSearchOperations","Click",objQuickOpenResults, "jbtn_Open")
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					Exit For
				 End If					 
			Else 
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform quick search operation for type [ " & Cstr(sSearchType) & " ] and value [ " & Cstr(sSearchText) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Dataset Name"
			If CStr(sObject) = CStr(sCellData) Then
				Call Fn_UI_JavaTable_Operations("RAC_Search_QuickSearchOperations","SelectRow",objQuickOpenResults,"jtbl_QuickSearchResultsTable",iCounter,"","","","")
				bReturn =Fn_UI_JavaButton_Operations("RAC_Search_QuickSearchOperations","Click",objQuickOpenResults, "jbtn_Open")
				Exit For
			Else 
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform quick search operation for type [ " & Cstr(sSearchType) & " ] and value [ " & Cstr(sSearchText) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		'- - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - -- - - - - - - - - - - - - 
	End Select
Next

'Releasing all objects
Set objDefaultWindow= Nothing
Set objNoAccessibleObjects= Nothing
Set objQuickOpenResults= Nothing
Set objQuickOpenResults2= Nothing

If bReturn = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform quick search operation for type [ " & Cstr(sSearchType) & " ] and value [ " & Cstr(sSearchText) & " ]","","","","","")
	Call Fn_ExitTest()
Else
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Quick Search","","Search Type",sSearchType)
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully perform quick search for type [ " & Cstr(sSearchType) & " ] and value [ " & Cstr(sSearchText) & " ]","","","","","")
End If

Function Fn_ExitTest()
	'Releasing all objects
	Set objDefaultWindow= Nothing
	Set objNoAccessibleObjects= Nothing
	Set objQuickOpenResults= Nothing
	Set objQuickOpenResults2= Nothing
	ExitTest
End Function

