'! @Name 			RAC_Common_NamedReferencesOperations
'! @Details 		Action word to perform operations on on Named References dialog
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam2 	sInvokeOption		: Named References dialog invoke option
'! @InputParam3 	sName				: Name of reference
'! @InputParam4 	sReference 			: Reference Name
'! @InputParam5 	sFilePath 			: File path to upload\download
'! @InputParam6 	sColumnName			: Table column name
'! @InputParam7 	sValue				: Table column value
'! @InputParam8 	sFileType			: File type
'! @InputParam9 	sTool				: Tool name
'! @InputParam10 	sButton				: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			02 Jan 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_NamedReferencesOperations","RAC_Common_NamedReferencesOperations",OneIteration,"Upload","menu","aaa.txt","","C:\VSEM_AUTOMATION\TestData\TextImport.txt","","","","","Close"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_NamedReferencesOperations","RAC_Common_NamedReferencesOperations",OneIteration,"VerifyReferenceSizeGreaterThanZero","menu","","BMP","","","","","","Close"

Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sInvokeOption,sName,sReference,sFilePath,sColumnName,sValue,sFileType,sTool,sButton	
Dim objNamedReferences,objUploadFile, objDownloadFile
Dim iCounter,iRanNo
Dim bFlag
Dim sTempValue,sPerspective

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sName= Parameter("sName")
sReference = Parameter("sReference")
sFilePath = Parameter("sFilePath")
sColumnName = Parameter("sColumnName")
sValue = Parameter("sValue")
sFileType = Parameter("sFileType")
sTool = Parameter("sTool")
sButton = Parameter("sButton")

'Invoking [ New Form ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","ViewNamedReferences"
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating object of [ Named References ] dialog
Select Case Lcase(sPerspective)
	Case "myteamcenter"
		'Creating object of [ Named References ]
		Set objNamedReferences = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR", "jdlg_NamedReferences","")
		Set objUploadFile = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR", "jwnd_TcDefaultApplet_jdlg_UploadFile","")
		Set objDownloadFile = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR", "jwnd_TcDefaultApplet_jdlg_DownloadFile","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_NamedReferencesOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of [ Named References ] dialog
If Fn_UI_Object_Operations("RAC_Common_NamedReferencesOperations","Exist",objNamedReferences,"","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Named References ] as [ Named References ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Upload Named Reference
	Case "Upload"
		'Clicking on [ Upload... ] button
		If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objNamedReferences,"jbtn_Upload") Then
			'Checking Existance of [ Upload File ] dialog
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If sFilePath<>"" Then
				'Set File Path from where to Upload 
				If Fn_UI_JavaEdit_Operations("RAC_Common_NamedReferencesOperations", "Set", objUploadFile, "jedt_FileName",sFilePath)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to upload file as fail to set file path [ " & Cstr(sFilePath) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			End If					
			
			'Click on Upload button
			If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objUploadFile,"jbtn_Upload")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to upload file as failed to click on [ Upload ] button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

			If sButton<>"" Then
				'Click on Upload button
				If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objNamedReferences,"jbtn_" & sButton)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to upload file as failed to click on [ " & Cstr(sButton) & " ] button","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)					
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully uploaded file from path [ " & Cstr(sFilePath) & " ]","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to upload file as failed to click on [ Upload ] button","","","","","")
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Download Named Reference
	Case "Download"
		GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow
		If sFilePath = ""  Then
			iRanNo = Fn_CommonUtil_GenerateRandomNumber(5)
			sFilePath = Fn_Setup_GetAutomationFolderPath("TestData") & "\PDFDataset" & iRanNo & ".pdf"
		End If
		'Clicking on [ Download... ] button
		If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objNamedReferences,"jbtn_Download") Then
			'Set [ Download File ] name
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			
			If sFilePath<>"" Then
				'Set File Path where to Download 
				If Fn_UI_JavaEdit_Operations("RAC_Common_NamedReferencesOperations", "Set", objDownloadFile, "jedt_FileName",sFilePath)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to download file as fail to set file path [ " & Cstr(sFilePath) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			End If					
			
			'Click on Download button
			If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objDownloadFile,"jbtn_Download")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Download file as failed to click on [ Download ] button","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

			If sButton<>"" Then
				'Click on Download button
				If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objDownloadFile,"jbtn_" & sButton)=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Download file as failed to click on [ " & Cstr(sButton) & " ] button","","","","","")
					Call Fn_ExitTest()
				End If
				
				'Add sFilePath to DataTable
				Call Fn_CommonUtil_DataTableOperations("AddColumn","NamedReferencesFilePath","","")
				Call Fn_CommonUtil_DataTableOperations("AddColumn","NamedReferencesFileCount","","")
				DataTable.SetCurrentRow 1
				iCountRowNumber = DataTable.Value("NamedReferencesFileCount","Global")
				If iCountRowNumber = "" Then
					iCountRowNumber = 1
					DataTable.Value("NamedReferencesFileCount","Global") = 1
				Else
					iCountRowNumber = Cint(iCountRowNumber) + 1
					DataTable.Value("NamedReferencesFileCount","Global") = iCountRowNumber
				End If
				DataTable.SetCurrentRow iCountRowNumber
				DataTable.Value("NamedReferencesFilePath","Global") = sFilePath
				DataTAble.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
				
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)					
			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully downloaded file from path [ " & Cstr(sFilePath) & " ]","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to download file as failed to click on [ Download ] button","","","","","")
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "Verify","VerifyCellData"
		bFlag=False
		For iCounter=0 to Cint(objNamedReferences.JavaTable("jtbl_References").GetROProperty("rows"))-1
			If Trim(Fn_UI_JavaTable_Operations("RAC_Common_NamedReferencesOperations","GetCellData",objNamedReferences,"jtbl_References",iCounter,"Name","","",""))=Trim(sName) Then
				If sAction="VerifyCellData" Then
					If Trim(Fn_UI_JavaTable_Operations("RAC_Common_NamedReferencesOperations","GetCellData",objNamedReferences,"jtbl_References",iCounter,sColumnName,"","",""))=Trim(sValue) Then
						bFlag=True
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified column [ " & Cstr(sColumnName) & " ] contain value [ " & Cstr(sValue) & " ] against file name [ " & Cstr(sName) & " ]","","","","","")
					End If
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified file with name [ " & Cstr(sName) & " ] exist in named references table","","","","","")
					bFlag=True
				End If
				Exit For
			End If
		Next
		If bFlag=False Then
			If sAction="VerifyCellData" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as column [ " & Cstr(sColumnName) & " ] does not contain value [ " & Cstr(sValue) & " ] against file name [ " & Cstr(sName) & " ]","","","","","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as file with name [ " & Cstr(sName) & " ] does not exist in named references table","","","","","")
			End If	
			Call Fn_ExitTest()
		End If
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objNamedReferences, "jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to complete operation [ " & Cstr(sAction) & " ] as fail to click on button [ " & Cstr(sButton) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Delete Named Reference
	Case "Delete"
		bFlag=False
		For iCounter=0 to Cint(Fn_UI_Object_GetROProperty("RAC_Common_NamedReferencesOperations",objNamedReferences.JavaTable("jtbl_References"),"rows"))-1
			If Trim(Fn_UI_JavaTable_Operations("RAC_Common_NamedReferencesOperations","GetCellData",objNamedReferences,"jtbl_References",iCounter,"Name","","",""))=Trim(sName) Then
				Call Fn_UI_JavaTable_Operations("RAC_Common_NamedReferencesOperations","SelectRow",objNamedReferences,"jtbl_References",iCounter,"","","","")
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				bFlag=True
				Exit For
			End If
		Next
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to delete file [ " & Cstr(sName) & " ] as file with name does not found in named references table","","","","","")
			Call Fn_ExitTest()
		End If
		'Clicking on [ Delete ] button
		If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objNamedReferences, "jbtn_Delete") Then
			If Fn_UI_Object_Operations("RAC_Common_NamedReferencesOperations", "Exist", JavaDialog("jdlg_Delete"),"") Then
				If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", JavaDialog("jdlg_Delete"), "jbtn_Yes")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to delete file [ " & Cstr(sName) & " ] as failed to click ok [ Yes ] button of [ Delete ] confirmation dialog","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				
				If sButton<>"" Then
					If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objNamedReferences, "jbtn_" & sButton)=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to delete file [ " & Cstr(sName) & " ] as failed to click ok [ " & Cstr(sButton) & " ] button","","","","","")
						Call Fn_ExitTest()
					End If
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to delete file [ " & Cstr(sName) & " ] as [ Delete ] confirmation dialog does not exist","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to delete file [ " & Cstr(sName) & " ] as failed to click on [ Delete ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully deleted file [ " & Cstr(sName) & " ] from named references table","","","","","")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "VerifyReferenceSizeGreaterThanZero"
		bFlag=False
		For iCounter=0 to Cint(objNamedReferences.JavaTable("jtbl_References").GetROProperty("rows"))-1
			If Trim(Fn_UI_JavaTable_Operations("RAC_Common_NamedReferencesOperations","GetCellData",objNamedReferences,"jtbl_References",iCounter,"Reference","","",""))=Trim(sReference) Then
				sTempValue=Trim(Fn_UI_JavaTable_Operations("RAC_Common_NamedReferencesOperations","GetCellData",objNamedReferences,"jtbl_References",iCounter,"Size","","",""))
				If sTempValue<>"" Then
					sTempValue=Replace(Lcase(sTempValue),"kb","")
					If Cint(sTempValue) >0 Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified Reference [ " & Cstr(sReference) & " ] contain Size grater than [ 0 Kb ], actual size is [ " & Cstr(sTempValue) & " Kb ]","","","","","")
						bFlag=True
						Exit For
					End If
				End If				
			End If
		Next
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as Reference [ " & Cstr(sReference) & " ] is contain [ 0 Kb ] Size","","","","","")
			Call Fn_ExitTest()
		End If
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_NamedReferencesOperations", "Click", objNamedReferences, "jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to complete operation [ " & Cstr(sAction) & " ] as fail to click on button [ " & Cstr(sButton) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
End Select

'Releasing objects
Set objNamedReferences=Nothing
Set objUploadFile =Nothing

Function Fn_ExitTest()
	'Releasing objects
	Set objNamedReferences=Nothing
	Set objUploadFile =Nothing
	ExitTest
End Function

