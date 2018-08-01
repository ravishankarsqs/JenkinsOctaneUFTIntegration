'! @Name RAC_Common_GenerateReport
'! @Details This actionword is used to perform menu operations in Teamcenter application.
'! @InputParam1. sAction = Action to be performed
'! @InputParam2. sInvokeOption = Invoke option to open the generate report dialog
'! @InputParam3. sNodePath = Nav tree node path for which report is to be generated
'! @InputParam4. sButtonName = Button to be clicked at the end of operation
'! @Author Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer Kundan Kudale kundan.kudale@sqs.com
'! @Date 25 Mar 2016
'! @Version 1.0
'! @Example  dictGenerateReportInformation("ReportType") = "Item Reports"
'! @Example  dictGenerateReportInformation("ReportName") = "DA Print Report"
'! @Example  dictGenerateReportInformation("ReportDisplayLocale") = "English"
'! @Example  dictGenerateReportInformation("ReportStylesheets") = "DA_Report_Excel.xsl"
'! @Example  dictGenerateReportInformation("ReportDatasetName") = ""
'! @Example  dictGenerateReportInformation("ReportDatasetCheckboxValue") = "ON"
'! @Example  LoadAndRunAction "RAC_Common\RAC_Common_GenerateReport","RAC_Common_GenerateReport",OneIteration,"AutoGenerateReport","PopupMenuSelect","Home~AutomatedTest~TestCaseFolder~ItemNode~ItemRevisionNode",""

Option Explicit

Dim sAction, sInvokeOption, sNodePath, sButtonName
Dim objGenerateReport
Dim sPerspective

'Get action parameters in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sNodePath = Parameter("sNodePath")
sButtonName = Parameter("sButtonName")

Select Case Lcase(sInvokeOption)
	Case "popupmenuselect"
		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "PopupMenuSelect",sNodePath ,"GenerateReport"
	Case "bomtablepopupmenu"
		LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "PopupMenuSelect",sNodePath,"","","GenerateReport"
		
End Select

sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating object of [ New Item ] dialog
Select Case Lcase(sPerspective)
	Case "myteamcenter","","my teamcenter"
		'Creating object of [ Generate Report Wizard ]
		Set objGenerateReport = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_GenerateReport","")
	Case "structuremanager","structure manager","StructureManager"
		'Creating object of [ Generate Report Wizard ]
		Set objGenerateReport = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_GenerateReport@2","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_GenerateReport"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Verify existence of generate report dialog
'Set objGenerateReport = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_GenerateReport","")
If Not Fn_UI_Object_Operations("RAC_Common_GenerateReport","Exist", objGenerateReport, GBL_DEFAULT_TIMEOUT,"","") Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ " & Cstr(sNodePath) & " ] node as [ Generate Reprot ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

Select Case sAction

	Case "autogeneratereportandverifyexcelcontents"
	
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		'Select report name
		If dictGenerateReportInformation("ReportName") <> "" Then
			If Fn_UI_JavaTable_Operations("RAC_Common_GenerateReport","selectrowext",objGenerateReport,"jtbl_SelectReport","","",dictGenerateReportInformation("ReportName") ,"LEFT","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as failed to select [" & dictGenerateReportInformation("ReportName") & "] in reports table","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		
		objGenerateReport.JavaButton("jbtn_Next").WaitProperty "enabled",1,5000		
'		'Click on Next button
'		If Fn_UI_Object_Operations("RAC_Common_GenerateReport", "enabled", objGenerateReport.JavaButton("jbtn_Next"),"", "", "") Then
		If Cint(objGenerateReport.JavaButton("jbtn_Next").GetROProperty("enabled"))=1 Or objGenerateReport.JavaButton("jbtn_Next").GetROProperty("enabled")="1" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_GenerateReport", "Click", objGenerateReport, "jbtn_Next") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as failed to click on next button ","","","","","")
				Call Fn_ExitTest()
			End If
		End If
'		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		
		objGenerateReport.JavaButton("jbtn_Finish").WaitProperty "enabled",1,20000
		'Click on Finish button
		If Fn_UI_JavaButton_Operations("RAC_Common_GenerateReport", "Click", objGenerateReport, "jbtn_Finish") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as failed to click on Finish button ","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		
		If Window("wwnd_Excel").Dialog("dlg_Excel").Exist(30) Then
			Window("wwnd_Excel").Dialog("dlg_Excel").WinButton("wbtn_Yes").Click
		End If
		
'		If Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition","",0,"",dictGenerateReportInformation("ItemIDInExcel"),True) = False Then
'			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictGenerateReportInformation("ItemIDInExcel")) & " ] was not found in excel file","","","","","")
'			Call Fn_ExitTest()
'		Else
'			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as [ " & Cstr(dictGenerateReportInformation("ItemIDInExcel")) & " ] was found in excel file","","","","","")
'		End If
'
	Case "basicautogeneratereport"	
		'Select report name
		If dictGenerateReportInformation("ReportName") <> "" Then
			If Fn_UI_JavaTable_Operations("RAC_Common_GenerateReport","selectrowext",objGenerateReport,"jtbl_SelectReport","","",dictGenerateReportInformation("ReportName") ,"LEFT","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as failed to select [" & dictGenerateReportInformation("ReportName") & "] in reports table","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		
		If dictGenerateReportInformation("ReportName")<>"BOM - Impacted Changes Report" Then		
			'Click on Next button
			If Fn_UI_JavaButton_Operations("RAC_Common_GenerateReport", "Click", objGenerateReport, "jbtn_Next") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as failed to click on next button ","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		End If
		
		'Click on Finish button
		objGenerateReport.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_GenerateReport", "Click", objGenerateReport, "jbtn_Finish") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as failed to click on Finish button ","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		
		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_GenerateReport", "Click", objGenerateReport, "jbtn_Close") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as failed to click on Close button ","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
End Select

